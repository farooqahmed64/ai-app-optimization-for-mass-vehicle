"""FastAPI — données 100 %% Supabase (HTTPS) quand DATA_BACKEND=supabase."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Query, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import excel_export
from .config import settings
from .local_store import LocalSheetStore

app = FastAPI(title="WGHT Vehicle Mass API", version="0.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origin_list,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

local_store = LocalSheetStore(settings.local_data_path)
DEFAULT_PROJECT = "default"
logger = logging.getLogger(__name__)


def _sheet_key(sheet_id: str) -> str:
    s = sheet_id.lower()
    if s not in ("bd", "synthesis"):
        raise HTTPException(404, f"Unknown sheet '{sheet_id}'")
    return s


def _cloud_sheet_name(sheet_id: str) -> str:
    return "BD" if sheet_id == "bd" else "SYNTHESIS"


def _remote_store():
    if settings.use_supabase_rest:
        from . import supabase_rest

        return supabase_rest
    if settings.use_postgres_direct:
        from . import supabase_store

        return supabase_store
    if settings.use_databricks:
        from . import databricks_store

        return databricks_store
    return None


def _require_remote():
    store = _remote_store()
    if not store:
        raise HTTPException(
            503,
            "Supabase non configure : SUPABASE_URL + SUPABASE_SERVICE_KEY dans api/.env",
        )
    return store


@app.get("/api/v1/health")
def health():
    return {
        "ok": True,
        "backend": settings.api_mode,
        "remoteOnly": settings.remote_only,
    }


def _store_has_grid_cache() -> bool:
    store = _remote_store()
    return bool(store and hasattr(store, "fetch_grid_pack"))


@app.get("/api/v1/config")
def config():
    return {
        "mode": settings.api_mode,
        "chunkedLoad": settings.use_remote_store,
        "cloudPersist": settings.use_remote_store,
        "remoteOnly": settings.remote_only,
        # When true the frontend can fetch the precomputed display grid from the DB
        # (GET /sheets/{id}/grid) instead of the static /public/data/*-grid.json file.
        "gridInDb": _store_has_grid_cache(),
        # Realtime (websocket) — le navigateur s'abonne directement a Supabase
        # pour recevoir en direct les modifs Synthesis (ligne 16 Curb mass) et
        # les refleter sur Options SP2 (ligne 14). anon key = publique, sans risque.
        "supabaseUrl": settings.supabase_url if settings.use_supabase else "",
        "supabaseAnonKey": settings.supabase_anon_key if settings.use_supabase else "",
        "version": 2,
        "projectId": DEFAULT_PROJECT,
    }


@app.get("/api/v1/sheets/{sheet_id}/meta")
def sheet_meta(sheet_id: str, project_id: str = DEFAULT_PROJECT):
    sheet_id = _sheet_key(sheet_id)
    if settings.remote_only:
        store = _require_remote()
        meta = store.fetch_meta(project_id, _cloud_sheet_name(sheet_id))
        if not meta:
            raise HTTPException(
                404,
                f"No meta for {sheet_id} — run: python tools/ingest_supabase_rest.py",
            )
        return meta
    store = _remote_store()
    if store:
        try:
            meta = store.fetch_meta(project_id, _cloud_sheet_name(sheet_id))
            if meta:
                return meta
        except Exception as exc:
            raise HTTPException(503, str(exc)) from exc
    try:
        return local_store.get_meta(sheet_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc


@app.get("/api/v1/sheets/{sheet_id}/cells")
def sheet_cells(
    sheet_id: str,
    rowMin: int = Query(..., ge=1),
    rowMax: int = Query(..., ge=1),
    project_id: str = DEFAULT_PROJECT,
):
    if rowMax < rowMin:
        raise HTTPException(400, "rowMax must be >= rowMin")
    sheet_id = _sheet_key(sheet_id)
    if settings.remote_only:
        store = _require_remote()
        cells = store.fetch_cells(
            project_id, _cloud_sheet_name(sheet_id), rowMin, rowMax
        )
        return {"cells": cells, "rowMin": rowMin, "rowMax": rowMax}
    store = _remote_store()
    if store:
        try:
            cells = store.fetch_cells(
                project_id, _cloud_sheet_name(sheet_id), rowMin, rowMax
            )
            if cells is not None:
                return {"cells": cells, "rowMin": rowMin, "rowMax": rowMax}
        except Exception as exc:
            raise HTTPException(503, str(exc)) from exc
    try:
        cells = local_store.get_cells(sheet_id, rowMin, rowMax)
        return {"cells": cells, "rowMin": rowMin, "rowMax": rowMax}
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc


@app.get("/api/v1/sheets/{sheet_id}/grid")
def sheet_grid(sheet_id: str, project_id: str = DEFAULT_PROJECT):
    """Precomputed display grid pack from the DB (replaces static *-grid.json)."""
    sheet_id = _sheet_key(sheet_id)
    store = _remote_store()
    if not store or not hasattr(store, "fetch_grid_pack"):
        raise HTTPException(501, "Grid cache requires Supabase REST")
    try:
        pack = store.fetch_grid_pack(project_id, _cloud_sheet_name(sheet_id))
    except Exception as exc:
        raise HTTPException(503, str(exc)) from exc
    if not pack:
        raise HTTPException(
            404,
            f"No grid pack for {sheet_id} — run: python -m tools.upload_grids_to_supabase",
        )
    return pack


class CellPatchBody(BaseModel):
    changes: list[dict[str, Any]]


@app.patch("/api/v1/sheets/{sheet_id}/cells")
def patch_sheet_cells(
    sheet_id: str,
    body: CellPatchBody,
    project_id: str = DEFAULT_PROJECT,
):
    sheet_id = _sheet_key(sheet_id)
    changes = body.changes or []
    if not changes:
        raise HTTPException(400, "changes[] required")
    store = _require_remote() if settings.remote_only else _remote_store()
    if not store or not hasattr(store, "upsert_cells"):
        raise HTTPException(501, "Cell PATCH requires Supabase REST")
    try:
        store.upsert_cells(project_id, _cloud_sheet_name(sheet_id), changes)
        return {"ok": True, "count": len(changes)}
    except Exception as exc:
        raise HTTPException(503, str(exc)) from exc


class ExportOverride(BaseModel):
    r: int
    c: str
    v: Any = None


class ExportBody(BaseModel):
    overrides: list[ExportOverride] = []


@app.post("/api/v1/sheets/{sheet_id}/export")
def export_sheet(sheet_id: str, body: ExportBody):
    """Stream the active sheet as a styled single-sheet .xlsx (openpyxl template).

    Independent of DATA_BACKEND: the client posts its current edits and we overlay
    them onto the pre-built template, so this works in local/supabase/databricks.
    """
    sheet_id = _sheet_key(sheet_id)
    try:
        data, filename = excel_export.build_export_bytes(
            sheet_id, [o.model_dump() for o in body.overrides]
        )
    except FileNotFoundError as exc:
        raise HTTPException(503, str(exc)) from exc
    return Response(
        content=data,
        media_type=excel_export.XLSX_MEDIA,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


class SessionPutBody(BaseModel):
    revision: int = 0
    structure_revision: int = 0
    bd: dict[str, Any] | None = None
    syn: dict[str, Any] | None = None
    updated_by: str = "web"


@app.get("/api/v1/sessions/{project_id}")
def get_session(project_id: str):
    store = _require_remote() if settings.remote_only else _remote_store()
    if not store:
        raise HTTPException(501, "Session API requires Supabase")
    try:
        session = store.fetch_session(project_id)
        if not session:
            raise HTTPException(404, "Session not found — run ingest")
        return session
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(503, str(exc)) from exc


@app.put("/api/v1/sessions/{project_id}")
def put_session(project_id: str, body: SessionPutBody):
    store = _require_remote() if settings.remote_only else _remote_store()
    if not store:
        raise HTTPException(501, "Session API requires Supabase")
    try:
        struct_rev = body.structure_revision
        if not struct_rev and body.bd:
            struct_rev = int(body.bd.get("structureRevision") or 0)
        return store.upsert_session(
            project_id,
            body.revision,
            body.bd,
            body.syn,
            body.updated_by,
            structure_revision=struct_rev,
        )
    except Exception as exc:
        raise HTTPException(503, str(exc)) from exc


# Code/données servis sans cache navigateur. Sans cela, Starlette n'envoie qu'un
# ETag/Last-Modified et Edge ressert un module ES périmé sous la même URL
# `?v=...` (ex: navConfig.js d'avant l'ajout de NAV_ROUTES) → "does not provide
# an export named ...". Le serveur node fait déjà ce `no-store` ; on l'aligne ici.
_NO_STORE_SUFFIXES = (".html", ".js", ".mjs", ".css", ".json")


class NoStoreStaticFiles(StaticFiles):
    def is_not_modified(self, response_headers, request_headers) -> bool:  # noqa: ANN001
        # Empêche les réponses 304 pour le code/données : on veut toujours frais.
        if response_headers.get("content-type", "").split(";")[0] in (
            "text/html",
            "text/javascript",
            "application/javascript",
            "text/css",
            "application/json",
        ):
            return False
        return super().is_not_modified(response_headers, request_headers)

    async def get_response(self, path: str, scope):  # noqa: ANN001
        response = await super().get_response(path, scope)
        if any(path.endswith(suffix) for suffix in _NO_STORE_SUFFIXES):
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            for validator in ("etag", "last-modified"):
                if validator in response.headers:
                    del response.headers[validator]
        return response


# Sert l'app web (index.html, /js, /css, /public/data, ...) sur le meme port que
# l'API. Monte en DERNIER pour que les routes /api/v1/... gardent la priorite.
# C'est ce qui corrige le "HTTP 404 for /public/data/bd-sheet.json" : le front
# n'a plus besoin de __WGHT_API_BASE__, tout est servi en same-origin.
WEB_DIR = Path(__file__).resolve().parent.parent.parent / "web"
if WEB_DIR.is_dir():
    app.mount("/", NoStoreStaticFiles(directory=str(WEB_DIR), html=True), name="web")
else:
    logger.warning("Dossier web introuvable, fichiers statiques non servis: %s", WEB_DIR)
