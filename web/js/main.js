import {
  createApp,
  ref,
  shallowRef,
  computed,
  watch,
  onMounted,
  onUnmounted,
  onErrorCaptured,
  nextTick,
  provide,
} from 'vue';
import BdGrid from './BdGrid.js?v=bd-ugap-edit1';
import { computeBodyDisplayRows } from './bdStore.js';
import SynthesisGrid from './SynthesisGrid.js?v=syn-row16-link1';
import { createEditHistory } from './editHistory.js?v=undo4';
import AppSidebar from './AppSidebar.js?v=syn-perf32';
import EmptyPage from './EmptyPage.js?v=syn-perf32';
import MnsGrid from './MnsGrid.js?v=sp2-col-a-meta1';
import CdcOutputGrid from './CdcOutputGrid.js?v=cdc-output7';
import { resolveSynRow16DisplayFromRaw } from './sp2CurbLink.js?v=5';
import MatrixModal from './MatrixModal.js?v=matrix-ca-syn1';
import { NAV_ROUTES, DEFAULT_ROUTE } from './navConfig.js?v=syn-perf32';
import { transformBdSheetAsync, transformSynthesisSheetAsync } from './sheetTransform.js?v=grid-perf9';
import {
  collectRowCells,
  collectColCells,
  pasteRowCells,
  pasteColCells,
} from './sheetRowColOps.js?v=axis2';
import { addUserRowGap, addUserColGap, isUserGapCol, cloneUserGapsSnapshot, clearUserGaps } from './gridUserGaps.js?v=ugap3';
import {
  synGridLooksHealthy,
  isSynStaleProjectExportText,
} from './synStore.js?v=syn-project-fix1';
import { createWorkbookSession } from './workbookSession.js?v=realtime-sync1';
import { startSynthesisRealtime } from './synthesisRealtime.js?v=realtime1';
import { createPerfBench } from './perfBench.js?v=1';
import { yieldToMain, deferHeavy } from './yieldMain.js?v=3';
import {
  fetchPrecomputedPack,
  resolveGridSheet,
  hasSheetEdits,
} from './gridBoot.js?v=row-del-persist1';
import {
  loadSheetRaw,
  probeSheetApi,
  patchSheetCells,
  fetchSession,
  exportSheetXlsx,
} from './sheetDataApi.js?v=xlsx-export1';
import {
  buildMatrixState,
  applyMatrixSave,
  alignSynModelToBd,
  cloneStructure,
} from './structureModel.js?v=bookmark-sync1';
import {
  extractSheetEdits,
  upsertRawCell,
  loadLocalSnapshot,
  clearLocalSnapshot,
  createAutoSave,
  applySheetEdits,
  saveSheetTransform,
  rawFingerprint,
  syncGridEditsToRaw,
  synRawLooksHealthy,
  clearSheetTransform,
  dropStaleStructure,
} from './sessionPersistence.js?v=canonical-structure1';

const App = {
  components: { BdGrid, SynthesisGrid, MnsGrid, CdcOutputGrid, AppSidebar, EmptyPage, MatrixModal },
  setup() {
    const bdLoading = ref(false);
    const gridPreparing = ref(false);
    const synthesisLoading = ref(false);
    /** @type {import('vue').Ref<{ sheetId: string, pct: number } | null>} */
    const dataLoadProgress = ref(null);
    const chunkedApi = ref(false);
    const serverCalc = ref(false);
    const remoteOnly = ref(false);
    const error = ref(null);
    // Large data objects (149k+ cells + cellMap): shallowRef avoids Vue deep-proxying
    // every cell, which otherwise blocks the main thread for seconds on load and on
    // every edit. Reassigning .value still triggers re-render; internal cell mutations
    // are signaled via dedicated tick refs (bdSheetRevision, externalEditTick, etc.).
    const bdSheet = shallowRef(null);
    const synthesisSheet = shallowRef(null);
    const bdRaw = shallowRef(null);
    const synRaw = shallowRef(null);
    const matrixOpen = ref(false);
    const matrixState = ref(null);
    const matrixSaving = ref(false);
    const dirty = ref(0);
    const saveStatus = ref('idle');
    const saveError = ref('');
    const loadedFromLocal = ref(false);
    const route = ref(DEFAULT_ROUTE);
    const menuOpen = ref(false);
    // Outline view cycle: 0 = full, 1 = sections + sub-sections, 2 = sections only.
    const outlineMode = ref(0);
    const outlineOnly = computed(() => outlineMode.value !== 0);
    const searchOpen = ref(false);
    const searchQuery = ref('');
    const searchInputRef = ref(null);
    const gridSearchCmd = ref(null);
    const searchStatus = ref('');
    let searchHiddenRetry = false;
    let searchDebounceTimer = 0;
    const session = createWorkbookSession();
    const editHistory = createEditHistory();
    const historyTick = ref(0);
    const externalEditTick = ref(0);
    /** >0 after Bookmark Matrix save — triggers full-sheet local persistence. */
    const structureRevision = ref(0);
    let applyingHistory = false;
    let engineStarted = false;
    let synthesisLoadPromise = null;
    let synthesisPreparePromise = null;
    let matrixPendingModel = null;
    let matrixSessionBefore = null;
    let matrixSessionDirty = false;
    let matrixApplyQueued = false;
    let matrixApplyDebounceTimer = null;
    /** Keep last known synthesis raw when BD-only saves run before Syn is opened. */
    let preservedSynRaw = null;
    let synRawLoadPromise = null;
    /** Skip re-running transformBdSheet / transformSynthesisSheet when raw unchanged. */
    const bdSheetRevision = ref(0);
    const synSheetRevision = ref(0);
    let bdSheetBuiltAt = -1;
    let synSheetBuiltAt = -1;

    function buildAutoSavePayload() {
      if (bdSheet.value && bdRaw.value) {
        syncGridEditsToRaw(bdSheet.value, bdRaw.value);
      }
      let syn = synRaw.value != null ? synRaw.value : preservedSynRaw;
      if (synthesisSheet.value) {
        if (!syn) syn = synRaw.value;
        if (syn) syncGridEditsToRaw(synthesisSheet.value, syn);
      }
      const structRev = structureRevision.value;
      const bd =
        bdRaw.value != null
          ? { ...bdRaw.value, structureRevision: structRev }
          : null;
      const synOut =
        syn != null ? { ...syn, structureRevision: structRev } : null;
      return {
        bd,
        syn: synOut,
        revision: dirty.value,
        structureRevision: structRev,
      };
    }

    const autoSave = createAutoSave(
      () => buildAutoSavePayload(),
      {
        debounceMs: 400,
        onStatus(status, err) {
          saveStatus.value = status;
          if (err) saveError.value = (err && err.message) || String(err);
          else if (status !== 'error') saveError.value = '';
        },
      }
    );

    const currentNav = computed(
      () => NAV_ROUTES.find((n) => n.id === route.value) || NAV_ROUTES[0]
    );

    const isDatabase = computed(() => route.value === 'database');
    const isSynthesis = computed(() => route.value === 'synthesis');
    const isMns = computed(() => route.value === 'cdc-mns');
    const isOptionsSp2 = computed(() => route.value === 'cdc-options-sp2');
    const isCdcOutput = computed(() => route.value === 'cdc-output');
    const isGridPage = computed(
      () => route.value === 'database' || route.value === 'synthesis'
    );
    // `session` is a plain object, so its nested `loading` ref does NOT auto-unwrap
    // in templates (Vue only unwraps refs nested in reactive objects). Expose it as
    // a top-level computed so `!sessionLoading` in the topbar evaluates correctly.
    const sessionLoading = computed(() => session.loading.value);
    // Legacy topbar actions (Enregistrer / Restaurer Synthesis / Fichiers projet)
    // are kept in the template but hidden — flip to true to show them again.
    // Edits still persist via autosave; Export Excel stays visible regardless.
    const showLegacyTopbarActions = false;

    const canUndo = computed(() => {
      void historyTick.value;
      return editHistory.canUndo;
    });
    const canRedo = computed(() => {
      void historyTick.value;
      return editHistory.canRedo;
    });

    const overlayMessage = computed(() => {
      if (isSynthesis.value && !synthesisSheet.value) {
        return synthesisLoading.value ? 'Loading synthesis…' : 'Preparing synthesis…';
      }
      if (bdLoading.value) return 'Loading database…';
      return 'Loading…';
    });

    const showContentOverlay = computed(() => {
      if (error.value) return false;
      if (isDatabase.value) return false;
      if (isSynthesis.value && !synthesisSheet.value) return true;
      return false;
    });

    function slimBdEnginePayload(sheet) {
      return {
        columns: sheet.columns,
        lastRow: sheet.lastRow,
        cells: sheet.cells,
        headerRows: sheet.headerRows,
        sectionHeaderRows: sheet.sectionHeaderRows,
        canonicalSectionByLabel: sheet.canonicalSectionByLabel,
        cellMap: sheet.cellMap,
      };
    }

    /** Sheet object last pushed into the calc engine — guards against redundant reloads. */
    let engineLoadedForSheet = null;

    /**
     * Build BD column index for live SYN SUMPRODUCT as soon as the grid sheet exists.
     * Reloading the engine re-indexes ~149k cells, clears every SUMPRODUCT cache and
     * re-posts the whole BD payload to the worker — so it must run ONCE per BD sheet,
     * never on every navigation. `bdSheet.value` is a fresh object whenever the data
     * actually changes (applyBdFromRaw / matrix / reset reassign it), so comparing the
     * reference is enough to skip the no-op reloads that made page switches crawl.
     */
    function syncBdCalcEngine(force = false) {
      if (!bdSheet.value) return;
      if (!force && engineStarted && engineLoadedForSheet === bdSheet.value) return;
      engineLoadedForSheet = bdSheet.value;
      void session
        .loadSheets([{ name: 'BD', data: slimBdEnginePayload(bdSheet.value) }])
        .then(() => {
          engineStarted = true;
          externalEditTick.value += 1;
        });
    }

    watch(
      () => bdSheet.value,
      (sheet) => {
        if (sheet) syncBdCalcEngine();
      }
    );

    function syncRawMetaOntoSheet(raw, sheet) {
      if (!sheet || !raw) return;
      if (Array.isArray(raw.deletedRows) && raw.deletedRows.length) {
        sheet.deletedRows = [...raw.deletedRows];
      }
      if (Array.isArray(raw.deletedCols) && raw.deletedCols.length) {
        sheet.deletedCols = [...raw.deletedCols];
      }
      sheet.userRowGaps = Array.isArray(raw.userRowGaps)
        ? raw.userRowGaps.map((g) => ({ ...g }))
        : [];
      sheet.userColGaps = Array.isArray(raw.userColGaps)
        ? raw.userColGaps.map((g) => ({ ...g }))
        : [];
      sheet.userGapCells = Array.isArray(raw.userGapCells)
        ? raw.userGapCells.map((c) => ({ ...c }))
        : [];
    }

    function syncDeletedRowsOntoSheet(raw, sheet) {
      syncRawMetaOntoSheet(raw, sheet);
    }

    async function applyBdFromRaw(force = false) {
      if (force) await clearSheetTransform('bd');
      if (!bdRaw.value) {
        if (!force) {
          const pre = await fetchPrecomputedPack('bd');
          if (pre && pre.sheet) {
            await yieldToMain();
            bdSheet.value = pre.sheet;
            bdSheetBuiltAt = bdSheetRevision.value;
            return;
          }
        }
        return;
      }
      if (!force && bdSheet.value && bdSheetBuiltAt === bdSheetRevision.value) {
        syncDeletedRowsOntoSheet(bdRaw.value, bdSheet.value);
        bdSheet.value.bodyDisplayRows = computeBodyDisplayRows(bdSheet.value);
        return;
      }
      const resolved = await resolveGridSheet('bd', bdRaw.value, { force });
      if (resolved) {
        await yieldToMain();
        bdSheet.value = resolved;
        syncDeletedRowsOntoSheet(bdRaw.value, bdSheet.value);
        bdSheet.value.bodyDisplayRows = computeBodyDisplayRows(bdSheet.value);
        bdSheetBuiltAt = bdSheetRevision.value;
        return;
      }
      await yieldToMain();
      const transformed = await transformBdSheetAsync(bdRaw.value);
      await yieldToMain();
      bdSheet.value = transformed;
      syncDeletedRowsOntoSheet(bdRaw.value, bdSheet.value);
      bdSheet.value.bodyDisplayRows = computeBodyDisplayRows(bdSheet.value);
      bdSheetBuiltAt = bdSheetRevision.value;
      const fp = rawFingerprint(bdRaw.value);
      const pre = await fetchPrecomputedPack('bd');
      void saveSheetTransform('bd', fp, bdSheet.value, {
        skipIfPrecomputed: Boolean(pre && pre.fingerprint === fp),
      });
    }

    /**
     * Seed a transform-built Synthesis sheet with the precomputed pack's materialized
     * live-mass values (M…AA / AC…AN / AP…BB, ADAPTATION total, AB Δ). Without this, a
     * sheet rebuilt from raw (because local edits changed the fingerprint and the pack no
     * longer matches) has empty M…AA cells until the slow live SUMPRODUCT warms up — which
     * is the "blank for ~1 minute on load" bug. With the seed, every cell shows its last
     * recorded value instantly; the live engine then refreshes impacted cells in real time.
     * Non-destructive: never touches user-edited cells nor cells that already hold a value.
     * @returns {number} count of cells seeded
     */
    async function seedSynMaterializedFromPack(sheet) {
      if (!sheet) return 0;
      const map = sheet.cellMap instanceof Map ? sheet.cellMap : null;
      if (!map) return 0;
      let pre = null;
      try {
        pre = await fetchPrecomputedPack('syn');
      } catch {
        return 0;
      }
      if (!pre || !pre.sheet || !Array.isArray(pre.sheet.cells)) return 0;
      let seeded = 0;
      for (const pc of pre.sheet.cells) {
        if (!pc || !pc.mat) continue;
        const key = `${pc.r}:${pc.c}`;
        const existing = map.get(key);
        if (existing) {
          if (existing.userEdited) continue;
          const cur = existing.v != null ? String(existing.v).trim() : '';
          if (cur !== '' && !isSynStaleProjectExportText(cur)) continue;
          existing.v = pc.v;
          existing.mat = true;
        } else {
          const cell = { r: pc.r, c: pc.c, v: pc.v, mat: true };
          sheet.cells.push(cell);
          map.set(key, cell);
        }
        seeded++;
      }
      return seeded;
    }

    async function applySynFromRaw(force = false) {
      if (force) await clearSheetTransform('syn');
      if (!synRaw.value) {
        if (!force) {
          const pre = await fetchPrecomputedPack('syn');
          if (pre && pre.sheet) {
            await yieldToMain();
            synthesisSheet.value = pre.sheet;
            synSheetBuiltAt = synSheetRevision.value;
            return;
          }
        }
        return;
      }
      if (!force && synthesisSheet.value && synSheetBuiltAt === synSheetRevision.value) {
        syncDeletedRowsOntoSheet(synRaw.value, synthesisSheet.value);
        return;
      }
      const resolved = await resolveGridSheet('syn', synRaw.value, { force });
      if (resolved) {
        await yieldToMain();
        await seedSynMaterializedFromPack(resolved);
        synthesisSheet.value = resolved;
        syncDeletedRowsOntoSheet(synRaw.value, synthesisSheet.value);
        synSheetBuiltAt = synSheetRevision.value;
        return;
      }
      await yieldToMain();
      const transformed = await transformSynthesisSheetAsync(synRaw.value);
      await yieldToMain();
      await seedSynMaterializedFromPack(transformed);
      synthesisSheet.value = transformed;
      syncDeletedRowsOntoSheet(synRaw.value, synthesisSheet.value);
      synSheetBuiltAt = synSheetRevision.value;
      const fp = rawFingerprint(synRaw.value);
      const pre = await fetchPrecomputedPack('syn');
      void saveSheetTransform('syn', fp, synthesisSheet.value, {
        skipIfPrecomputed: Boolean(pre && pre.fingerprint === fp),
      });
    }

    function bumpBdSheetRevision() {
      bdSheetRevision.value += 1;
    }

    function bumpSynSheetRevision() {
      synSheetRevision.value += 1;
    }

    /** Load synthesis-sheet.json once — required before any edit can persist. */
    async function ensureSynRaw() {
      if (synRaw.value) return synRaw.value;
      if (synRawLoadPromise) return synRawLoadPromise;
      synRawLoadPromise = (async () => {
        const raw = await loadSheetRaw('synthesis');
        synRaw.value = raw;
        preservedSynRaw = raw;
        return raw;
      })()
        .catch((e) => {
          console.warn('ensureSynRaw failed:', e);
          throw e;
        })
        .finally(() => {
          synRawLoadPromise = null;
        });
      return synRawLoadPromise;
    }

    // Live link for the Options SP2 page: lets the CDC grid read Synthesis cells
    // (e.g. the Curb mass row 16) reactively. synSheetRevision bumps on every
    // synthesis change (edits, server recalc, Supabase session load), so any cell
    // bound through this link updates in real time.
    /** SynthesisGrid registers cellDisplay(16, col) — exact on-screen Curb mass row. */
    const synRow16DisplayFn = shallowRef(null);

    function getSynRow16Display(excelCol) {
      const fn = synRow16DisplayFn.value;
      if (typeof fn === 'function') {
        try {
          const live = fn(String(excelCol));
          if (live != null && String(live).trim() !== '') return String(live);
        } catch (e) {
          console.warn('[syn-row16]', e);
        }
      }
      const raw = synRaw.value;
      if (raw && Array.isArray(raw.cells)) {
        return resolveSynRow16DisplayFromRaw(raw.cells, excelCol);
      }
      return '';
    }

    provide('synthesisCellLink', {
      synRaw,
      synRevision: synSheetRevision,
      // Bumped on every synthesis cell edit / recalc (synSheetRevision is not),
      // and synRaw is a shallowRef mutated in place — so consumers must watch this
      // tick to react to live edits.
      synEditTick: externalEditTick,
      session,
      ensureSyn: ensureSynRaw,
      getSynRow16Display,
      registerSynRow16Display(fn) {
        synRow16DisplayFn.value = fn;
      },
      unregisterSynRow16Display(fn) {
        if (synRow16DisplayFn.value === fn) synRow16DisplayFn.value = null;
      },
      ensureSynGrid: startSynBackgroundPrepare,
    });

    /** Options SP2 row 15 (F…T) — shared live values for MNS row 15 (Y…AM). */
    const sp2Row15FtSums = shallowRef({});
    const sp2Row15Tick = ref(0);

    provide('optionsSp2CellLink', {
      row15FtSums: sp2Row15FtSums,
      sp2Row15Tick,
      publishRow15FtSums(sums) {
        sp2Row15FtSums.value =
          sums && typeof sums === 'object' ? { ...sums } : {};
        sp2Row15Tick.value += 1;
      },
    });

    // ── Supabase Realtime ──────────────────────────────────────────────────
    // Quand une cellule Synthesis change cote base (autre onglet / utilisateur),
    // on l'applique a synRaw en memoire et on bump le tick d'edition : la ligne 14
    // "Curb mass" d'Options SP2 (liee a la ligne 16 de Synthesis) se met a jour en
    // direct, sans rechargement. Best-effort : ne casse rien si Realtime indisponible.
    let synRealtimeHandle = null;

    function applyRealtimeSynCell({ row, col, v }) {
      const raw = synRaw.value;
      if (!raw || !Array.isArray(raw.cells)) return;
      let cell = raw.cells.find(
        (x) => Number(x.r) === Number(row) && String(x.c) === String(col)
      );
      if (cell) {
        cell.v = v == null ? '' : v;
        cell.userEdited = true;
        delete cell.f;
        delete cell.mat;
      } else if (v != null && String(v) !== '') {
        raw.cells.push({ r: Number(row), c: String(col), v, userEdited: true });
      } else {
        return;
      }
      if (synthesisSheet.value) {
        syncSheetCell(synthesisSheet.value, Number(row), String(col), v);
      }
      // synRaw is a shallowRef mutated in place — bump the tick so consumers
      // (Options SP2 Curb mass link) recompute.
      externalEditTick.value += 1;
      bumpSynSheetRevision();
    }

    async function startRealtimeSync(apiCfg) {
      if (synRealtimeHandle || !apiCfg) return;
      if (!apiCfg.supabaseUrl || !apiCfg.supabaseAnonKey) return;
      try {
        synRealtimeHandle = await startSynthesisRealtime({
          supabaseUrl: apiCfg.supabaseUrl,
          supabaseAnonKey: apiCfg.supabaseAnonKey,
          projectId: apiCfg.projectId || 'default',
          onSynCellChange: applyRealtimeSynCell,
        });
      } catch (e) {
        console.warn('[syn-realtime] demarrage impossible:', e);
      }
    }

    onUnmounted(() => {
      if (synRealtimeHandle && synRealtimeHandle.stop) synRealtimeHandle.stop();
      synRealtimeHandle = null;
    });

    /** Fetch + transform Synthesis when user opens the page (never while on Database). */
    function startSynBackgroundPrepare() {
      if (
        synthesisSheet.value &&
        synRaw.value &&
        synSheetBuiltAt === synSheetRevision.value &&
        synGridLooksHealthy(synthesisSheet.value)
      ) {
        return Promise.resolve();
      }
      if (synthesisPreparePromise) return synthesisPreparePromise;
      synthesisPreparePromise = (async () => {
        await ensureSynRaw();
        if (!synRawLooksHealthy(synRaw.value)) {
          await fetchSynFromServer();
          preservedSynRaw = synRaw.value;
        }
        // Prefer precomputed / IndexedDB grid — never force a full in-browser transform
        // on every navigation (that was the ~2 min "Calcul Synthesis…" wait).
        await applySynFromRaw(false);
        if (!synGridLooksHealthy(synthesisSheet.value)) {
          await clearSheetTransform('syn');
          await fetchSynFromServer();
          preservedSynRaw = synRaw.value;
          bumpSynSheetRevision();
          await applySynFromRaw(true);
        }
      })().catch((e) => {
        synthesisPreparePromise = null;
        console.warn('Synthesis prepare failed:', e);
        throw e;
      });
      return synthesisPreparePromise;
    }

    function mergeExtraCellsIntoBd(fullRaw) {
      const sheet = bdSheet.value;
      if (!sheet || !fullRaw || !fullRaw.cells || !fullRaw.cells.length) return;
      const map = sheet.cellMap instanceof Map ? sheet.cellMap : null;
      if (!map) {
        bumpBdSheetRevision();
        return applyBdFromRaw(false);
      }
      let added = 0;
      for (const c of fullRaw.cells) {
        const k = `${c.r}:${c.c}`;
        if (!map.has(k)) {
          sheet.cells.push(c);
          map.set(k, c);
          added++;
        }
      }
      if (added > 0) externalEditTick.value += 1;
    }

    /**
     * Load BD raw — chunked API when available (meta + row windows), else full JSON.
     * @param {{ progressive?: boolean }} opts progressive=true → show grid after 1st chunk
     */
    async function loadBdFromServer({ progressive = true } = {}) {
      bdLoading.value = true;
      let earlyRender = false;
      try {
        const full = await loadSheetRaw('bd', {
          onProgress(info) {
            if (info.lastRow > 0) {
              dataLoadProgress.value = {
                sheetId: 'bd',
                pct: Math.min(
                  100,
                  Math.round((info.rowMax / info.lastRow) * 100)
                ),
              };
            }
          },
          async onFirstChunk(partial) {
            if (!progressive || earlyRender) return;
            earlyRender = true;
            bdRaw.value = partial;
            bumpBdSheetRevision();
            await applyBdFromRaw(false);
            gridPreparing.value = false;
            syncBdCalcEngine();
          },
        });
        bdRaw.value = full;
        if (!earlyRender) {
          bumpBdSheetRevision();
          await applyBdFromRaw(false);
          gridPreparing.value = false;
          syncBdCalcEngine();
        } else {
          mergeExtraCellsIntoBd(full);
        }
      } finally {
        bdLoading.value = false;
        dataLoadProgress.value = null;
      }
    }

    async function fetchBdFromServer() {
      await loadBdFromServer({ progressive: false });
    }

    /** Force a full Synthesis grid rebuild from synRaw (bypasses IDB / stale Vue sheet). */
    async function rebuildSynthesisGridFromRaw() {
      await clearSheetTransform('syn');
      if (!synRaw.value) await fetchSynFromServer();
      preservedSynRaw = synRaw.value;
      synthesisSheet.value = null;
      synthesisPreparePromise = null;
      synthesisLoadPromise = null;
      synSheetBuiltAt = -1;
      bumpSynSheetRevision();
      await applySynFromRaw(true);
      if (!synGridLooksHealthy(synthesisSheet.value)) {
        throw new Error(
          'Synthesis grid still invalid after rebuild (ADAPTATION band missing).'
        );
      }
    }

    /** Reload synthesis-sheet.json from the project (fixes a corrupted local copy). */
    async function restoreSynthesisFromProject() {
      if (
        !confirm(
          "Restaurer Synthesis depuis le fichier projet ?\n\nEfface la copie locale Synthesis corrompue et recharge ADAPTATION (ligne Excel 25, affichage ~26).\nDatabase / Bookmark Matrix sont conservés."
        )
      ) {
        return;
      }
      synthesisLoading.value = true;
      try {
        await clearSheetTransform('syn');
        await fetchSynFromServer();
        preservedSynRaw = synRaw.value;
        if (session && session.liveBdEdited) session.liveBdEdited.value = false;
        structureRevision.value = 0;
        await rebuildSynthesisGridFromRaw();
        externalEditTick.value += 1;
        dirty.value += 1;
        await autoSave.saveNow();
        saveStatus.value = 'saved';
      } catch (e) {
        error.value = (e && e.message) || String(e);
        console.error('restoreSynthesisFromProject failed:', e);
      } finally {
        synthesisLoading.value = false;
      }
    }

    async function healSynRawIfCorrupted(reason) {
      const rawBad = synRaw.value && !synRawLooksHealthy(synRaw.value);
      const gridBad =
        synthesisSheet.value && !synGridLooksHealthy(synthesisSheet.value);
      if (!rawBad && !gridBad) return false;
      console.warn(
        `[syn] Copie locale invalide (${reason}) — rechargement synthesis-sheet.json`
      );
      if (rawBad || !synRaw.value) {
        await fetchSynFromServer();
        preservedSynRaw = synRaw.value;
      }
      await rebuildSynthesisGridFromRaw();
      return true;
    }

    async function fetchSynFromServer() {
      synthesisLoading.value = true;
      try {
        synRaw.value = await loadSheetRaw('synthesis', {
          onProgress(info) {
            if (info.lastRow > 0) {
              dataLoadProgress.value = {
                sheetId: 'syn',
                pct: Math.min(
                  100,
                  Math.round((info.rowMax / info.lastRow) * 100)
                ),
              };
            }
          },
        });
        clearUserGaps(synRaw.value);
        preservedSynRaw = synRaw.value;
      } finally {
        dataLoadProgress.value = null;
        synthesisLoading.value = false;
      }
    }

    async function loadSynthesis() {
      if (
        synthesisSheet.value &&
        synGridLooksHealthy(synthesisSheet.value)
      ) {
        return;
      }
      if (synthesisLoadPromise) return synthesisLoadPromise;
      synthesisLoading.value = true;
      synthesisLoadPromise = (async () => {
        // Instant paint: DB / build-time grid already holds materialized values.
        if (!synthesisSheet.value) {
          const pre = await fetchPrecomputedPack('syn');
          if (pre && pre.sheet && synGridLooksHealthy(pre.sheet)) {
            synthesisSheet.value = pre.sheet;
            synSheetBuiltAt = synSheetRevision.value;
            void startSynBackgroundPrepare();
            return;
          }
        }
        await startSynBackgroundPrepare();
      })()
        .catch((e) => {
          error.value = (e && e.message) || String(e);
          synthesisLoadPromise = null;
        })
        .finally(() => {
          synthesisLoading.value = false;
        });
      return synthesisLoadPromise;
    }

    onErrorCaptured((err) => {
      error.value = (err && err.message) || String(err);
      synthesisLoading.value = false;
      bdLoading.value = false;
      console.error('Grid render failed:', err);
      return false;
    });

    onMounted(() => {
      document.title = 'WGHT Dashboard';
      const routeParam = new URLSearchParams(location.search).get('route');
      if (routeParam && NAV_ROUTES.some((n) => n.id === routeParam)) {
        route.value = routeParam;
      }
      if (typeof window !== 'undefined') {
        syncInitialRouteToUrl();
        window.addEventListener('popstate', onPopState);
      }
      // Optional perf tracing (edit → calc → render-ish).
      // Enable via URL: ?perf=1
      if (typeof window !== 'undefined') {
        window.__perfEdits = new URLSearchParams(location.search).get('perf') === '1';
      }
      const bench = createPerfBench(() => ({
        session,
        bdSheet: bdSheet.value,
        synthesisSheet: synthesisSheet.value,
        applyBdFromRaw,
        applySynFromRaw,
      }));
      if (typeof window !== 'undefined') {
        window.__runPerfBench = () => bench.run();
        // Debug helpers (console) for verifying formula dependencies.
        // Usage:
        //   __synFormula(27,'G')  -> describes how the cell is computed
        //   __bdCell(9,'V')       -> current BD raw value at row/col
        //   __synCell(27,'G')     -> current SYN raw value at row/col (may be empty if computed)
        window.__synFormula = (row, col) => {
          try {
            if (session && typeof session.getSynFormula === 'function') {
              return session.getSynFormula(Number(row), String(col));
            }
            return '';
          } catch (e) {
            return String((e && e.message) || e);
          }
        };
        window.__bdCell = (row, col) => {
          try {
            if (!bdRaw.value) return null;
            const r = Number(row);
            const c = String(col);
            const hit = (bdRaw.value.cells || []).find((x) => x.r === r && x.c === c);
            return hit ? hit.v : '';
          } catch (e) {
            return String((e && e.message) || e);
          }
        };
        window.__synCell = (row, col) => {
          try {
            if (!synRaw.value) return null;
            const r = Number(row);
            const c = String(col);
            const hit = (synRaw.value.cells || []).find((x) => x.r === r && x.c === c);
            return hit ? hit.v : '';
          } catch (e) {
            return String((e && e.message) || e);
          }
        };
        window.__restoreSynthesis = () => restoreSynthesisFromProject();
      }
      void bootApp();
      window.addEventListener('beforeunload', onBeforeUnload);
      window.addEventListener('pagehide', onBeforeUnload);
      window.addEventListener('keydown', onGlobalKeydown, true);
    });

    onUnmounted(() => {
      if (typeof window !== 'undefined') {
        window.removeEventListener('popstate', onPopState);
      }
    });

    async function bootApp() {
      const wantSynthesis = route.value === 'synthesis';
      void fetchPrecomputedPack('syn');

      try {
        await nextTick();
        const [snapshot, apiCfg, preBd] = await Promise.all([
          loadLocalSnapshot('default', { timeoutMs: 1500 }),
          probeSheetApi(),
          fetchPrecomputedPack('bd'),
        ]);
        // A frozen full-BD snapshot from an older structure schema would override the
        // fresh base data + current code and bring back old/duplicated bookmarks. Drop
        // its structural part so the structure is rebuilt from the project JSON; the
        // user's plain cell edits are kept.
        if (dropStaleStructure(snapshot)) {
          console.info(
            'Local session: stale structure snapshot discarded — rebuilding from project files.'
          );
        }

        chunkedApi.value = Boolean(apiCfg && apiCfg.chunkedLoad);
        serverCalc.value = Boolean(apiCfg && apiCfg.serverCalc);
        remoteOnly.value = Boolean(apiCfg && apiCfg.remoteOnly);

        // Abonnement Supabase Realtime (Synthesis -> Options SP2 ligne 14, en direct).
        void startRealtimeSync(apiCfg);

        if (remoteOnly.value && (!apiCfg || apiCfg.mode === 'static')) {
          error.value =
            'Supabase requis : 1) run-supabase-server.bat  2) run-bd-server.bat  3) SUPABASE_SERVICE_KEY dans api\\.env';
          return;
        }

        let remoteWon = false;
        let remoteRev = -1;
        let remoteStruct = 0;

        const cloudApi =
          apiCfg &&
          (apiCfg.cloudPersist ||
            apiCfg.mode === 'databricks' ||
            apiCfg.mode === 'supabase' ||
            apiCfg.mode === 'postgres');
        if (cloudApi) {
          try {
            const remote = await fetchSession(apiCfg.projectId || 'default');
            if (remote && remote.bd) {
              remoteRev = Number(remote.revision) || 0;
              const localRev =
                snapshot && snapshot.revision != null ? Number(snapshot.revision) : -1;
              remoteStruct =
                Number(remote.structureRevision) ||
                Number(remote.bd.structureRevision) ||
                0;
              const localStruct =
                snapshot && snapshot.structureRevision != null
                  ? Number(snapshot.structureRevision)
                  : 0;
              if (
                !snapshot ||
                remoteRev > localRev ||
                (remoteRev === localRev && remoteStruct > localStruct)
              ) {
                bdRaw.value = remote.bd;
                if (remote.syn) {
                  synRaw.value = remote.syn;
                  preservedSynRaw = remote.syn;
                }
                structureRevision.value = remoteStruct;
                remoteWon = true;
                dirty.value = Math.max(dirty.value, remoteRev);
                saveStatus.value = 'saved';
                bumpBdSheetRevision();
                bumpSynSheetRevision();
              }
            }
          } catch (e) {
            console.warn('Cloud session load:', e);
          }
        }

        const bdHasEdits =
          (snapshot && snapshot.version === 3 && hasSheetEdits(snapshot.bdEdits)) ||
          (snapshot && snapshot.version === 2 && hasSheetEdits(snapshot.bdEdits));
        const synHasEdits =
          (snapshot && snapshot.version === 3 && hasSheetEdits(snapshot.synEdits)) ||
          (snapshot && snapshot.version === 2 && hasSheetEdits(snapshot.synEdits));
        const hasBdFull = snapshot && snapshot.version === 3 && snapshot.bdFull;
        const hasSynFull = snapshot && snapshot.version === 3 && snapshot.synFull;

        const localRev =
          snapshot && snapshot.revision != null ? Number(snapshot.revision) : 0;
        const localStruct =
          snapshot && snapshot.structureRevision != null
            ? Number(snapshot.structureRevision)
            : 0;
        const localWinsData =
          remoteOnly.value
            ? false
            : !remoteWon ||
              localRev > remoteRev ||
              (localRev === remoteRev && localStruct > remoteStruct);

        function mergeLocalEditsOntoCurrent() {
          if (bdHasEdits && bdRaw.value) {
            applySheetEdits(bdRaw.value, snapshot.bdEdits);
            bumpBdSheetRevision();
          }
          if (synHasEdits && synRaw.value) {
            applySheetEdits(synRaw.value, snapshot.synEdits);
            bumpSynSheetRevision();
          }
          if (bdHasEdits || synHasEdits) {
            loadedFromLocal.value = true;
          }
        }

        if (snapshot && snapshot.version === 3) {
          if (localWinsData) {
            if (snapshot.bdFull) {
              bdRaw.value = snapshot.bdFull;
            } else if (bdHasEdits) {
              if (!bdRaw.value) await fetchBdFromServer();
              applySheetEdits(bdRaw.value, snapshot.bdEdits);
              bumpBdSheetRevision();
            }
            if (!synRaw.value) await fetchSynFromServer();
            preservedSynRaw = synRaw.value;
            if (synHasEdits && synRaw.value) {
              applySheetEdits(synRaw.value, snapshot.synEdits);
            }
            bumpSynSheetRevision();
            structureRevision.value = localStruct;
            loadedFromLocal.value = true;
            saveStatus.value = 'saved';
            dirty.value = Math.max(dirty.value, localRev);
          } else if (remoteWon) {
            mergeLocalEditsOntoCurrent();
            structureRevision.value = Math.max(structureRevision.value, localStruct);
            dirty.value = Math.max(dirty.value, localRev);
          }
        } else if (snapshot && snapshot.version === 2) {
          if (localWinsData || !remoteWon) {
            if (snapshot.bdEdits && hasSheetEdits(snapshot.bdEdits)) {
              if (!bdRaw.value) await fetchBdFromServer();
              applySheetEdits(bdRaw.value, snapshot.bdEdits);
              bumpBdSheetRevision();
            }
            if (!synRaw.value) await fetchSynFromServer();
            preservedSynRaw = synRaw.value;
            if (snapshot.synEdits && hasSheetEdits(snapshot.synEdits)) {
              applySheetEdits(synRaw.value, snapshot.synEdits);
            }
            bumpSynSheetRevision();
            loadedFromLocal.value = true;
            saveStatus.value = 'saved';
            dirty.value = Math.max(dirty.value, localRev);
          } else if (remoteWon) {
            mergeLocalEditsOntoCurrent();
            dirty.value = Math.max(dirty.value, localRev);
          }
        } else if (snapshot && snapshot.bd && !remoteWon) {
          bdRaw.value = snapshot.bd;
          loadedFromLocal.value = true;
          saveStatus.value = 'saved';
        }

        if (!synRaw.value) {
          await fetchSynFromServer();
          preservedSynRaw = synRaw.value;
        }

        gridPreparing.value = true;

        const synBootPromise = wantSynthesis ? loadSynthesis() : null;

        if (!remoteOnly.value && !bdRaw.value && preBd && preBd.sheet) {
          await yieldToMain();
          bdSheet.value = preBd.sheet;
          bdSheetBuiltAt = 0;
          gridPreparing.value = false;
          syncBdCalcEngine();
          void loadSheetRaw('bd').then((raw) => {
            if (raw && !bdRaw.value) bdRaw.value = raw;
          });
        } else if (bdRaw.value) {
          bumpBdSheetRevision();
          await applyBdFromRaw(false);
          gridPreparing.value = false;
          syncBdCalcEngine();
        } else {
          await loadBdFromServer({ progressive: chunkedApi.value });
        }

        if (synBootPromise) {
          await synBootPromise;
        } else {
          void deferHeavy(() => startSynBackgroundPrepare());
        }

        if (
          (synRaw.value && !synRawLooksHealthy(synRaw.value)) ||
          (synthesisSheet.value && !synGridLooksHealthy(synthesisSheet.value))
        ) {
          await healSynRawIfCorrupted('post-boot');
          if (route.value === 'synthesis') {
            synthesisLoadPromise = null;
            await loadSynthesis();
          }
        }
      } catch (e) {
        error.value = e.message;
        gridPreparing.value = false;
        console.error('Startup failed:', e);
      }
    }

    /**
     * Force-commit a cell that is still being edited (focused input) so no
     * keystroke is lost when the user navigates, saves, or leaves the page.
     * In-cell edits otherwise only commit on blur/Enter; blurring the active
     * input runs that same commit path synchronously.
     */
    function flushActiveGridEdit() {
      if (typeof document === 'undefined') return;
      const el = document.activeElement;
      if (
        el &&
        el.tagName === 'INPUT' &&
        el.classList &&
        el.classList.contains('grid-cell-input') &&
        typeof el.blur === 'function'
      ) {
        el.blur();
      }
    }

    function onBeforeUnload() {
      flushActiveGridEdit();
      void autoSave.saveNow();
    }

    onUnmounted(() => {
      window.removeEventListener('beforeunload', onBeforeUnload);
      window.removeEventListener('pagehide', onBeforeUnload);
      window.removeEventListener('keydown', onGlobalKeydown, true);
      autoSave.destroy();
      session.destroy();
    });

    function syncSheetCell(gridSheet, row, col, value) {
      if (!gridSheet) return;
      const key = `${row}:${col}`;
      const map = gridSheet.cellMap instanceof Map ? gridSheet.cellMap : null;
      if (!map) return;
      let cell = map.get(key);
      if (!cell) {
        cell = { r: row, c: col, v: value, userEdited: true };
        gridSheet.cells.push(cell);
        map.set(key, cell);
      } else {
        cell.v = value;
        cell.userEdited = true;
        delete cell.f;
      }
    }

    function syncSheetCellPayload(gridSheet, row, col, payload) {
      if (!gridSheet) return;
      const map = gridSheet.cellMap instanceof Map ? gridSheet.cellMap : null;
      if (!map) return;
      const key = `${row}:${col}`;
      const hasContent = Boolean(payload && (payload.v != null || payload.f));
      let cell = map.get(key);
      if (!hasContent) {
        if (cell) {
          map.delete(key);
          gridSheet.cells = gridSheet.cells.filter((c) => c !== cell);
        }
        return;
      }
      if (!cell) {
        cell = { r: row, c: col };
        gridSheet.cells.push(cell);
        map.set(key, cell);
      }
      if (payload.v != null) cell.v = String(payload.v);
      else delete cell.v;
      if (payload.f) cell.f = String(payload.f);
      else delete cell.f;
      cell.userEdited = true;
    }

    function clearGridRowCells(gridSheet, row) {
      if (!gridSheet || !(gridSheet.cellMap instanceof Map)) return;
      const map = gridSheet.cellMap;
      const r = Number(row);
      const remove = [];
      for (const [key, cell] of map.entries()) {
        if (Number(cell.r) === r) remove.push(key);
      }
      for (const key of remove) {
        const cell = map.get(key);
        map.delete(key);
        gridSheet.cells = gridSheet.cells.filter((c) => c !== cell);
      }
    }

    function clearGridColCells(gridSheet, col) {
      if (!gridSheet || !(gridSheet.cellMap instanceof Map)) return;
      const map = gridSheet.cellMap;
      const c = String(col);
      const remove = [];
      for (const [key, cell] of map.entries()) {
        if (cell.c === c) remove.push(key);
      }
      for (const key of remove) {
        const cell = map.get(key);
        map.delete(key);
        gridSheet.cells = gridSheet.cells.filter((x) => x !== cell);
      }
    }

    const axisSelection = ref({ sheet: null, row: null, col: null });
    const axisClipboard = ref(null);

    function sheetRawAndGrid(sheetName) {
      if (sheetName === 'SYNTHESIS') {
        return { raw: synRaw.value, grid: synthesisSheet.value, engine: 'SYNTHESIS' };
      }
      if (sheetName === 'BD') {
        return { raw: bdRaw.value, grid: bdSheet.value, engine: 'BD' };
      }
      return { raw: null, grid: null, engine: null };
    }

    function onAxisSelect({ sheet, row, col }) {
      axisSelection.value = {
        sheet: sheet || null,
        row: row != null ? row : null,
        col: col != null ? col : null,
      };
    }

    function onAxisCopy() {
      const sel = axisSelection.value;
      if (!sel.sheet) return;
      const { raw } = sheetRawAndGrid(sel.sheet);
      if (!raw) return;
      if (sel.row != null) {
        axisClipboard.value = {
          sheet: sel.sheet,
          type: 'row',
          row: sel.row,
          col: null,
          cells: collectRowCells(raw, sel.row),
        };
      } else if (sel.col) {
        axisClipboard.value = {
          sheet: sel.sheet,
          type: 'col',
          row: null,
          col: sel.col,
          cells: collectColCells(raw, sel.col),
        };
      }
    }

    function snapshotAxisDeleteState(raw) {
      return {
        deletedRows: [...(raw.deletedRows || [])].map(Number).filter(Number.isFinite),
        deletedCols: [...(raw.deletedCols || [])].map(String).filter(Boolean),
        cells: [],
      };
    }

    function applyAxisPasteCells(sheetName, axis, target, cells) {
      const { raw, grid, engine } = sheetRawAndGrid(sheetName);
      if (!raw || !engine) return;
      if (axis === 'row') {
        pasteRowCells(raw, target, cells);
        clearGridRowCells(grid, target);
        for (const p of cells || []) {
          if (!p.c) continue;
          syncSheetCellPayload(grid, target, p.c, p);
          const v = p.v != null ? String(p.v) : '';
          void session
            .setCellValue(engine, target, p.c, v)
            .catch((e) => console.warn('Axis paste row:', e));
        }
      } else if (axis === 'col') {
        pasteColCells(raw, target, cells);
        clearGridColCells(grid, target);
        for (const p of cells || []) {
          if (p.r == null) continue;
          syncSheetCellPayload(grid, p.r, target, p);
          const v = p.v != null ? String(p.v) : '';
          void session
            .setCellValue(engine, p.r, target, v)
            .catch((e) => console.warn('Axis paste col:', e));
        }
      }
      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
    }

    async function onAxisPaste() {
      const clip = axisClipboard.value;
      const sel = axisSelection.value;
      if (!clip || !sel.sheet || clip.sheet !== sel.sheet) return;
      const { raw } = sheetRawAndGrid(sel.sheet);
      if (!raw) return;

      let before = [];
      let target = null;
      let axis = null;
      if (clip.type === 'row' && sel.row != null) {
        axis = 'row';
        target = sel.row;
        before = collectRowCells(raw, sel.row);
      } else if (clip.type === 'col' && sel.col) {
        axis = 'col';
        target = sel.col;
        before = collectColCells(raw, sel.col);
      } else {
        return;
      }

      if (!applyingHistory) {
        editHistory.push({
          type: 'axisPaste',
          sheet: sel.sheet,
          axis,
          target,
          before,
          after: clip.cells || [],
        });
        historyTick.value = editHistory.revision;
      }

      applyAxisPasteCells(sel.sheet, axis, target, clip.cells || []);
    }

    function applyUserGapsToGrid(raw, grid) {
      if (!raw || !grid) return;
      grid.userRowGaps = Array.isArray(raw.userRowGaps)
        ? raw.userRowGaps.map((g) => ({ ...g }))
        : [];
      grid.userColGaps = Array.isArray(raw.userColGaps)
        ? raw.userColGaps.map((g) => ({ ...g }))
        : [];
      grid.userGapCells = Array.isArray(raw.userGapCells)
        ? raw.userGapCells.map((c) => ({ ...c }))
        : [];
    }

    function applyUserGapsSnapshot(sheetName, snapshot) {
      const { raw, grid } = sheetRawAndGrid(sheetName);
      if (!raw || !snapshot) return;
      raw.userRowGaps = (snapshot.rowGaps || []).map((g) => ({ ...g }));
      raw.userColGaps = (snapshot.colGaps || []).map((g) => ({ ...g }));
      raw.userGapCells = (snapshot.gapCells || []).map((c) => ({ ...c }));
      if (grid) {
        applyUserGapsToGrid(raw, grid);
        if (sheetName === 'BD') {
          grid.bodyDisplayRows = computeBodyDisplayRows(grid);
        }
      }
      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
    }

    function applyAxisDeleteSnapshot(sheetName, snapshot) {
      const { raw, grid } = sheetRawAndGrid(sheetName);
      if (!raw || !snapshot) return;
      raw.deletedRows = [...(snapshot.deletedRows || [])];
      raw.deletedCols = [...(snapshot.deletedCols || [])];
      if (grid) {
        grid.deletedRows = raw.deletedRows;
        grid.deletedCols = raw.deletedCols;
      }
      if (sheetName === 'SYNTHESIS' && preservedSynRaw && preservedSynRaw !== raw) {
        preservedSynRaw.deletedRows = [...raw.deletedRows];
        preservedSynRaw.deletedCols = [...raw.deletedCols];
      }
      const engineSheet = sheetName === 'SYNTHESIS' ? 'SYNTHESIS' : 'BD';
      for (const cell of snapshot.cells || []) {
        const v = cell.v != null ? String(cell.v) : '';
        upsertRawCell(raw, cell.r, cell.c, v);
        if (grid) syncSheetCell(grid, cell.r, cell.c, v);
        void session
          .setCellValue(engineSheet, cell.r, cell.c, v)
          .catch((e) => console.warn('Axis delete restore:', e));
      }
      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
      if (sheetName === 'BD') {
        bumpBdSheetRevision();
        void applyBdFromRaw(false);
      } else if (sheetName === 'SYNTHESIS') {
        bumpSynSheetRevision();
        void applySynFromRaw(false);
      }
    }

    function onRowInsert({ sheet, afterExcelRow }) {
      if (sheet !== 'BD' && sheet !== 'SYNTHESIS') return;
      const row = Number(afterExcelRow);
      if (!Number.isFinite(row)) return;
      const { raw, grid } = sheetRawAndGrid(sheet);
      if (!raw) return;
      const before = cloneUserGapsSnapshot(raw);
      addUserRowGap(raw, row);
      if (grid) {
        applyUserGapsToGrid(raw, grid);
        if (sheet === 'BD') {
          grid.bodyDisplayRows = computeBodyDisplayRows(grid);
        }
      }
      if (!applyingHistory) {
        editHistory.push({
          type: 'userGaps',
          sheet,
          before,
          after: cloneUserGapsSnapshot(raw),
        });
        historyTick.value = editHistory.revision;
      }
      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
    }

    function onColumnInsert({ sheet, afterCol }) {
      if (sheet !== 'BD' && sheet !== 'SYNTHESIS') return;
      if (!afterCol || isUserGapCol(afterCol)) return;
      const { raw, grid } = sheetRawAndGrid(sheet);
      if (!raw) return;
      const before = cloneUserGapsSnapshot(raw);
      addUserColGap(raw, afterCol);
      if (grid) applyUserGapsToGrid(raw, grid);
      if (!applyingHistory) {
        editHistory.push({
          type: 'userGaps',
          sheet,
          before,
          after: cloneUserGapsSnapshot(raw),
        });
        historyTick.value = editHistory.revision;
      }
      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
    }

    function cloneRawSheet(sheet) {
      if (!sheet) return null;
      return JSON.parse(JSON.stringify(sheet));
    }

    async function applyMatrixSnapshot(bd, syn) {
      if (!bd) return;
      bdRaw.value = cloneRawSheet(bd);
      bumpBdSheetRevision();
      await applyBdFromRaw(true);
      if (syn) {
        synRaw.value = cloneRawSheet(syn);
        preservedSynRaw = synRaw.value;
        bumpSynSheetRevision();
        await applySynFromRaw(true);
      }
      externalEditTick.value += 1;
      if (engineStarted && bdSheet.value) {
        await session.loadSheets([
          { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
        ]);
        engineLoadedForSheet = bdSheet.value;
      }
      dirty.value += 1;
      autoSave.schedule();
      void autoSave.saveNow();
    }

    function applyHistoryCell({ sheet, row, col, value }) {
      const raw =
        sheet === 'SYNTHESIS'
          ? synRaw.value
          : sheet === 'BD'
            ? bdRaw.value
            : null;
      const gridSheet =
        sheet === 'SYNTHESIS' ? synthesisSheet.value : bdSheet.value;
      if (!raw || !gridSheet) return;

      upsertRawCell(raw, row, col, value);
      syncSheetCell(gridSheet, row, col, value);
      externalEditTick.value += 1;

      const engineSheet = sheet === 'SYNTHESIS' ? 'SYNTHESIS' : 'BD';
      void session
        .setCellValue(engineSheet, row, col, value)
        .catch((e) => console.warn('History setCellValue:', e));

      dirty.value += 1;
      autoSave.schedule();
    }

    async function performUndo() {
      flushActiveGridEdit();
      await nextTick();
      const entry = editHistory.undo();
      if (!entry) return;
      applyingHistory = true;
      try {
        if (entry.type === 'matrix') {
          await applyMatrixSnapshot(entry.bdBefore, entry.synBefore);
        } else if (entry.type === 'userGaps') {
          applyUserGapsSnapshot(entry.sheet, entry.before);
        } else if (entry.type === 'axisDelete') {
          applyAxisDeleteSnapshot(entry.sheet, entry.before);
        } else if (entry.type === 'axisPaste') {
          applyAxisPasteCells(entry.sheet, entry.axis, entry.target, entry.before);
        } else {
          applyHistoryCell({ ...entry, value: entry.oldValue });
        }
      } finally {
        applyingHistory = false;
        historyTick.value = editHistory.revision;
      }
    }

    async function performRedo() {
      flushActiveGridEdit();
      await nextTick();
      const entry = editHistory.redo();
      if (!entry) return;
      applyingHistory = true;
      try {
        if (entry.type === 'matrix') {
          await applyMatrixSnapshot(entry.bdAfter, entry.synAfter);
        } else if (entry.type === 'userGaps') {
          applyUserGapsSnapshot(entry.sheet, entry.after);
        } else if (entry.type === 'axisDelete') {
          applyAxisDeleteSnapshot(entry.sheet, entry.after);
        } else if (entry.type === 'axisPaste') {
          applyAxisPasteCells(entry.sheet, entry.axis, entry.target, entry.after);
        } else {
          applyHistoryCell({ ...entry, value: entry.newValue });
        }
      } finally {
        applyingHistory = false;
        historyTick.value = editHistory.revision;
      }
    }

    function isEditingGridInput() {
      const el = document.activeElement;
      return Boolean(
        el &&
          el.tagName === 'INPUT' &&
          el.classList &&
          el.classList.contains('grid-cell-input')
      );
    }

    function onGlobalKeydown(e) {
      if (!isGridPage.value) return;
      if (!(e.ctrlKey || e.metaKey)) return;
      if (isEditingGridInput()) return;
      const key = e.key.toLowerCase();
      if (key === 'z') {
        e.preventDefault();
        if (e.shiftKey) performRedo();
        else performUndo();
      } else if (key === 'y') {
        e.preventDefault();
        performRedo();
      } else if (key === 'c') {
        const sel = axisSelection.value;
        if (sel.sheet && (sel.row != null || sel.col)) {
          e.preventDefault();
          onAxisCopy();
        }
      } else if (key === 'v') {
        const sel = axisSelection.value;
        if (
          axisClipboard.value &&
          sel.sheet &&
          (sel.row != null || sel.col)
        ) {
          e.preventDefault();
          void onAxisPaste();
        }
      }
    }

    function applySynPatches(patches) {
      if (!patches || !patches.length) return;
      const raw = synRaw.value;
      const grid = synthesisSheet.value;
      if (!raw) return;
      for (const p of patches) {
        upsertRawCell(raw, p.r, p.c, p.v);
        if (grid) {
          const key = `${p.r}:${p.c}`;
          const map = grid.cellMap instanceof Map ? grid.cellMap : null;
          let cell = map ? map.get(key) : null;
          if (!cell) {
            cell = { r: p.r, c: p.c, v: p.v };
            grid.cells.push(cell);
            if (map) map.set(key, cell);
          } else if (!cell.userEdited) {
            cell.v = p.v;
            delete cell.mat;
          }
        }
      }
      externalEditTick.value += 1;
    }

    function onCellChange({ row, col, value, sheet, previousValue, live }) {
      if (!applyingHistory && !live) {
        editHistory.push({
          sheet,
          row,
          col,
          oldValue: previousValue != null ? previousValue : '',
          newValue: value,
        });
        historyTick.value = editHistory.revision;
      }
      if (sheet === 'SYNTHESIS' && !synRaw.value) {
        void ensureSynRaw()
          .then((raw) => {
            if (!raw) return;
            upsertRawCell(raw, row, col, value);
            if (synthesisSheet.value) syncSheetCell(synthesisSheet.value, row, col, value);
            externalEditTick.value += 1;
            dirty.value += 1;
            autoSave.schedule();
          })
          .catch((e) => console.warn('Syn raw load on edit:', e));
        return;
      }
      const raw =
        sheet === 'SYNTHESIS'
          ? synRaw.value
          : sheet === 'BD'
            ? bdRaw.value
            : null;
      if (raw) upsertRawCell(raw, row, col, value);
      if (sheet === 'BD') {
        if (bdSheet.value) {
          void session
            .setCellValue('BD', row, col, value)
            .then(() => {
              externalEditTick.value += 1;
            })
            .catch((e) => console.warn('BD → Syn live calc:', e));
        }
        if (serverCalc.value && chunkedApi.value) {
          void patchSheetCells('bd', [{ r: row, c: col, v: value }])
            .then((res) => applySynPatches(res.synPatches))
            .catch((e) => console.warn('Server Syn recalc:', e));
        }
      } else if (sheet === 'SYNTHESIS') {
        externalEditTick.value += 1;
      }
      dirty.value += 1;
      if (chunkedApi.value && (sheet === 'BD' || sheet === 'SYNTHESIS')) {
        const sheetId = sheet === 'BD' ? 'bd' : 'synthesis';
        void patchSheetCells(sheetId, [{ r: row, c: col, v: value }]).catch((e) =>
          console.warn('Remote cell PATCH:', e)
        );
      }
      // Debounced only: a full-sheet snapshot + server POST on every keystroke made
      // editing crawl. The 400ms debounce, navigation flush and unload handler all
      // call flush/saveNow, so nothing is lost — but typing stays instant. Per-cell
      // realtime sync still happens above via patchSheetCells (single-cell PATCH).
      autoSave.schedule();
    }

    /**
     * Derived Synthesis cells (e.g. green-row auto-totals) computed in the grid.
     * Persisted like a normal edit — real-time single-cell PATCH to Supabase plus
     * the debounced session snapshot — but kept out of the undo history since the
     * user never typed them.
     */
    function onDerivedChange(changes) {
      if (!Array.isArray(changes) || !changes.length) return;
      const raw = synRaw.value;
      const patch = [];
      for (const ch of changes) {
        if (!ch) continue;
        const row = Number(ch.row);
        const col = String(ch.col);
        const value = ch.value != null ? ch.value : '';
        if (!Number.isFinite(row)) continue;
        if (raw) upsertRawCell(raw, row, col, value);
        if (synthesisSheet.value) syncSheetCell(synthesisSheet.value, row, col, value);
        patch.push({ r: row, c: col, v: value });
      }
      if (!patch.length) return;
      dirty.value += 1;
      if (patch.some((p) => Number(p.r) === 16)) {
        externalEditTick.value += 1;
      }
      if (chunkedApi.value) {
        void patchSheetCells('synthesis', patch).catch((e) =>
          console.warn('Derived Synthesis PATCH:', e)
        );
      }
      autoSave.schedule();
    }

    /**
     * Soft-delete a grid column: hide it (persisted in raw.deletedCols) and clear
     * every cell it holds so it no longer feeds calculations.
     */
    function onColDelete({ sheet, col }) {
      const c = String(col || '');
      if (!c) return;
      const raw =
        sheet === 'SYNTHESIS'
          ? synRaw.value
          : sheet === 'BD'
            ? bdRaw.value
            : null;
      const gridSheet =
        sheet === 'SYNTHESIS' ? synthesisSheet.value : bdSheet.value;
      if (!raw) return;

      const before = snapshotAxisDeleteState(raw);
      before.cells = collectColCells(raw, c).map((p) => ({
        r: p.r,
        c,
        v: p.v != null ? String(p.v) : '',
        f: p.f,
      }));

      const set = new Set((raw.deletedCols || []).map(String).filter(Boolean));
      set.add(c);
      raw.deletedCols = [...set];
      if (gridSheet) gridSheet.deletedCols = raw.deletedCols;
      if (sheet === 'SYNTHESIS' && preservedSynRaw && preservedSynRaw !== raw) {
        const pset = new Set(
          (preservedSynRaw.deletedCols || []).map(String).filter(Boolean)
        );
        pset.add(c);
        preservedSynRaw.deletedCols = [...pset];
      }

      const rows = new Set();
      for (const cell of raw.cells || []) {
        if (cell.c === c && cell.v != null && String(cell.v) !== '') {
          rows.add(Number(cell.r));
        }
      }
      const hdr = raw.headerRows || {};
      for (const [rowKey, cols] of Object.entries(hdr)) {
        const cell = cols[c];
        if (cell && cell.v != null && String(cell.v) !== '') {
          rows.add(Number(rowKey));
        }
      }
      const engineSheet = sheet === 'SYNTHESIS' ? 'SYNTHESIS' : 'BD';
      for (const row of rows) {
        upsertRawCell(raw, row, c, '');
        if (gridSheet) syncSheetCell(gridSheet, row, c, '');
        void session
          .setCellValue(engineSheet, row, c, '')
          .catch((e) => console.warn('Col delete recalc:', e));
      }

      const after = snapshotAxisDeleteState(raw);
      after.cells = before.cells.map((cell) => ({ r: cell.r, c: cell.c, v: '' }));

      if (!applyingHistory) {
        editHistory.push({
          type: 'axisDelete',
          sheet,
          axis: 'col',
          id: c,
          before,
          after,
        });
        historyTick.value = editHistory.revision;
      }

      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
      if (sheet === 'BD') {
        bumpBdSheetRevision();
        void applyBdFromRaw(false);
      } else if (sheet === 'SYNTHESIS') {
        bumpSynSheetRevision();
        void applySynFromRaw(false);
      }
    }

    /**
     * Soft-delete a grid row: hide it (persisted in raw.deletedRows so it survives
     * reload) and clear every cell it holds so it no longer feeds calculations.
     */
    function onRowDelete({ sheet, excelRow }) {
      const row = Number(excelRow);
      if (!Number.isFinite(row)) return;
      const raw =
        sheet === 'SYNTHESIS'
          ? synRaw.value
          : sheet === 'BD'
            ? bdRaw.value
            : null;
      const gridSheet =
        sheet === 'SYNTHESIS' ? synthesisSheet.value : bdSheet.value;
      if (!raw) return;

      const before = snapshotAxisDeleteState(raw);
      before.cells = collectRowCells(raw, row).map((p) => ({
        r: row,
        c: p.c,
        v: p.v != null ? String(p.v) : '',
        f: p.f,
      }));

      const set = new Set(
        (raw.deletedRows || []).map(Number).filter(Number.isFinite)
      );
      set.add(row);
      raw.deletedRows = [...set];
      if (gridSheet) gridSheet.deletedRows = raw.deletedRows;
      if (sheet === 'SYNTHESIS' && preservedSynRaw && preservedSynRaw !== raw) {
        const pset = new Set(
          (preservedSynRaw.deletedRows || []).map(Number).filter(Number.isFinite)
        );
        pset.add(row);
        preservedSynRaw.deletedRows = [...pset];
      }

      const cols = new Set();
      for (const c of raw.cells || []) {
        if (Number(c.r) === row && c.v != null && String(c.v) !== '') {
          cols.add(c.c);
        }
      }
      const engineSheet = sheet === 'SYNTHESIS' ? 'SYNTHESIS' : 'BD';
      for (const col of cols) {
        upsertRawCell(raw, row, col, '');
        if (gridSheet) syncSheetCell(gridSheet, row, col, '');
        void session
          .setCellValue(engineSheet, row, col, '')
          .catch((e) => console.warn('Row delete recalc:', e));
      }

      const after = snapshotAxisDeleteState(raw);
      after.cells = before.cells.map((cell) => ({ r: cell.r, c: cell.c, v: '' }));

      if (!applyingHistory) {
        editHistory.push({
          type: 'axisDelete',
          sheet,
          axis: 'row',
          id: row,
          before,
          after,
        });
        historyTick.value = editHistory.revision;
      }

      externalEditTick.value += 1;
      dirty.value += 1;
      autoSave.schedule();
      if (sheet === 'BD') {
        bumpBdSheetRevision();
        void applyBdFromRaw(false);
      } else if (sheet === 'SYNTHESIS') {
        bumpSynSheetRevision();
        void applySynFromRaw(false);
      }
    }

    async function saveNow() {
      flushActiveGridEdit();
      await autoSave.saveNow();
    }

    const exporting = ref(false);

    /** Edited cells of the active sheet, shaped for the export endpoint. */
    function collectOverrides(raw) {
      if (!raw) return [];
      return extractSheetEdits(raw).cells.map(({ r, c, v }) => ({ r, c, v }));
    }

    async function exportToExcel() {
      if (exporting.value) return;
      const sheetId = isDatabase.value ? 'bd' : isSynthesis.value ? 'synthesis' : null;
      if (!sheetId) return;
      flushActiveGridEdit();
      const raw = sheetId === 'bd' ? bdRaw.value : synRaw.value;
      exporting.value = true;
      try {
        const blob = await exportSheetXlsx(sheetId, collectOverrides(raw));
        const filename = sheetId === 'bd' ? 'Database.xlsx' : 'Synthesis.xlsx';
        const href = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = href;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(href);
      } catch (e) {
        error.value =
          e && e.message === 'EXPORT_UNAVAILABLE'
            ? 'Export Excel indisponible — lancez run-supabase-server.bat (backend Python).'
            : `Export Excel échoué — ${(e && e.message) || e}`;
      } finally {
        exporting.value = false;
      }
    }

    async function resetFromServerFiles() {
      if (
        !confirm(
          'Effacer la copie locale et recharger les fichiers JSON du projet ? Les modifications non exportées ailleurs seront perdues.'
        )
      ) {
        return;
      }
      await clearLocalSnapshot();
      // Also drop the cached grid transforms (precomputed grid) so the reset truly
      // rebuilds from the project JSON instead of re-serving a stale cached grid.
      await Promise.all([clearSheetTransform('bd'), clearSheetTransform('syn')]);
      editHistory.clear();
      historyTick.value = editHistory.revision;
      if (session && session.liveBdEdited) session.liveBdEdited.value = false;
      structureRevision.value = 0;
      loadedFromLocal.value = false;
      engineStarted = false;
      engineLoadedForSheet = null;
      synthesisLoadPromise = null;
      synthesisPreparePromise = null;
      synthesisSheet.value = null;
      bdLoading.value = false;
      try {
        await fetchBdFromServer();
        bumpBdSheetRevision();
        await applyBdFromRaw(true);
        synRaw.value = null;
        synthesisSheet.value = null;
        synSheetBuiltAt = -1;
        if (route.value === 'synthesis') {
          await fetchSynFromServer();
          bumpSynSheetRevision();
          await applySynFromRaw(true);
        }
        await session.loadSheets([
          { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
        ]);
        engineStarted = true;
        engineLoadedForSheet = bdSheet.value;
        dirty.value = 0;
        saveStatus.value = 'saved';
      } catch (e) {
        error.value = e.message;
      } finally {
        bdLoading.value = false;
      }
    }

    function closeMenu() {
      menuOpen.value = false;
    }

    function toggleMenu() {
      menuOpen.value = !menuOpen.value;
    }

    function urlForRoute(id) {
      const u = new URL(window.location.href);
      u.searchParams.set('route', id);
      return u;
    }

    function applyRoute(id, { source = 'ui' } = {}) {
      // Commit any in-progress cell edit before leaving the current grid so the
      // value (and its dependent recalculations) is never lost on navigation.
      flushActiveGridEdit();
      closeMenu();
      closeSearch();
      if (id === route.value) return;
      route.value = id;
      if (typeof window !== 'undefined' && source !== 'pop') {
        const u = urlForRoute(id);
        window.history.pushState({ route: id }, '', u.toString());
      }
      requestAnimationFrame(() => {
        if (id === 'synthesis' && !synthesisSheet.value) {
          void loadSynthesis();
        }
        syncBdCalcEngine();
      });
    }

    function navigate(id) {
      applyRoute(id, { source: 'ui' });
    }

    function onPopState() {
      const id = new URLSearchParams(window.location.search).get('route');
      const target = id && NAV_ROUTES.some((n) => n.id === id) ? id : DEFAULT_ROUTE;
      applyRoute(target, { source: 'pop' });
    }

    function syncInitialRouteToUrl() {
      const u = urlForRoute(route.value);
      window.history.replaceState({ route: route.value }, '', u.toString());
    }

    function toggleOutline() {
      requestAnimationFrame(() => {
        outlineMode.value = (outlineMode.value + 1) % 3;
      });
    }

    const outlineTitle = computed(() => {
      if (outlineMode.value === 1) return 'Afficher uniquement les sections';
      if (outlineMode.value === 2) return 'Tout déplier';
      return 'Afficher sections et sous-sections';
    });

    function closeSearch() {
      searchOpen.value = false;
      searchQuery.value = '';
      searchStatus.value = '';
      gridSearchCmd.value = { q: '', step: 0, t: Date.now() };
    }

    function toggleSearch() {
      if (searchOpen.value) closeSearch();
      else {
        searchOpen.value = true;
        nextTick(() => {
          if (searchInputRef.value && typeof searchInputRef.value.focus === 'function') {
            searchInputRef.value.focus();
          }
        });
      }
    }

    function updateSearchStatus(result) {
      if (!searchQuery.value.trim()) {
        searchStatus.value = '';
        return;
      }
      if (result && result.searching) {
        searchStatus.value = 'Recherche…';
        return;
      }
      if (!result || !result.count) searchStatus.value = 'Aucun résultat';
      else searchStatus.value = `${result.index + 1} / ${result.count}`;
    }

    function dispatchGridSearch(step = 0) {
      const q = searchQuery.value.trim();
      gridSearchCmd.value = { q, step, t: Date.now() };
    }

    function runGridSearch(step = 0) {
      if (!searchQuery.value.trim()) {
        searchStatus.value = '';
        gridSearchCmd.value = { q: '', step: 0, t: Date.now() };
        return;
      }
      // Immediately reflect that a search is in progress for better UX.
      if (step === 0) searchStatus.value = 'Recherche…';
      dispatchGridSearch(step);
    }

    function onSearchNavigated(result) {
      if (!searchQuery.value.trim()) {
        searchStatus.value = '';
        return;
      }
      if (result && result.hidden && !searchHiddenRetry) {
        searchHiddenRetry = true;
        if (outlineOnly.value) {
          outlineMode.value = 0;
          nextTick(() => dispatchGridSearch(0));
          return;
        }
        searchHiddenRetry = false;
      } else {
        searchHiddenRetry = false;
      }
      updateSearchStatus(result);
    }

    function onSearchInput() {
      const q = searchQuery.value.trim();
      if (!q) {
        searchStatus.value = '';
        gridSearchCmd.value = { q: '', step: 0, t: Date.now() };
        return;
      }

      // Auto-search as you type (debounced) to avoid misleading "no results"
      // and to keep the UI responsive on large grids.
      searchStatus.value = 'Recherche…';
      if (searchDebounceTimer) clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(() => {
        searchDebounceTimer = 0;
        // step=0 => recompute matches + scroll to first hit (if any)
        runGridSearch(0);
      }, 180);
    }

    function onSearchKeydown(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (searchDebounceTimer) {
          clearTimeout(searchDebounceTimer);
          searchDebounceTimer = 0;
        }
        const step = e.shiftKey ? -1 : 1;
        runGridSearch(step);
      } else if (e.key === 'F3') {
        e.preventDefault();
        if (searchDebounceTimer) {
          clearTimeout(searchDebounceTimer);
          searchDebounceTimer = 0;
        }
        runGridSearch(e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        closeSearch();
      }
    }

    function onSearchRowHidden() {
      if (outlineOnly.value) {
        outlineMode.value = 0;
        nextTick(() => dispatchGridSearch(0));
      }
    }

    function deferClick(fn) {
      requestAnimationFrame(() => {
        void deferHeavy(fn);
      });
    }

    async function openMatrix() {
      if (!bdSheet.value) return;
      matrixPendingModel = null;
      matrixState.value = null;
      matrixOpen.value = true;
      try {
        if (!synRaw.value) {
          try {
            await ensureSynRaw();
          } catch {
            /* BD-only matrix still usable */
          }
        } else {
          void startSynBackgroundPrepare();
        }
        let synSheetForMatrix = synthesisSheet.value;
        if (!synSheetForMatrix && synRaw.value) {
          synSheetForMatrix = await transformSynthesisSheetAsync(synRaw.value);
        }
        matrixState.value = buildMatrixState(bdSheet.value, synSheetForMatrix);
        matrixSessionBefore = {
          bd: cloneRawSheet(bdRaw.value),
          syn: synRaw.value ? cloneRawSheet(synRaw.value) : null,
        };
        matrixSessionDirty = false;
      } catch (e) {
        error.value = (e && e.message) || String(e);
        console.error('Matrix open failed:', e);
        matrixOpen.value = false;
      }
    }

    function commitMatrixSessionHistory() {
      if (!matrixSessionBefore || applyingHistory || !matrixSessionDirty) {
        matrixSessionBefore = null;
        matrixSessionDirty = false;
        return;
      }
      editHistory.push({
        type: 'matrix',
        bdBefore: matrixSessionBefore.bd,
        synBefore: matrixSessionBefore.syn,
        bdAfter: cloneRawSheet(bdRaw.value),
        synAfter: synRaw.value ? cloneRawSheet(synRaw.value) : null,
      });
      historyTick.value = editHistory.revision;
      matrixSessionBefore = null;
      matrixSessionDirty = false;
    }

    async function applyMatrixFromModel(bdModel) {
      if (!bdRaw.value || !bdModel || !bdModel.sections || !bdModel.sections.length) return;
      if (bdModel.sections.length < 2) return;
      if (matrixSaving.value) {
        matrixApplyQueued = true;
        matrixPendingModel = bdModel;
        return;
      }
      matrixSaving.value = true;
      try {
        if (!synRaw.value) {
          try {
            const raw = await loadSheetRaw('synthesis');
            if (raw) synRaw.value = raw;
          } catch {
            /* apply BD structure only */
          }
        }
        const synBase = matrixState.value && matrixState.value.syn != null ? matrixState.value.syn : null;
        const synModel = synBase
          ? alignSynModelToBd(cloneStructure(synBase), bdModel)
          : null;
        const bdModelClone = cloneStructure(bdModel);
        const result = applyMatrixSave(
          bdRaw.value,
          synRaw.value,
          bdModelClone,
          synModel
        );
        bdRaw.value = result.bdRaw;
        bumpBdSheetRevision();
        await applyBdFromRaw(true);
        if (result.synRaw) {
          synRaw.value = result.synRaw;
          preservedSynRaw = result.synRaw;
          if (!synRawLooksHealthy(synRaw.value)) {
            await healSynRawIfCorrupted('matrix-apply');
          }
          bumpSynSheetRevision();
          await applySynFromRaw(true);
        }
        externalEditTick.value += 1;
        if (session && session.liveBdEdited) session.liveBdEdited.value = true;
        if (engineStarted && bdSheet.value) {
          await session.loadSheets([
            { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
          ]);
          engineLoadedForSheet = bdSheet.value;
        }
        if (bdSheet.value) {
          const synForMatrix =
            synthesisSheet.value ||
            (synRaw.value ? await transformSynthesisSheetAsync(synRaw.value) : null);
          matrixState.value = buildMatrixState(bdSheet.value, synForMatrix);
        } else {
          matrixState.value = {
            bd: cloneStructure(result.bdModel),
            syn: result.synModel
              ? cloneStructure(result.synModel)
              : matrixState.value && matrixState.value.syn != null
                ? matrixState.value.syn
                : null,
          };
        }
        matrixSessionDirty = true;
        structureRevision.value = Math.max(structureRevision.value, 1);
        dirty.value += 1;
        autoSave.schedule();
        void autoSave.saveNow();
      } catch (e) {
        error.value = (e && e.message) || String(e);
        console.error('Matrix apply failed:', e);
      } finally {
        matrixSaving.value = false;
        if (matrixApplyQueued && matrixPendingModel) {
          matrixApplyQueued = false;
          const next = matrixPendingModel;
          void applyMatrixFromModel(next);
        }
      }
    }

    function onMatrixChange({ bd }) {
      if (!bd) return;
      matrixPendingModel = bd;
      if (matrixApplyDebounceTimer) clearTimeout(matrixApplyDebounceTimer);
      matrixApplyDebounceTimer = setTimeout(() => {
        matrixApplyDebounceTimer = null;
        const pending = matrixPendingModel;
        if (pending && matrixOpen.value) void applyMatrixFromModel(pending);
      }, 400);
    }

    async function closeMatrix() {
      if (matrixApplyDebounceTimer) {
        clearTimeout(matrixApplyDebounceTimer);
        matrixApplyDebounceTimer = null;
      }
      const pending = matrixPendingModel;
      matrixPendingModel = null;
      if (pending) {
        await applyMatrixFromModel(pending);
      }
      commitMatrixSessionHistory();
      matrixOpen.value = false;
      matrixState.value = null;
    }

    // Options SP2 row 14 mirrors Synthesis row 16 — preload Syn + react to BD recalc.
    // MNS row 14 (Y…AM) uses the same live link.
    watch(isOptionsSp2, (on) => {
      if (on) void startSynBackgroundPrepare();
    });
    watch(isMns, (on) => {
      if (on) void startSynBackgroundPrepare();
    });
    watch(
      () => session.synCalcTick?.value,
      () => {
        externalEditTick.value += 1;
      }
    );
    watch(
      () => session.displayTick?.value,
      () => {
        if (isOptionsSp2.value || isMns.value) externalEditTick.value += 1;
      }
    );

    return {
      bdLoading,
      synthesisLoading,
      overlayMessage,
      showContentOverlay,
      error,
      bdSheet,
      synthesisSheet,
      isDatabase,
      isSynthesis,
      isMns,
      isOptionsSp2,
      isCdcOutput,
      isGridPage,
      sessionLoading,
      showLegacyTopbarActions,
      dirty,
      saveStatus,
      saveError,
      loadedFromLocal,
      saveNow,
      exporting,
      exportToExcel,
      resetFromServerFiles,
      restoreSynthesisFromProject,
      route,
      gridPreparing,
      dataLoadProgress,
      chunkedApi,
      menuOpen,
      toggleMenu,
      closeMenu,
      outlineOnly,
      outlineMode,
      outlineTitle,
      currentNav,
      session,
      onCellChange,
      onRowDelete,
      onDerivedChange,
      onAxisSelect,
      onRowInsert,
      onColumnInsert,
      axisClipboard,
      navigate,
      toggleOutline,
      searchOpen,
      searchQuery,
      searchInputRef,
      gridSearchCmd,
      searchStatus,
      toggleSearch,
      onSearchInput,
      onSearchKeydown,
      onSearchNavigated,
      onSearchRowHidden,
      deferClick,
      matrixOpen,
      matrixState,
      matrixSaving,
      openMatrix,
      closeMatrix,
      onMatrixChange,
      canUndo,
      canRedo,
      performUndo,
      performRedo,
      externalEditTick,
    };
  },
  template: `
    <div class="app-shell" :class="{ 'menu-open': menuOpen }">
      <AppSidebar
        :current="route"
        :open="menuOpen"
        @navigate="navigate"
        @close="closeMenu"
      />
      <header class="app-topbar">
        <div class="topbar-left">
          <button
            type="button"
            class="burger"
            aria-label="Open menu"
            @click="toggleMenu"
          >
            <span></span><span></span><span></span>
          </button>
          <span v-if="isDatabase" class="page-title">Database</span>
          <span v-else-if="isSynthesis" class="page-title">Synthesis</span>
          <span v-else class="page-title">{{ currentNav.label }}</span>
          <template v-if="isDatabase || isSynthesis">
            <div class="topbar-history" role="group" aria-label="Historique des modifications">
              <button
                type="button"
                class="icon-btn icon-btn-sm"
                :disabled="!canUndo"
                title="Annuler (Ctrl+Z)"
                aria-label="Annuler"
                @click.prevent="performUndo"
              >
                <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                  <path
                    fill="none"
                    stroke="currentColor"
                    stroke-width="1.35"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                    d="M6 4 2 8l4 4V9h3a3 3 0 1 1 0 6H6"
                  />
                </svg>
              </button>
              <button
                type="button"
                class="icon-btn icon-btn-sm"
                :disabled="!canRedo"
                title="Rétablir (Ctrl+Y)"
                aria-label="Rétablir"
                @click.prevent="performRedo"
              >
                <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                  <path
                    fill="none"
                    stroke="currentColor"
                    stroke-width="1.35"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                    d="M10 4l4 4-4 4V9H7a3 3 0 1 0 0 6h3"
                  />
                </svg>
              </button>
            </div>
            <button
              type="button"
              class="icon-btn icon-btn-sm"
              title="Bookmark Matrix — reorder sections and sub-sections"
              aria-label="Open Bookmark Matrix"
              @click="openMatrix"
            >
              <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                <rect x="1" y="1" width="6" height="6" fill="currentColor"/>
                <rect x="9" y="1" width="6" height="6" fill="currentColor"/>
                <rect x="1" y="9" width="6" height="6" fill="currentColor"/>
                <rect x="9" y="9" width="6" height="6" fill="currentColor"/>
              </svg>
            </button>
            <button
              type="button"
              class="icon-btn icon-btn-sm icon-btn-outline"
              :class="{ active: outlineOnly }"
              :title="outlineTitle"
              aria-label="Toggle outline view"
              @click="toggleOutline"
            >
              <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                <ellipse cx="8" cy="8" rx="6.5" ry="4" fill="none" stroke="currentColor" stroke-width="1.2"/>
                <circle cx="8" cy="8" r="2" fill="currentColor"/>
              </svg>
              <span v-if="outlineMode > 0" class="icon-btn-badge">{{ outlineMode }}</span>
            </button>
            <button
              type="button"
              class="icon-btn icon-btn-sm"
              :class="{ active: searchOpen }"
              title="Rechercher dans la grille"
              aria-label="Rechercher dans la grille"
              @click="toggleSearch"
            >
              <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                <circle cx="7" cy="7" r="4.25" fill="none" stroke="currentColor" stroke-width="1.25"/>
                <path
                  fill="none"
                  stroke="currentColor"
                  stroke-width="1.25"
                  stroke-linecap="round"
                  d="M10.2 10.2 14 14"
                />
              </svg>
            </button>
            <div v-if="searchOpen" class="topbar-search" role="search">
              <input
                ref="searchInputRef"
                v-model="searchQuery"
                type="text"
                class="topbar-search-input"
                placeholder="Texte puis Entree"
                aria-label="Rechercher dans la grille"
                autocomplete="off"
                spellcheck="false"
                @input="onSearchInput"
                @keydown="onSearchKeydown"
              />
              <span v-if="searchStatus" class="topbar-search-status">{{ searchStatus }}</span>
            </div>
          </template>
        </div>
        <div class="topbar-right">
          <span class="status" v-if="dataLoadProgress">
            Données {{ dataLoadProgress.sheetId === 'syn' ? 'Synthesis' : 'Database' }}
            {{ dataLoadProgress.pct }}%
          </span>
          <span class="status" v-else-if="isSynthesis && synthesisLoading && !synthesisSheet">Préparation Synthesis…</span>
          <span class="status error-text" v-else-if="error">{{ error }}</span>
          <span class="status" v-else-if="isGridPage && saveStatus === 'saving'">Enregistrement…</span>
          <span class="status" v-else-if="isGridPage && saveStatus === 'dirty'">Enregistrement automatique…</span>
          <span class="status error-text" v-else-if="isGridPage && saveStatus === 'error'">
            Erreur — {{ saveError }}
          </span>
          <template v-if="isGridPage && !sessionLoading">
            <button
              v-show="showLegacyTopbarActions"
              type="button"
              class="topbar-save-btn"
              @click="saveNow"
            >
              Enregistrer
            </button>
            <button
              type="button"
              class="topbar-save-btn topbar-save-btn-muted"
              title="Télécharger la feuille active au format Excel (.xlsx), mise en forme identique"
              :disabled="exporting"
              @click="exportToExcel"
            >
              {{ exporting ? 'Export…' : 'Export Excel' }}
            </button>
            <button
              v-show="showLegacyTopbarActions && isSynthesis"
              type="button"
              class="topbar-save-btn topbar-save-btn-muted"
              title="Recharger synthesis-sheet.json (bande ADAPTATION + sous-sections)"
              @click="restoreSynthesisFromProject"
            >
              Restaurer Synthesis
            </button>
            <button
              v-show="showLegacyTopbarActions"
              type="button"
              class="topbar-save-btn topbar-save-btn-muted"
              title="Recharger les JSON du projet (efface la copie locale)"
              @click="resetFromServerFiles"
            >
              Fichiers projet
            </button>
          </template>
        </div>
      </header>
      <div class="app-body">
        <main class="app-content">
          <div v-if="bdSheet" v-show="isDatabase" class="grid-route-pane">
            <BdGrid
              key="bd-grid"
              :sheet="bdSheet"
              sheet-name="BD"
              :session="session"
              :raw-bd="bdRaw"
              :outline-only="outlineOnly"
              :outline-mode="outlineMode"
              :pane-visible="isDatabase"
              :external-edit-tick="externalEditTick"
              :search-cmd="gridSearchCmd"
              :axis-clipboard="axisClipboard"
              @cell-change="onCellChange"
              @row-delete="onRowDelete"
              @column-delete="onColDelete"
              @row-insert="onRowInsert"
              @column-insert="onColumnInsert"
              @axis-select="onAxisSelect"
              @axis-copy="onAxisCopy"
              @axis-paste="onAxisPaste"
              @search-navigated="onSearchNavigated"
              @search-row-hidden="onSearchRowHidden"
            />
          </div>
          <div v-if="synthesisSheet" v-show="isSynthesis" class="grid-route-pane">
            <SynthesisGrid
              key="syn-grid"
              :sheet="synthesisSheet"
              :session="session"
              :raw-syn="synRaw"
              :outline-only="outlineOnly"
              :outline-mode="outlineMode"
              :pane-visible="isSynthesis"
              :external-edit-tick="externalEditTick"
              :search-cmd="gridSearchCmd"
              :axis-clipboard="axisClipboard"
              @cell-change="onCellChange"
              @row-delete="onRowDelete"
              @column-delete="onColDelete"
              @row-insert="onRowInsert"
              @column-insert="onColumnInsert"
              @axis-select="onAxisSelect"
              @axis-copy="onAxisCopy"
              @axis-paste="onAxisPaste"
              @derived-change="onDerivedChange"
              @search-navigated="onSearchNavigated"
              @search-row-hidden="onSearchRowHidden"
            />
          </div>
          <div
            v-if="gridPreparing && isDatabase && !bdSheet"
            class="loading-overlay loading-overlay-subtle"
          >
            Préparation de la grille…
          </div>
          <div v-show="isMns" class="grid-route-pane">
            <MnsGrid v-if="isMns" key="mns-grid" storage-key="mns-grid-cells-v3" />
          </div>
          <div v-show="isOptionsSp2" class="grid-route-pane">
            <MnsGrid v-if="isOptionsSp2" key="options-sp2-grid" storage-key="options-sp2-grid-cells-v2" />
          </div>
          <div v-show="isCdcOutput" class="grid-route-pane">
            <CdcOutputGrid v-if="isCdcOutput" key="cdc-output-grid" />
          </div>
          <div v-if="error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="showContentOverlay" class="loading-overlay">{{ overlayMessage }}</div>
          <EmptyPage
            v-else-if="!isDatabase && !isSynthesis && !isMns && !isOptionsSp2 && !isCdcOutput"
            :title="currentNav.label"
          />
        </main>
      </div>
      <MatrixModal
        :open="matrixOpen"
        :state="matrixState"
        :saving="matrixSaving"
        @close="closeMatrix"
        @change="onMatrixChange"
      />
    </div>
  `,
};

createApp(App).mount('#app');
