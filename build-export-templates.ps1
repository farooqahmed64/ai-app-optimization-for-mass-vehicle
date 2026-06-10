# Pre-build the Excel export templates from the master workbook.
# Run offline, one-time, and again whenever workbooks\base-de-donnees-complete-avec-liens.xlsm
# changes. Outputs api\export_templates\{bd,synthesis}-export-template.xlsx (commit them).
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$apiDir = Join-Path $root "api"
$python = Join-Path $apiDir ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { $python = "python" }

Push-Location $apiDir
try {
    $env:PYTHONPATH = $apiDir
    & $python -m tools.build_export_templates
} finally {
    Pop-Location
}
exit $LASTEXITCODE
