# Rebuilds the workbook, recalculates it in Excel (so cached values are stored for previewers),
# and verifies all three scenarios against the independent Python model. Requires Windows + desktop Excel.
#
#   powershell -File src\finalize.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$xlsx = Join-Path $root "Northlight_Scenario_Budget_Model.xlsx"
python (Join-Path $root "src\build_workbook.py")

# 1) verify every scenario on the freshly built file
foreach ($sc in "Base", "Upside", "Downside") {
    $dump = Join-Path $env:TEMP "northlight_dump_$sc.json"
    powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "src\recalc_and_dump.ps1") -Path $xlsx -Dump $dump -Scenario $sc | Out-Null
    Write-Host "--- $sc"
    python (Join-Path $root "src\verify.py") $dump $sc
    if ($LASTEXITCODE -ne 0) { throw "verification failed for $sc" }
}

# 2) save the recalculated Base-case file with cached values
$tmp = Join-Path $env:TEMP "northlight_baked.xlsx"
if (Test-Path $tmp) { Remove-Item $tmp -Force }
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "src\recalc_and_dump.ps1") -Path $xlsx -Dump (Join-Path $env:TEMP "northlight_dump_final.json") -Scenario Base -SaveAs $tmp | Out-Null
Copy-Item $tmp $xlsx -Force
Write-Host "Saved recalculated workbook: $xlsx"
