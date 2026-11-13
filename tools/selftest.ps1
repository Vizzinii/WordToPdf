# Regression self-test for WordToPdfTool.
#
#   .\tools\selftest.ps1 -Sample A.docx                 # run from source
#   .\tools\selftest.ps1 -Exe dist\WordToPdfTool.exe -Sample A.docx
#   .\tools\selftest.ps1 -Sample .\docs\                 # whole folder
#   .\tools\selftest.ps1 -CheckOnly                      # environment only
#
# What it verifies, per document:
#   * a final PDF was produced
#   * page count matches the export
#   * the interim PDF was disposed of (or explicitly kept)
#   * the fidelity counters report zero pixel differences
# Results are read back from the tool's own structured log.
[CmdletBinding()]
param(
    [string]$Sample,
    [string]$Exe,
    [ValidateSet("wps", "word", "cli")]
    [string]$Engine = "word",
    [switch]$KeepInterim,
    [switch]$CheckOnly,
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$log = Join-Path $env:TEMP "pdf_tool_selftest.log"

function Invoke-Tool {
    param([string[]]$ToolArgs)
    if ($Exe) {
        # A --windowed PyInstaller build detaches from the console, so `& $Exe`
        # returns before the work finishes. Start it and wait for the process.
        $proc = Start-Process -FilePath $Exe -ArgumentList $ToolArgs `
            -PassThru -Wait -NoNewWindow:$false
        $script:ToolExit = $proc.ExitCode
    } else {
        & $Python (Join-Path $root "src\pdf_tool.py") @ToolArgs
        $script:ToolExit = $LASTEXITCODE
    }
}

if ($CheckOnly) {
    Write-Host "Environment check ($Engine)..." -ForegroundColor Cyan
    & $Python (Join-Path $root "src\wps_to_pdf.py") --engine $Engine --check
    exit $LASTEXITCODE
}

if (-not $Sample) { throw "Specify -Sample <file.docx | folder>" }

Remove-Item $log -Force -ErrorAction SilentlyContinue
$toolArgs = @("--selftest", (Resolve-Path $Sample).Path, $Engine)
if ($KeepInterim) { $toolArgs += "keep" }

Write-Host "Running: $($toolArgs -join ' ')" -ForegroundColor Cyan
Invoke-Tool -ToolArgs $toolArgs
$code = $script:ToolExit

if (-not (Test-Path $log)) {
    throw "No selftest log at $log (tool exit code $code)"
}
$lines = Get-Content $log -Encoding UTF8
$lines | Write-Host

# --- assertions -----------------------------------------------------------
# NOTE: `$array -notmatch 'x'` returns the elements that do NOT match, which is
# truthy whenever any line differs - so membership must be tested explicitly.
$failures = @()
if (-not ($lines -contains 'RESULT: OK')) { $failures += "RESULT is not OK" }

$reports = ($lines | Select-String '^\[\d+\] out = ').Count
if ($reports -eq 0) { $failures += "no per-document report found" }

# Documents without pictures skip optimisation entirely: the exported PDF *is*
# the deliverable, so `pdf_removed = False` is expected for them.
$skippedDocs = @()
foreach ($line in $lines) {
    if ($line -match '^\[(\d+)\] skipped = True$') { $skippedDocs += $matches[1] }
}

foreach ($line in $lines) {
    if ($line -match '^\[(\d+)\] diff = (\d+)$' -and [int]$matches[2] -ne 0) {
        $failures += "document $($matches[1]) reports $($matches[2]) pixel differences"
    }
    if ($line -match '^\[(\d+)\] final_exists = False$') {
        $failures += "document $($matches[1]) produced no final PDF"
    }
    if (-not $KeepInterim -and $line -match '^\[(\d+)\] interim_exists_after = True$') {
        $failures += "document $($matches[1]) left the interim PDF behind"
    }
    if ($line -match '^\[(\d+)\] pages = (\d+)$' -and [int]$matches[2] -le 0) {
        $failures += "document $($matches[1]) exported zero pages"
    }
    if ($line -match '^\[(\d+)\] pdf_removed = False$' -and
        $skippedDocs -notcontains $matches[1] -and -not $KeepInterim) {
        $failures += "document $($matches[1]) kept the interim PDF without being asked to"
    }
}

Write-Host ""
if ($failures.Count -eq 0) {
    Write-Host "SELFTEST PASSED ($reports document(s))" -ForegroundColor Green
    exit 0
}
Write-Host "SELFTEST FAILED:" -ForegroundColor Red
$failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
exit 1
