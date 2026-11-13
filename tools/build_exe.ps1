# Build a single-file WordToPdfTool.exe with PyInstaller.
#
#   .\tools\build_exe.ps1                 # default: onefile, windowed
#   .\tools\build_exe.ps1 -Console        # keep a console window (debugging)
#   .\tools\build_exe.ps1 -OneDir         # folder build: faster startup
#
# Requirements:  python -m pip install pyinstaller
# The entry point is src\pdf_tool.py; wps_to_pdf.py and optimize_pdf_images.py
# are imported by it and must be collected alongside.
[CmdletBinding()]
param(
    [switch]$Console,
    [switch]$OneDir,
    [string]$Python = "python",
    [string]$OutputDir = "dist"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    & $Python -c "import PyInstaller, sys; print('PyInstaller', PyInstaller.__version__)"
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller not installed. Run: $Python -m pip install pyinstaller"
    }

    $mode = if ($OneDir) { "--onedir" } else { "--onefile" }
    $windowed = if ($Console) { "--console" } else { "--windowed" }

    $args = @(
        "-m", "PyInstaller",
        $mode,
        $windowed,
        "--name", "WordToPdfTool",
        "--distpath", $OutputDir,
        "--workpath", "build",
        "--specpath", "build",
        "--noconfirm",
        "--clean",
        # Only these three modules; the tool does not use any other project code.
        "--paths", "src",
        # numpy arrives solely through PIL.ImageFilter.Color3DLUT, which this
        # tool never calls; excluding it keeps the payload ~26 MB smaller.
        "--exclude-module", "numpy",
        "--exclude-module", "PIL.ImageFilter",
        "--exclude-module", "tkinter",
        "--exclude-module", "unittest",
        "--exclude-module", "pydoc",
        "src\pdf_tool.py"
    )

    Write-Host "Building: $mode $windowed" -ForegroundColor Cyan
    & $Python @args
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

    $exe = Join-Path $OutputDir "WordToPdfTool.exe"
    if (Test-Path $exe) {
        $mb = [math]::Round((Get-Item $exe).Length / 1MB, 1)
        Write-Host "Built $exe ($mb MB)" -ForegroundColor Green
        Write-Host "Next: .\tools\selftest.ps1 -Exe '$exe' -Sample '<A.docx>'" -ForegroundColor Yellow
    } else {
        throw "Build finished but $exe was not produced"
    }
}
finally {
    Pop-Location
}
