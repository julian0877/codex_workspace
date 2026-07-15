$ErrorActionPreference = "Stop"
$env:QT_QPA_PLATFORM = "offscreen"

python -m pytest -v
if ($LASTEXITCODE -ne 0) { throw "Automated tests failed" }

python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name TenderWordFormatter `
    --paths src `
    --collect-all win32com `
    src/tender_formatter/main.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed" }

$executable = "dist/TenderWordFormatter/TenderWordFormatter.exe"
if (-not (Test-Path $executable)) { throw "Build artifact missing: $executable" }
Write-Host "Build complete: $executable"
