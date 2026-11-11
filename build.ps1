# Builds release\QuoteGrid\QuoteGrid.exe and release\QuoteGrid-v<version>-win64.zip
param([string]$Version = "1.0.0")
# pip and PyInstaller log to stderr, which Windows PowerShell 5.1 turns into
# errors under "Stop" -- so check $LASTEXITCODE instead.
$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot

python -m pip install -q --disable-pip-version-check -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }

# Absolute paths: PyInstaller resolves --icon/--add-data relative to --specpath
$root = $PSScriptRoot
# --noupx: UPX-packed DLLs are a common antivirus false-positive trigger
python -m PyInstaller --noconfirm --clean --onefile --windowed --noupx `
  --name QuoteGrid --icon "$root\assets\logo.ico" `
  --paths "$root\src" `
  --add-data "$root\assets\logo_128.png;assets" `
  --collect-all pdfplumber --collect-all pypdfium2 --collect-data docx `
  --hidden-import defusedxml --hidden-import defusedxml.ElementTree `
  --collect-all winrt --hidden-import ocr `
  --workpath build --specpath build --distpath dist `
  src\quotegrid_app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$out = "release\QuoteGrid"
New-Item -ItemType Directory -Force $out | Out-Null
Copy-Item dist\QuoteGrid.exe $out -Force
$zip = "release\QuoteGrid-v$Version-win64.zip"
if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path $out -DestinationPath $zip
Write-Host "Built $out\QuoteGrid.exe and $zip"
