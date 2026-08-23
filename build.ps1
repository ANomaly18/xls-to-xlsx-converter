$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $Root

# Step 1: Validate the GUI build environment. PyInstaller can otherwise finish
# successfully while silently excluding Tk from the executable.
Write-Host ">>> Checking GUI build dependencies..."
python -c "import tkinter; import PIL; import PyInstaller; print(f'Tk {tkinter.TkVersion} available')"
if ($LASTEXITCODE -ne 0) {
    throw "GUI build dependencies are unavailable. Activate a Python environment with Tk and install .[build]."
}

# Step 2: Generate icon
Write-Host ">>> Generating icon..."
python .\tools\create_icon.py
if ($LASTEXITCODE -ne 0) {
    throw "Icon generation failed."
}

# Step 3: Clean previous build
$OutputExe = Join-Path $Root "dist\XLS to XLSX Converter.exe"
if (Test-Path -LiteralPath $OutputExe) {
    Remove-Item -LiteralPath $OutputExe -Force
}

# Step 4: Build with PyInstaller
Write-Host ">>> Building EXE with PyInstaller..."
python -m PyInstaller --noconfirm --clean .\XlsToXlsxConverter.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

Write-Host ""
Write-Host ">>> Build selesai:"
Write-Host $OutputExe
Write-Host ""

# Show file size
$Size = (Get-Item -LiteralPath $OutputExe).Length / 1MB
Write-Host ("Ukuran: {0:N1} MB" -f $Size)
