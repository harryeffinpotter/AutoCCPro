# Build script for Auto CapCut Pro Bypass using Nuitka
# Requires Python 3.12.x

Write-Host "Building Auto CapCut Pro Bypass with Nuitka..." -ForegroundColor Cyan

# Check Python version
$pythonVersion = python --version 2>&1
Write-Host "Using: $pythonVersion" -ForegroundColor Yellow

# Clean previous builds
if (Test-Path "app_gui.exe") {
    Write-Host "Removing old app_gui.exe..." -ForegroundColor Yellow
    Remove-Item "app_gui.exe" -Force
}
if (Test-Path "app_gui.dist") {
    Write-Host "Removing old app_gui.dist..." -ForegroundColor Yellow
    Remove-Item "app_gui.dist" -Recurse -Force
}
if (Test-Path "app_gui.build") {
    Write-Host "Removing old app_gui.build..." -ForegroundColor Yellow
    Remove-Item "app_gui.build" -Recurse -Force
}
if (Test-Path "app_gui.onefile-build") {
    Write-Host "Removing old app_gui.onefile-build..." -ForegroundColor Yellow
    Remove-Item "app_gui.onefile-build" -Recurse -Force
}

# Build with Nuitka
Write-Host "Starting Nuitka compilation..." -ForegroundColor Green
python -m nuitka `
    --onefile `
    --windows-console-mode=force `
    --enable-plugin=tk-inter `
    --include-package=comtypes `
    --include-package=win32com `
    --include-package=pywinauto `
    --include-package=psutil `
    --include-package=pyperclip `
    --include-data-file=icon.ico=icon.ico `
    --windows-icon-from-ico=icon.ico `
    --company-name="YouStayGold" `
    --product-name="Auto Capcut Pro Bypass" `
    --file-version=2.0.0.0 `
    --product-version=2.0.0.0 `
    --file-description="Auto Capcut Pro Bypass" `
    app_gui.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nBuild successful! Executable: app_gui.exe" -ForegroundColor Green
    Write-Host "File size: $((Get-Item app_gui.exe).Length / 1MB) MB" -ForegroundColor Cyan
} else {
    Write-Host "`nBuild failed with exit code: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}
