@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\bin\python.exe"
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python executable not found: "%PYTHON_EXE%"
    echo [INFO] Create the .venv first, then install requirements.
    exit /b 1
)

set "UPX_EXE="
set "UPX_DIR="

for /f "delims=" %%I in ('where upx 2^>nul') do (
    set "UPX_EXE=%%~fI"
    goto :upx_found
)

for /f "delims=" %%I in ('dir /b /s "%LOCALAPPDATA%\Microsoft\WinGet\Packages\UPX.UPX_*\upx-*\upx.exe" 2^>nul') do (
    set "UPX_EXE=%%~fI"
    goto :upx_found
)

goto :build

:upx_found
for %%D in ("%UPX_EXE%") do set "UPX_DIR=%%~dpD"
if "%UPX_DIR:~-1%"=="\" set "UPX_DIR=%UPX_DIR:~0,-1%"
echo [INFO] Using UPX: "%UPX_EXE%"

:build
echo [INFO] Cleaning old build artifacts...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo [INFO] Building EXE with PyInstaller...
if defined UPX_DIR (
    "%PYTHON_EXE%" -m PyInstaller --clean --noconfirm --upx-dir "%UPX_DIR%" keykey.spec
) else (
    echo [WARN] UPX not found. Building without UPX compression.
    "%PYTHON_EXE%" -m PyInstaller --clean --noconfirm keykey.spec
)

if errorlevel 1 (
    echo [ERROR] Build failed.
    exit /b 1
)

if exist "dist\keykey.exe" (
    for %%F in ("dist\keykey.exe") do echo [OK] Build complete. EXE size: %%~zF bytes
) else (
    echo [WARN] Build finished but dist\keykey.exe was not found.
)

exit /b 0
