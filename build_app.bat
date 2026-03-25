@echo off
REM ============================================================
REM  MP^&L Hub  –  Desktop Application Build Script (Windows)
REM  Run from the project root:  build_app.bat
REM ============================================================

echo === MP^&L Hub Build Script ===
echo.

REM -- Check Python --
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found on PATH. Install Python 3.9+ and try again.
    pause
    exit /b 1
)

REM -- Install / upgrade build dependencies --
echo Installing build dependencies...
pip install --upgrade pyinstaller
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller.
    pause
    exit /b 1
)

REM -- Install app runtime dependencies --
echo Installing runtime dependencies...
pip install PyQt5 pandas numpy bcrypt
if errorlevel 1 (
    echo WARNING: Some dependencies may have failed. Continuing...
)

REM -- Clean previous build artifacts --
echo Cleaning previous build...
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

REM -- Run PyInstaller --
echo Building executable...
pyinstaller "MP&L_Hub.spec"
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo  Distributable folder:  dist\MPL_Hub\
echo  Share the entire MPL_Hub folder with your users.
echo  Users launch the app by running:  MPL_Hub.exe
echo ============================================================
echo.
pause
