@echo off
REM ============================================================
REM  MP^&L Hub  –  Desktop Application Build Script (Windows)
REM  Run from the project root:  build_app.bat
REM ============================================================

REM Shared-drive path where updates are published for auto-update
set "UPDATES_DIR=W:\_US Operations\P&M Americas\21200 Supplier Quality & Logistics\21220 Material Planning & Logistic\12 MP&L Hub\Data\updates"

echo === MP^&L Hub Build Script ===
echo.

REM -- Read version from version.txt --
set /p APP_VERSION=<version.txt
echo Version: %APP_VERSION%
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
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller.
    pause
    exit /b 1
)

REM -- Install app runtime dependencies --
echo Installing runtime dependencies...
python -m pip install PyQt5 pandas numpy bcrypt matplotlib tqdm
if errorlevel 1 (
    echo WARNING: Some dependencies may have failed. Continuing...
)

REM -- Clean previous build artifacts --
echo Cleaning previous build...
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

REM -- Run PyInstaller --
echo Building executable...
python -m PyInstaller "MP&L_Hub.spec"
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!  Version: %APP_VERSION%
echo  Distributable folder:  dist\MPL_Hub\
echo ============================================================
echo.

REM -- Offer to push the update to the shared drive --
set /p PUSH=Push this build to the shared drive for auto-update? (y/n):
if /i not "%PUSH%"=="y" goto :done

echo.
echo Pushing to shared drive...

REM Create updates folder if it doesn't exist
if not exist "%UPDATES_DIR%" mkdir "%UPDATES_DIR%"

REM Copy the new build
robocopy "dist\MPL_Hub" "%UPDATES_DIR%\MPL_Hub" /E /IS /IT /COPYALL /R:3 /W:2
if errorlevel 8 (
    echo ERROR: robocopy failed. Check that the shared drive is accessible.
    pause
    exit /b 1
)

REM Write version file so clients know an update is available
echo %APP_VERSION%>"%UPDATES_DIR%\version.txt"

echo.
echo ============================================================
echo  Update published!
echo  Version %APP_VERSION% is now available to all users.
echo  They will be prompted on their next app launch.
echo ============================================================
echo.

:done
pause
