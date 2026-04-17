@echo off
REM ============================================================
REM  MP^&L Hub  –  Publish Update to Shared Drive
REM  Run this after build_app.bat when you're ready to release.
REM  Requires:  dist\MPL_Hub\  to exist from a previous build.
REM ============================================================

set "UPDATES_DIR=W:\_US Operations\P&M Americas\21200 Supplier Quality & Logistics\21220 Material Planning & Logistic\12 MP&L Hub\Data\updates"

echo === MP^&L Hub Publish Update ===
echo.

REM -- Read version --
set /p APP_VERSION=<version.txt
echo Publishing version: %APP_VERSION%
echo Destination: %UPDATES_DIR%
echo.

REM -- Confirm --
set /p CONFIRM=Are you sure you want to publish v%APP_VERSION% to all users? (y/n):
if /i not "%CONFIRM%"=="y" (
    echo Cancelled.
    pause
    exit /b 0
)

REM -- Check build exists --
if not exist "dist\MPL_Hub\MPL_Hub.exe" (
    echo ERROR: dist\MPL_Hub\MPL_Hub.exe not found.
    echo Run build_app.bat first, then come back here.
    pause
    exit /b 1
)

REM -- Check shared drive is reachable --
if not exist "%UPDATES_DIR:~0,3%" (
    echo ERROR: Shared drive not accessible. Make sure W: is mapped.
    pause
    exit /b 1
)

REM -- Create updates folder if needed --
if not exist "%UPDATES_DIR%" mkdir "%UPDATES_DIR%"

REM -- Copy build --
echo Copying build files...
robocopy "dist\MPL_Hub" "%UPDATES_DIR%\MPL_Hub" /E /IS /IT /COPYALL /R:3 /W:2
if errorlevel 8 (
    echo ERROR: robocopy failed. Check network connection and permissions.
    pause
    exit /b 1
)

REM -- Write version file (this is what triggers the update prompt for users) --
echo %APP_VERSION%>"%UPDATES_DIR%\version.txt"

echo.
echo ============================================================
echo  Done! Version %APP_VERSION% is now live.
echo  Users will be prompted to update on their next launch.
echo ============================================================
echo.
pause
