@echo off
chcp 65001 > nul

echo.
echo ====================================================
echo   MySchool Checks - NSIS Installer Compiler
echo ====================================================
echo.

if not exist "dist\MySchoolChecks.exe" (
    echo [ERROR] dist\MySchoolChecks.exe not found.
    echo         Run build_executable.bat first.
    pause
    exit /b 1
)
echo [OK] dist\MySchoolChecks.exe found.

set MAKENSIS=
for %%P in (
    "C:\Program Files (x86)\NSIS\makensis.exe"
    "C:\Program Files\NSIS\makensis.exe"
) do (
    if exist %%P if not defined MAKENSIS set MAKENSIS=%%P
)

if not defined MAKENSIS (
    echo [ERROR] NSIS not found.
    echo         Download from: https://nsis.sourceforge.io/Download
    pause
    exit /b 1
)
echo [OK] NSIS: %MAKENSIS%

:: Version from myschool-checks.nsi (!define APP_VERSION "x.y.z")
set APP_VER=
for /f "tokens=3" %%V in ('findstr /c:"!define APP_VERSION" myschool-checks.nsi') do set APP_VER=%%~V
set SETUP_FILE=myschool-checks-%APP_VER%-setup.exe
echo [OK] Version: %APP_VER%

echo.
echo [STEP] Compiling installer...
echo.

%MAKENSIS% myschool-checks.nsi

if errorlevel 1 (
    echo.
    echo [ERROR] Compilation failed. Check messages above.
    pause
    exit /b 1
)

if exist "%SETUP_FILE%" (
    echo.
    echo ====================================================
    echo   [SUCCESS] Installer created!
    echo.
    for %%A in ("%SETUP_FILE%") do echo   File: %SETUP_FILE%  (%%~zA bytes^)
    echo.
    echo   Ready to distribute!
    echo ====================================================
) else (
    echo [ERROR] Setup file not created.
    pause
    exit /b 1
)

echo.
echo Press any key to exit...
pause > nul
exit /b 0
