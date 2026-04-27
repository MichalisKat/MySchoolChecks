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

if exist "myschool-checks-0.9.7-setup.exe" (
    echo.
    echo ====================================================
    echo   [SUCCESS] Installer created!
    echo.
    for %%A in ("myschool-checks-0.9.7-setup.exe") do echo   File: myschool-checks-0.9.7-setup.exe  (%%~zA bytes^)
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
