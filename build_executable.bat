@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo.
echo ====================================================
echo   MySchool Checks - PyInstaller Build
echo ====================================================
echo.

if not exist "MySchoolChecks\main.py" (
    echo [ERROR] MySchoolChecks\main.py not found.
    echo         Run from the myschool-checks\ folder.
    pause
    exit /b 1
)
echo [OK] MySchoolChecks\main.py found.

:: Backup settings - check both locations
echo.
echo [STEP] Backing up local_settings.json ...
set SETTINGS_BACKUP=%TEMP%\myschool_settings_backup.json

:: Priority 1: LOCALAPPDATA (installed version)
set LOCALAPPDATA_SETTINGS=%LOCALAPPDATA%\MySchoolChecks\data\local_settings.json
:: Priority 2: dist\data (portable/dev version)
set DIST_SETTINGS=dist\data\local_settings.json

if exist "%LOCALAPPDATA_SETTINGS%" (
    copy /y "%LOCALAPPDATA_SETTINGS%" "%SETTINGS_BACKUP%" > nul
    echo [OK] Settings backed up from LOCALAPPDATA.
) else if exist "%DIST_SETTINGS%" (
    copy /y "%DIST_SETTINGS%" "%SETTINGS_BACKUP%" > nul
    echo [OK] Settings backed up from dist\data.
) else (
    echo [INFO] No existing settings to backup.
    set SETTINGS_BACKUP=
)

:: Clean old builds
echo.
echo [STEP] Cleaning old builds...
if exist "build\"  rd /s /q "build"
if exist "dist\"   rd /s /q "dist"
echo [OK] Clean done.

:: Find PyInstaller
set PYINSTALLER=
for %%P in (
    "%APPDATA%\Python\Python314\Scripts\pyinstaller.exe"
    "C:\Program Files\Python314\Scripts\pyinstaller.exe"
    "%LOCALAPPDATA%\Python\pythoncore-3.14-64\Scripts\pyinstaller.exe"
    "%LOCALAPPDATA%\Programs\Python\Python314\Scripts\pyinstaller.exe"
    "%LOCALAPPDATA%\Programs\Python\Python313\Scripts\pyinstaller.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\Scripts\pyinstaller.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\Scripts\pyinstaller.exe"
) do (
    if exist %%P if not defined PYINSTALLER set PYINSTALLER=%%P
)
if not defined PYINSTALLER (
    where pyinstaller > nul 2>&1
    if not errorlevel 1 set PYINSTALLER=pyinstaller
)
if not defined PYINSTALLER (
    echo [ERROR] PyInstaller not found. Run: pip install pyinstaller
    pause
    exit /b 1
)
echo [OK] PyInstaller: %PYINSTALLER%

:: Run PyInstaller with spec
echo.
echo [STEP] Running PyInstaller (2-5 min)...
echo.

%PYINSTALLER% MySchoolChecks.spec

if errorlevel 1 (
    echo [ERROR] PyInstaller build failed!
    goto :restore
)
if not exist "dist\MySchoolChecks.exe" (
    echo [ERROR] dist\MySchoolChecks.exe not created.
    goto :restore
)

echo.
echo ====================================================
echo   [SUCCESS] Build complete!
for %%A in ("dist\MySchoolChecks.exe") do echo   Size: %%~zA bytes
echo ====================================================
echo.
echo [STEP] Cleaning build/ folder...
if exist "build\" rd /s /q "build"

:restore
:: Restore settings to BOTH locations
echo.
echo [STEP] Restoring settings...
if defined SETTINGS_BACKUP (
    if exist "%SETTINGS_BACKUP%" (
        :: Restore to dist\data (portable)
        if exist "dist\" (
            if not exist "dist\data\" mkdir "dist\data"
            copy /y "%SETTINGS_BACKUP%" "dist\data\local_settings.json" > nul
            echo [OK] Settings restored to dist\data\
        )
        :: Restore to LOCALAPPDATA (installed version)
        if not exist "%LOCALAPPDATA%\MySchoolChecks\data\" mkdir "%LOCALAPPDATA%\MySchoolChecks\data"
        copy /y "%SETTINGS_BACKUP%" "%LOCALAPPDATA_SETTINGS%" > nul
        echo [OK] Settings restored to LOCALAPPDATA\MySchoolChecks\data\
        del /q "%SETTINGS_BACKUP%"
    )
) else (
    if exist "dist\" (
        if not exist "dist\data\" mkdir "dist\data"
        echo [OK] dist\data\ created.
    )
)

echo.
echo Press any key to exit...
pause > nul
exit /b 0
