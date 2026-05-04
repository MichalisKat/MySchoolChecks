@echo off
cd /d "%~dp0"

:: ── Shortcut Desktop (μόνο μία φορά) ─────────────────────────────────────
if not exist ".shortcut_ok" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $desk=[Environment]::GetFolderPath('Desktop'); $sc=$ws.CreateShortcut($desk+'\MySchool Checks.lnk'); $sc.TargetPath='%~f0'; $sc.WorkingDirectory='%~dp0'; $sc.IconLocation='%~dp0app.ico,0'; $sc.Description='MySchool Checks - D/nsh P.E. Anatolikhs Thessalonikhs'; $sc.Save()"
    echo. > .shortcut_ok
)

:: ── Εκκίνηση ελαχιστοποιημένη ώστε το terminal να μην φαίνεται ───────────
if "%HIDDEN%"=="" (
    set HIDDEN=1
    start /MIN "" cmd /c "%~f0"
    exit
)

:: ── Έλεγχος Python ────────────────────────────────────────────────────────
set PYTHON_CMD=
python --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=python

if "%PYTHON_CMD%"=="" (
    py --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=py
)

if "%PYTHON_CMD%"=="" (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('H Python den einai egkatesthmene.`nEpiske8teite: https://www.python.org/downloads/`n`nTsekare ''Add Python to PATH'' kata thn egkatastash.', 'MySchool - Sfalma')"
    powershell -Command "Start-Process 'https://www.python.org/downloads/'"
    exit /b 1
)

:: ── Εγκατάσταση βιβλιοθηκών (μόνο πρώτη φορά) ───────────────────────────
if not exist ".libs_ok" (
    %PYTHON_CMD% -m pip install -r requirements.txt --disable-pip-version-check >nul 2>&1
    echo. > .libs_ok
)

:: ── Εκκίνηση εφαρμογής ────────────────────────────────────────────────────
start "" pythonw main.py
