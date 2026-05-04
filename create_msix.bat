@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

:: ============================================================
:: create_msix.bat
:: Δημιουργεί MSIX package για Microsoft Store
:: Απαιτεί: Windows 10/11 SDK (makeappx.exe, signtool.exe)
:: Τρέξε ΜΕΤΑ το build_executable.bat
:: ============================================================

echo.
echo ====================================================
echo   MySchool Checks - MSIX Package Creator
echo   Version: 1.0
echo ====================================================
echo.

:: --- Μεταβλητές ---
set APP_NAME=MySchoolChecksPlus
set APP_VERSION=1.0.0.0
set DISPLAY_NAME=MySchool Checks
set PUBLISHER=CN=MichalisKat, O=Education, C=GR
set PACKAGE_DIR=msix_package
set OUTPUT_MSIX=MySchoolChecksPlus-1.0.msix

:: --- Βήμα 1: Έλεγχος dist\MySchoolChecksPlus.exe ---
if not exist "dist\MySchoolChecksPlus.exe" (
    echo [ERROR] Δεν βρέθηκε dist\MySchoolChecksPlus.exe
    echo         Τρέξε πρώτα το build_executable.bat
    pause
    exit /b 1
)
echo [OK] Βρέθηκε dist\MySchoolChecksPlus.exe

:: --- Βήμα 2: Εύρεση makeappx.exe (Windows SDK) ---
set MAKEAPPX=
for %%D in (
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64"
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64"
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64"
    "C:\Program Files (x86)\Windows Kits\10\bin\x64"
) do (
    if exist "%%~D\makeappx.exe" (
        set MAKEAPPX=%%~D\makeappx.exe
        echo [OK] Βρέθηκε makeappx.exe: %%~D
        goto :found_sdk
    )
)

:sdk_not_found
echo [ERROR] Δεν βρέθηκε makeappx.exe
echo.
echo   Εγκατέστησε το Windows 10/11 SDK:
echo   https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
echo.
echo   Εναλλακτικά, χρησιμοποίησε το MSIX Packaging Tool:
echo   https://aka.ms/msixpackagingtool
pause
exit /b 1

:found_sdk

:: --- Βήμα 3: Καθαρισμός και δημιουργία φακέλου package ---
echo.
echo [STEP] Δημιουργία δομής MSIX package...
if exist "%PACKAGE_DIR%"   rmdir /s /q "%PACKAGE_DIR%"
if exist "%OUTPUT_MSIX%"   del /q "%OUTPUT_MSIX%"

mkdir "%PACKAGE_DIR%"
mkdir "%PACKAGE_DIR%\Assets"

:: --- Βήμα 4: Αντιγραφή αρχείων ---
echo [STEP] Αντιγραφή αρχείων...
copy "dist\MySchoolChecksPlus.exe"      "%PACKAGE_DIR%\MySchoolChecksPlus.exe" > nul

if exist "MySchoolChecksPlus\drivers"   xcopy /e /q /y "MySchoolChecksPlus\drivers"     "%PACKAGE_DIR%\drivers\" > nul
if exist "MySchoolChecksPlus\screenshots" xcopy /e /q /y "MySchoolChecksPlus\screenshots" "%PACKAGE_DIR%\screenshots\" > nul
if exist "MySchoolChecksPlus\startup.mp3" copy "MySchoolChecksPlus\startup.mp3"          "%PACKAGE_DIR%\startup.mp3" > nul
echo [OK] Αρχεία αντιγράφηκαν.

:: --- Βήμα 5: Δημιουργία placeholder assets (required by MSIX) ---
echo [STEP] Δημιουργία assets...
:: Αν δεν υπάρχουν PNG assets, δημιουργούμε placeholder text files
:: ΣΗΜΑΝΤΙΚΟ: Αντικατέστησε με πραγματικά PNG πριν την υποβολή στο Store
if not exist "%PACKAGE_DIR%\Assets\Square44x44Logo.png" (
    echo Placeholder - replace with real PNG 44x44 > "%PACKAGE_DIR%\Assets\Square44x44Logo.png"
)
if not exist "%PACKAGE_DIR%\Assets\Square150x150Logo.png" (
    echo Placeholder - replace with real PNG 150x150 > "%PACKAGE_DIR%\Assets\Square150x150Logo.png"
)
if not exist "%PACKAGE_DIR%\Assets\StoreLogo.png" (
    echo Placeholder - replace with real PNG 50x50 > "%PACKAGE_DIR%\Assets\StoreLogo.png"
)
if not exist "%PACKAGE_DIR%\Assets\Wide310x150Logo.png" (
    echo Placeholder - replace with real PNG 310x150 > "%PACKAGE_DIR%\Assets\Wide310x150Logo.png"
)
echo [OK] Assets folder έτοιμος.
echo [WARN] ΑΝΤΙΚΑΤΕΣΤΗΣΕ τα placeholder .png με πραγματικές εικόνες πριν το Store!

:: --- Βήμα 6: Δημιουργία AppxManifest.xml ---
echo.
echo [STEP] Δημιουργία AppxManifest.xml...

(
echo ^<?xml version="1.0" encoding="utf-8"?^>
echo ^<Package
echo   xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
echo   xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
echo   xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
echo   IgnorableNamespaces="uap rescap"^>
echo.
echo   ^<Identity
echo     Name="%APP_NAME%"
echo     Publisher="%PUBLISHER%"
echo     Version="%APP_VERSION%"
echo     ProcessorArchitecture="x64" /^>
echo.
echo   ^<Properties^>
echo     ^<DisplayName^>%DISPLAY_NAME%^</DisplayName^>
echo     ^<PublisherDisplayName^>Michalis Katsirintakis^</PublisherDisplayName^>
echo     ^<Description^>Αυτοματοποιημένοι έλεγχοι δεδομένων MySchool για εκπαιδευτικούς^</Description^>
echo     ^<Logo^>Assets\StoreLogo.png^</Logo^>
echo   ^</Properties^>
echo.
echo   ^<Dependencies^>
echo     ^<TargetDeviceFamily
echo       Name="Windows.Desktop"
echo       MinVersion="10.0.17763.0"
echo       MaxVersionTested="10.0.22621.0" /^>
echo   ^</Dependencies^>
echo.
echo   ^<Resources^>
echo     ^<Resource Language="el-GR" /^>
echo     ^<Resource Language="en-US" /^>
echo   ^</Resources^>
echo.
echo   ^<Applications^>
echo     ^<Application Id="MySchoolChecksPlus"
echo       Executable="MySchoolChecksPlus.exe"
echo       EntryPoint="Windows.FullTrustApplication"^>
echo       ^<uap:VisualElements
echo         DisplayName="%DISPLAY_NAME%"
echo         Description="Αυτοματοποιημένοι έλεγχοι δεδομένων MySchool"
echo         BackgroundColor="transparent"
echo         Square150x150Logo="Assets\Square150x150Logo.png"
echo         Square44x44Logo="Assets\Square44x44Logo.png"^>
echo         ^<uap:DefaultTile
echo           Wide310x150Logo="Assets\Wide310x150Logo.png"
echo           ShortName="%DISPLAY_NAME%" /^>
echo         ^<uap:SplashScreen Image="Assets\StoreLogo.png" /^>
echo       ^</uap:VisualElements^>
echo     ^</Application^>
echo   ^</Applications^>
echo.
echo   ^<Capabilities^>
echo     ^<rescap:Capability Name="runFullTrust" /^>
echo   ^</Capabilities^>
echo.
echo ^</Package^>
) > "%PACKAGE_DIR%\AppxManifest.xml"

echo [OK] AppxManifest.xml δημιουργήθηκε.

:: --- Βήμα 7: Δημιουργία MSIX package ---
echo.
echo [STEP] Δημιουργία MSIX package...
"%MAKEAPPX%" pack /d "%PACKAGE_DIR%" /p "%OUTPUT_MSIX%" /overwrite

if errorlevel 1 (
    echo.
    echo [ERROR] Αποτυχία δημιουργίας MSIX!
    echo         Ελέγξε το AppxManifest.xml για σφάλματα.
    echo.
    echo   Tip: Χρησιμοποίησε το MSIX Packaging Tool για γραφικό περιβάλλον:
    echo   https://aka.ms/msixpackagingtool
    pause
    exit /b 1
)

:: --- Βήμα 8: Επαλήθευση output ---
if exist "%OUTPUT_MSIX%" (
    echo.
    echo ====================================================
    echo   [SUCCESS] MSIX δημιουργήθηκε επιτυχώς!
    echo.
    echo   Output: %OUTPUT_MSIX%
    for %%A in ("%OUTPUT_MSIX%") do echo   Μέγεθος: %%~zA bytes
    echo.
    echo   ΕΠΟΜΕΝΑ ΒΗΜΑΤΑ για Microsoft Store:
    echo   1. Υπόγραψε το .msix με πιστοποιητικό (signtool)
    echo      ή χρησιμοποίησε Partner Center για αυτόματη υπογραφή
    echo   2. Πήγαινε στο Partner Center:
    echo      https://partner.microsoft.com/dashboard
    echo   3. Create new app -> Upload .msix
    echo   4. Αντικατάστησε τα placeholder assets με πραγματικά PNG
    echo ====================================================
) else (
    echo [ERROR] Το %OUTPUT_MSIX% δεν δημιουργήθηκε.
    pause
    exit /b 1
)

echo.
echo Πάτα οποιοδήποτε πλήκτρο για έξοδο...
pause > nul
exit /b 0
