; ============================================================
; myschool-checks.nsi
; NSIS Installer Script - MySchool Checks 1.0.0
; Compile: makensis myschool-checks.nsi
; Output:  myschool-checks-1.0.0-setup.exe
; ============================================================

Unicode True
SetCompressor /SOLID lzma
SetCompressorDictSize 64

; --- Metadata ---
!define APP_NAME      "MySchool Checks"
!define APP_VERSION   "3.1.0"
!define APP_PUBLISHER "Michalis Katsirintakis"
!define APP_URL       "https://github.com/mkatsirntakis/myschool-checks"
!define APP_EXE       "MySchoolChecks.exe"
!define APP_ICON      "MySchoolChecks\app.ico"
!define INSTALL_DIR   "$PROGRAMFILES64\MySchoolChecks"
!define REG_KEY       "Software\Microsoft\Windows\CurrentVersion\Uninstall\MySchoolChecks"
!define SETUP_EXE     "myschool-checks-${APP_VERSION}-setup.exe"

; --- NSIS Modern UI 2 ---
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "x64.nsh"

; --- Installer Info ---
Name             "${APP_NAME} ${APP_VERSION}"
OutFile          "${SETUP_EXE}"
InstallDir       "${INSTALL_DIR}"
InstallDirRegKey HKLM "${REG_KEY}" "InstallLocation"
RequestExecutionLevel admin
BrandingText     "${APP_PUBLISHER}"

; --- MUI Settings ---
!define MUI_ABORTWARNING
!define MUI_ICON   "${APP_ICON}"
!define MUI_UNICON "${APP_ICON}"

; --- Pages ---
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_RUN          "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT     "Launch MySchool Checks"
!define MUI_FINISHPAGE_LINK         "GitHub"
!define MUI_FINISHPAGE_LINK_LOCATION "${APP_URL}"
!insertmacro MUI_PAGE_FINISH

; --- Uninstaller Pages ---
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; --- Languages ---
!insertmacro MUI_LANGUAGE "Greek"
!insertmacro MUI_LANGUAGE "English"

; ============================================================
; SECTION: Main Install
; ============================================================
Section "MySchool Checks" SecMain
    SectionIn RO

    ; Απεγκατάσταση παλιάς έκδοσης MySchool Checks (αν υπάρχει)
    ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\MySchoolChecks" "UninstallString"
    StrCmp $0 "" skip_old_uninstall
        ExecWait '"$0" /S _?=$INSTDIR'
    skip_old_uninstall:

    SetOutPath "$INSTDIR"
    SetOverwrite on

    ; Main executable
    File "dist\${APP_EXE}"

    ; PDF guide (if exists)
    File /nonfatal "MySchoolChecks_Odigos.pdf"

    ; Icon
    File /nonfatal "${APP_ICON}"

    ; Startup sound
    File /nonfatal "MySchoolChecks\startup.mp3"

    ; Drivers (ChromeDriver fallback)
    SetOutPath "$INSTDIR\drivers"
    File /nonfatal /r "MySchoolChecks\drivers\*.*"

    ; Screenshots (help images)
    SetOutPath "$INSTDIR\screenshots"
    File /nonfatal /r "MySchoolChecks\screenshots\*.*"

    ; Create data folder (for settings)
    SetOutPath "$INSTDIR\data"

    ; Back to root
    SetOutPath "$INSTDIR"

    ; --- Registry: Add/Remove Programs ---
    WriteRegStr   HKLM "${REG_KEY}" "DisplayName"     "${APP_NAME}"
    WriteRegStr   HKLM "${REG_KEY}" "DisplayVersion"  "${APP_VERSION}"
    WriteRegStr   HKLM "${REG_KEY}" "Publisher"       "${APP_PUBLISHER}"
    WriteRegStr   HKLM "${REG_KEY}" "URLInfoAbout"    "${APP_URL}"
    WriteRegStr   HKLM "${REG_KEY}" "InstallLocation" "$INSTDIR"
    WriteRegStr   HKLM "${REG_KEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
    WriteRegStr   HKLM "${REG_KEY}" "DisplayIcon"     '"$INSTDIR\${APP_EXE}"'
    WriteRegDWORD HKLM "${REG_KEY}" "NoModify"        1
    WriteRegDWORD HKLM "${REG_KEY}" "NoRepair"        1
    WriteRegDWORD HKLM "${REG_KEY}" "EstimatedSize"   65536

    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

; ============================================================
; SECTION: Desktop Shortcut
; ============================================================
Section "Desktop Shortcut" SecDesktop
    CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0 SW_SHOWNORMAL "" "${APP_NAME}"
SectionEnd

; ============================================================
; SECTION: Start Menu
; ============================================================
Section "Start Menu" SecStartMenu
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0
    CreateShortcut "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
SectionEnd

; ============================================================
; FUNCTION: Init checks
; ============================================================
Function .onInit
    ${IfNot} ${RunningX64}
        MessageBox MB_OK|MB_ICONSTOP "MySchool Checks requires 64-bit Windows."
        Abort
    ${EndIf}

    ReadRegStr $R0 HKLM "${REG_KEY}" "InstallLocation"
    ${If} $R0 != ""
        MessageBox MB_YESNO|MB_ICONQUESTION "${APP_NAME} is already installed at:$\n$R0$\n$\nReplace existing installation?" IDYES +2
        Abort
        ExecWait 'taskkill /f /im "${APP_EXE}"'
    ${EndIf}
FunctionEnd

; ============================================================
; UNINSTALLER
; ============================================================
Section "Uninstall"
    ExecWait 'taskkill /f /im "${APP_EXE}"'

    Delete "$INSTDIR\${APP_EXE}"
    Delete "$INSTDIR\startup.mp3"
    Delete "$INSTDIR\app.ico"
    Delete "$INSTDIR\MySchoolChecks_Odigos.pdf"
    Delete "$INSTDIR\Uninstall.exe"

    RMDir /r "$INSTDIR\drivers"
    RMDir /r "$INSTDIR\screenshots"
    RMDir /r "$INSTDIR\data"
    RMDir    "$INSTDIR"

    Delete "$DESKTOP\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"
    RMDir  "$SMPROGRAMS\${APP_NAME}"

    DeleteRegKey HKLM "${REG_KEY}"

    MessageBox MB_OK "${APP_NAME} was successfully uninstalled."
SectionEnd
