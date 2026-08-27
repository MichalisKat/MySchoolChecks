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
!define APP_VERSION   "3.7.0"
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
!include "nsDialogs.nsh"
!include "WordFunc.nsh"
!include "WinMessages.nsh"
!insertmacro WordFind

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

; --- Vars: Επιλογή Διεύθυνσης Εκπαίδευσης ---
Var Dialog
Var RadioPE
Var RadioDE
Var ComboBox
Var DirType
Var DirName
Var HasPrevDir

; --- Pages ---
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE"
!insertmacro MUI_PAGE_DIRECTORY
Page custom DirectoratePageCreate DirectoratePageLeave
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

    ; Ξ‘Ο€ΞµΞ³ΞΊΞ±Ο„Ξ¬ΟƒΟ„Ξ±ΟƒΞ· Ο€Ξ±Ξ»ΞΉΞ¬Ο‚ Ξ­ΞΊΞ΄ΞΏΟƒΞ·Ο‚ MySchool Checks (Ξ±Ξ½ Ο…Ο€Ξ¬ΟΟ‡ΞµΞΉ)
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

    ; Αποθήκευση επιλεγμένης Διεύθυνσης Εκπαίδευσης (βλ. σελίδα DirectoratePage)
    FileOpen $4 "$INSTDIR\data\directorate.txt" w
    FileWrite $4 "$DirType|$DirName"
    FileClose $4

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

    StrCpy $DirType "ΠΕ"
FunctionEnd

; ============================================================
; PAGE: Επιλογή Διεύθυνσης Εκπαίδευσης
; ============================================================
Function DirectoratePageCreate
    !insertmacro MUI_HEADER_TEXT "Διεύθυνση Εκπαίδευσης" "Επιλέξτε τη Διεύθυνση Εκπαίδευσης στην οποία ανήκετε"

    ; Αν υπάρχει ήδη καταχωρημένη επιλογή (π.χ. update πάνω σε ήδη
    ; εγκατεστημένη έκδοση), την κρατάμε ΩΣ ΕΧΕΙ — δεν εμφανίζουμε καθόλου
    ; επιλογές προς αλλαγή (μόνο ενημερωτικό μήνυμα), ώστε να μην υπάρχει
    ; περίπτωση να αλλάξει κατά λάθος σε update.
    StrCpy $DirType "ΠΕ"
    StrCpy $DirName ""
    StrCpy $HasPrevDir ""
    ${If} ${FileExists} "$INSTDIR\data\directorate.txt"
        FileOpen $5 "$INSTDIR\data\directorate.txt" r
        FileRead $5 $6
        FileClose $5
        ${WordFind} "$6" "|" "+1" $DirType
        ${WordFind} "$6" "|" "+2" $DirName
        ${If} $DirName != ""
            StrCpy $HasPrevDir "1"
        ${EndIf}
    ${EndIf}

    nsDialogs::Create 1018
    Pop $Dialog
    ${If} $Dialog == error
        Abort
    ${EndIf}

    ${If} $HasPrevDir == "1"
        ; Ήδη καταχωρημένη Διεύθυνση — μόνο ενημερωτικό μήνυμα, τίποτα προς
        ; επιλογή· ο χρήστης απλά πατάει «Επόμενο».
        ${NSD_CreateLabel} 0 0u 100% 60u "Η εφαρμογή είναι ήδη καταχωρημένη για τη Διεύθυνση:$\n$\n$DirType   $DirName$\n$\nΔεν χρειάζεται καμία ενέργεια εδώ — πάτησε «Επόμενο»."
        Pop $0
    ${Else}
        ${NSD_CreateLabel} 0 0u 100% 20u "Η επιλογή χρησιμοποιείται μόνο για να φαίνεται ποιες Διευθύνσεις χρησιμοποιούν το πρόγραμμα (καμία σχέση με τα στοιχεία σύνδεσης MySchool)."
        Pop $0

        ${NSD_CreateRadioButton} 0 26u 45% 12u "Πρωτοβάθμια (Π.Ε.)"
        Pop $RadioPE
        ${NSD_SetState} $RadioPE ${BST_CHECKED}
        ${NSD_OnClick} $RadioPE OnTypeChange

        ${NSD_CreateRadioButton} 50% 26u 45% 12u "Δευτεροβάθμια (Δ.Ε.)"
        Pop $RadioDE
        ${NSD_OnClick} $RadioDE OnTypeChange

        ${NSD_CreateLabel} 0 44u 100% 12u "Διεύθυνση:"
        Pop $0

        ${NSD_CreateDropList} 0 58u 100% 120u ""
        Pop $ComboBox

        ; Ξεκινάει ΑΔΕΙΟ (καμία προεπιλογή) — ο χρήστης πρέπει να διαλέξει ο
        ; ίδιος συνειδητά, ώστε να μην περάσει κατά λάθος λάθος Διεύθυνση.
        Call PopulatePE
    ${EndIf}

    nsDialogs::Show
FunctionEnd

Function OnTypeChange
    ${NSD_GetState} $RadioPE $0
    ${If} $0 == ${BST_CHECKED}
        StrCpy $DirType "ΠΕ"
        Call PopulatePE
    ${Else}
        StrCpy $DirType "ΔΕ"
        Call PopulateDE
    ${EndIf}
FunctionEnd

Function DirectoratePageLeave
    ${If} $HasPrevDir == "1"
        ; Τίποτα να ελέγξουμε — τα $DirType/$DirName είναι ήδη σωστά, όπως
        ; διαβάστηκαν από το προηγούμενο directorate.txt.
        Return
    ${EndIf}
    ${NSD_GetText} $ComboBox $DirName
    ${If} $DirName == ""
        MessageBox MB_OK|MB_ICONEXCLAMATION "Επιλέξτε τη Διεύθυνση Εκπαίδευσης."
        Abort
    ${EndIf}
FunctionEnd

Function PopulatePE
    SendMessage $ComboBox ${CB_RESETCONTENT} 0 0
    ${NSD_CB_AddString} $ComboBox "Α' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Β' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Γ' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Δ' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Ανατολικής Αττικής"
    ${NSD_CB_AddString} $ComboBox "Δυτικής Αττικής"
    ${NSD_CB_AddString} $ComboBox "Πειραιά"
    ${NSD_CB_AddString} $ComboBox "Ανατολικής Θεσσαλονίκης"
    ${NSD_CB_AddString} $ComboBox "Δυτικής Θεσσαλονίκης"
    ${NSD_CB_AddString} $ComboBox "Ημαθίας"
    ${NSD_CB_AddString} $ComboBox "Κιλκίς"
    ${NSD_CB_AddString} $ComboBox "Πέλλας"
    ${NSD_CB_AddString} $ComboBox "Πιερίας"
    ${NSD_CB_AddString} $ComboBox "Σερρών"
    ${NSD_CB_AddString} $ComboBox "Χαλκιδικής"
    ${NSD_CB_AddString} $ComboBox "Δράμας"
    ${NSD_CB_AddString} $ComboBox "Έβρου"
    ${NSD_CB_AddString} $ComboBox "Καβάλας"
    ${NSD_CB_AddString} $ComboBox "Ξάνθης"
    ${NSD_CB_AddString} $ComboBox "Ροδόπης"
    ${NSD_CB_AddString} $ComboBox "Βοιωτίας"
    ${NSD_CB_AddString} $ComboBox "Εύβοιας"
    ${NSD_CB_AddString} $ComboBox "Ευρυτανίας"
    ${NSD_CB_AddString} $ComboBox "Φθιώτιδας"
    ${NSD_CB_AddString} $ComboBox "Φωκίδας"
    ${NSD_CB_AddString} $ComboBox "Λάρισας"
    ${NSD_CB_AddString} $ComboBox "Μαγνησίας"
    ${NSD_CB_AddString} $ComboBox "Τρικάλων"
    ${NSD_CB_AddString} $ComboBox "Καρδίτσας"
    ${NSD_CB_AddString} $ComboBox "Ηρακλείου"
    ${NSD_CB_AddString} $ComboBox "Λασιθίου"
    ${NSD_CB_AddString} $ComboBox "Ρεθύμνης"
    ${NSD_CB_AddString} $ComboBox "Χανίων"
    ${NSD_CB_AddString} $ComboBox "Άρτας"
    ${NSD_CB_AddString} $ComboBox "Θεσπρωτίας"
    ${NSD_CB_AddString} $ComboBox "Ιωαννίνων"
    ${NSD_CB_AddString} $ComboBox "Πρέβεζας"
    ${NSD_CB_AddString} $ComboBox "Γρεβενών"
    ${NSD_CB_AddString} $ComboBox "Καστοριάς"
    ${NSD_CB_AddString} $ComboBox "Κοζάνης"
    ${NSD_CB_AddString} $ComboBox "Φλώρινας"
    ${NSD_CB_AddString} $ComboBox "Ζακύνθου"
    ${NSD_CB_AddString} $ComboBox "Κέρκυρας"
    ${NSD_CB_AddString} $ComboBox "Κεφαλληνίας"
    ${NSD_CB_AddString} $ComboBox "Λευκάδας"
    ${NSD_CB_AddString} $ComboBox "Αιτωλοακαρνανίας"
    ${NSD_CB_AddString} $ComboBox "Αχαΐας"
    ${NSD_CB_AddString} $ComboBox "Ηλείας"
    ${NSD_CB_AddString} $ComboBox "Λέσβου"
    ${NSD_CB_AddString} $ComboBox "Σάμου"
    ${NSD_CB_AddString} $ComboBox "Χίου"
    ${NSD_CB_AddString} $ComboBox "Λήμνου"
    ${NSD_CB_AddString} $ComboBox "Δωδεκανήσου"
    ${NSD_CB_AddString} $ComboBox "Κυκλάδων"
    ${NSD_CB_AddString} $ComboBox "Αργολίδας"
    ${NSD_CB_AddString} $ComboBox "Αρκαδίας"
    ${NSD_CB_AddString} $ComboBox "Κορινθίας"
    ${NSD_CB_AddString} $ComboBox "Λακωνίας"
    ${NSD_CB_AddString} $ComboBox "Μεσσηνίας"
FunctionEnd

Function PopulateDE
    SendMessage $ComboBox ${CB_RESETCONTENT} 0 0
    ${NSD_CB_AddString} $ComboBox "Α' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Β' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Γ' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Δ' Αθήνας"
    ${NSD_CB_AddString} $ComboBox "Ανατολικής Αττικής"
    ${NSD_CB_AddString} $ComboBox "Δυτικής Αττικής"
    ${NSD_CB_AddString} $ComboBox "Πειραιά"
    ${NSD_CB_AddString} $ComboBox "Ανατολικής Θεσσαλονίκης"
    ${NSD_CB_AddString} $ComboBox "Δυτικής Θεσσαλονίκης"
    ${NSD_CB_AddString} $ComboBox "Ημαθίας"
    ${NSD_CB_AddString} $ComboBox "Κιλκίς"
    ${NSD_CB_AddString} $ComboBox "Πέλλας"
    ${NSD_CB_AddString} $ComboBox "Πιερίας"
    ${NSD_CB_AddString} $ComboBox "Σερρών"
    ${NSD_CB_AddString} $ComboBox "Χαλκιδικής"
    ${NSD_CB_AddString} $ComboBox "Δράμας"
    ${NSD_CB_AddString} $ComboBox "Έβρου"
    ${NSD_CB_AddString} $ComboBox "Καβάλας"
    ${NSD_CB_AddString} $ComboBox "Ξάνθης"
    ${NSD_CB_AddString} $ComboBox "Ροδόπης"
    ${NSD_CB_AddString} $ComboBox "Βοιωτίας"
    ${NSD_CB_AddString} $ComboBox "Εύβοιας"
    ${NSD_CB_AddString} $ComboBox "Ευρυτανίας"
    ${NSD_CB_AddString} $ComboBox "Φθιώτιδας"
    ${NSD_CB_AddString} $ComboBox "Φωκίδας"
    ${NSD_CB_AddString} $ComboBox "Λάρισας"
    ${NSD_CB_AddString} $ComboBox "Μαγνησίας"
    ${NSD_CB_AddString} $ComboBox "Τρικάλων"
    ${NSD_CB_AddString} $ComboBox "Καρδίτσας"
    ${NSD_CB_AddString} $ComboBox "Ηρακλείου"
    ${NSD_CB_AddString} $ComboBox "Λασιθίου"
    ${NSD_CB_AddString} $ComboBox "Ρεθύμνης"
    ${NSD_CB_AddString} $ComboBox "Χανίων"
    ${NSD_CB_AddString} $ComboBox "Άρτας"
    ${NSD_CB_AddString} $ComboBox "Θεσπρωτίας"
    ${NSD_CB_AddString} $ComboBox "Ιωαννίνων"
    ${NSD_CB_AddString} $ComboBox "Πρέβεζας"
    ${NSD_CB_AddString} $ComboBox "Γρεβενών"
    ${NSD_CB_AddString} $ComboBox "Καστοριάς"
    ${NSD_CB_AddString} $ComboBox "Κοζάνης"
    ${NSD_CB_AddString} $ComboBox "Φλώρινας"
    ${NSD_CB_AddString} $ComboBox "Ζακύνθου"
    ${NSD_CB_AddString} $ComboBox "Κέρκυρας"
    ${NSD_CB_AddString} $ComboBox "Κεφαλληνίας"
    ${NSD_CB_AddString} $ComboBox "Λευκάδας"
    ${NSD_CB_AddString} $ComboBox "Αιτωλοακαρνανίας"
    ${NSD_CB_AddString} $ComboBox "Αχαΐας"
    ${NSD_CB_AddString} $ComboBox "Ηλείας"
    ${NSD_CB_AddString} $ComboBox "Λέσβου"
    ${NSD_CB_AddString} $ComboBox "Σάμου"
    ${NSD_CB_AddString} $ComboBox "Χίου"
    ${NSD_CB_AddString} $ComboBox "Δωδεκανήσου"
    ${NSD_CB_AddString} $ComboBox "Κυκλάδων"
    ${NSD_CB_AddString} $ComboBox "Αργολίδας"
    ${NSD_CB_AddString} $ComboBox "Αρκαδίας"
    ${NSD_CB_AddString} $ComboBox "Κορινθίας"
    ${NSD_CB_AddString} $ComboBox "Λακωνίας"
    ${NSD_CB_AddString} $ComboBox "Μεσσηνίας"
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
