# MySchool Checks - Release Checklist v1.0

> **Αυτό το checklist χρησιμοποιείται για κάθε release.**  
> Ακολούθησε τα βήματα με τη σειρά. Κάθε ✅ = ολοκληρώθηκε.

---

## 1. PRE-RELEASE CHECKLIST

Πριν φτιάξεις οτιδήποτε, ελέγξε τα παρακάτω:

- [ ] **Test locally** — Τρέξε `MySchoolChecksPlus\main.py` με Python και βεβαιώσου ότι όλα δουλεύουν
- [ ] **ChromeDriver** — Βεβαιώσου ότι `MySchoolChecksPlus\drivers\chromedriver-win64\` περιέχει το `chromedriver.exe`
- [ ] **Chrome version** — Το ChromeDriver να ταιριάζει με την έκδοση Chrome που έχει ο χρήστης
- [ ] **Version bump** — Άλλαξε τον αριθμό έκδοσης στα παρακάτω αρχεία:
  - `MySchoolChecksPlus\config.py` (αν υπάρχει `VERSION =`)
  - `myschool-checks.nsi` → `!define APP_VERSION "1.1"`
  - `create_msix.bat` → `set APP_VERSION=1.1.0.0`
  - `AppxManifest.xml` → `Version="1.1.0.0"`
- [ ] **Icon** — Υπάρχει `MySchoolChecksPlus\8ball.ico`
- [ ] **Assets για Store** — Υπάρχουν PNG αρχεία (βλ. Φάση 3)
- [ ] **LICENSE** — Υπάρχει `LICENSE` file στο root (MIT)
- [ ] **Requirements** — Ενημερωμένο `MySchoolChecksPlus\requirements.txt`
- [ ] **Git status** — Όλες οι αλλαγές committed

```bash
# Version bump example (από v1.0 σε v1.1)
git add .
git commit -m "chore: bump version to 1.1"
git tag v1.1
```

---

## 2. BUILD PROCESS

### 2.1 — PyInstaller Executable

```batch
# Τρέξε από τον φάκελο myschool-checks\
build_executable.bat
```

**Αναμενόμενο output:** `dist\MySchoolChecksPlus.exe`

**Έλεγχος:**
- [ ] `dist\MySchoolChecksPlus.exe` υπάρχει
- [ ] Τρέξε το `.exe` — ανοίγει χωρίς console παράθυρο
- [ ] Test σε καθαρό Windows χωρίς Python (VM ή άλλο PC)

### 2.2 — NSIS Installer

**Απαιτεί:** [NSIS 3.x](https://nsis.sourceforge.io/Download)

```batch
# Επιλογή 1: Από NSIS GUI
# Right-click στο myschool-checks.nsi -> "Compile NSIS Script"

# Επιλογή 2: Command line
"C:\Program Files (x86)\NSIS\makensis.exe" myschool-checks.nsi
```

**Αναμενόμενο output:** `myschool-checks-1.0-setup.exe`

**Έλεγχος:**
- [ ] `myschool-checks-1.0-setup.exe` υπάρχει
- [ ] Τρέξε το setup και εγκατάστησε σε test φάκελο
- [ ] Desktop shortcut δημιουργήθηκε
- [ ] Start Menu entry δημιουργήθηκε
- [ ] Add/Remove Programs δείχνει την εφαρμογή
- [ ] Uninstall δουλεύει

### 2.3 — MSIX Package (Microsoft Store)

**Απαιτεί:** [Windows 10/11 SDK](https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/)

**Πριν τρέξεις:** Αντικατάστησε τα placeholder PNG στο `msix_package\Assets\`:

| File | Μέγεθος | Χρήση |
|------|---------|-------|
| `Square44x44Logo.png` | 44×44 px | Taskbar icon |
| `Square150x150Logo.png` | 150×150 px | Start menu tile |
| `StoreLogo.png` | 50×50 px | Store listing |
| `Wide310x150Logo.png` | 310×150 px | Wide tile |

```batch
create_msix.bat
```

**Αναμενόμενο output:** `MySchoolChecksPlus-1.0.msix`

**Έλεγχος:**
- [ ] `MySchoolChecksPlus-1.0.msix` υπάρχει
- [ ] `msix_package\AppxManifest.xml` δείχνει σωστά metadata
- [ ] Δοκιμαστική εγκατάσταση: `Add-AppxPackage MySchoolChecksPlus-1.0.msix` (PowerShell)

### 2.4 — Verify All Outputs

```
dist\
  MySchoolChecksPlus.exe            ← Standalone executable
myschool-checks-1.0-setup.exe  ← NSIS installer
MySchoolChecksPlus-1.0.msix        ← Microsoft Store package
```

---

## 3. GITHUB RELEASE

### 3.1 — Δημιουργία Tag

```bash
git tag v1.0
git push origin v1.0
```

### 3.2 — Δημιουργία Release

```bash
# Με GitHub CLI
gh release create v1.0 \
  dist/MySchoolChecksPlus.exe \
  myschool-checks-1.0-setup.exe \
  --title "MySchool Checks v1.0" \
  --notes "$(cat <<'EOF'
## MySchool Checks v1.0

Πρώτη επίσημη έκδοση της εφαρμογής MySchool Checks.

### Τι είναι νέο
- Αυτοματοποιημένοι έλεγχοι δεδομένων MySchool
- Εξαγωγή αποτελεσμάτων σε Excel
- Windows 10/11 υποστήριξη

### Εγκατάσταση
Κατέβασε το `myschool-checks-1.0-setup.exe` για γρήγορη εγκατάσταση,
ή το `MySchoolChecksPlus.exe` για portable χρήση.

### Απαιτήσεις
- Windows 10/11 (64-bit)
- Google Chrome (για Selenium automation)
EOF
)"
```

**Checklist GitHub:**
- [ ] Tag `v1.0` push έγινε
- [ ] Release δημιουργήθηκε
- [ ] `MySchoolChecksPlus.exe` upload
- [ ] `myschool-checks-1.0-setup.exe` upload
- [ ] Release notes γράφτηκαν (Ελληνικά + Αγγλικά)
- [ ] Release marked as "Latest"

---

## 4. MICROSOFT STORE SUBMISSION

### 4.1 — Προετοιμασία

1. [ ] Εγγραφή στο **Microsoft Partner Center**:  
   👉 https://partner.microsoft.com/dashboard
   
2. [ ] **Κόστος εγγραφής:** ~$19 USD (εφάπαξ, individual account)

3. [ ] Δημιουργία νέας εφαρμογής:
   - Partner Center → Apps & Games → "New product" → "MSIX or PWA app"
   - App name: `MySchool Checks`

4. [ ] **Υπογραφή MSIX** (απαιτείται για Store):
   ```batch
   # Ο Partner Center υπογράφει αυτόματα κατά την υποβολή
   # Για local testing, χρησιμοποίησε self-signed certificate:
   
   # Δημιουργία self-signed cert (PowerShell ως Admin):
   New-SelfSignedCertificate -Type Custom -Subject "CN=MichalisKat, O=Education, C=GR" `
     -KeyUsage DigitalSignature -FriendlyName "MySchoolChecksPlus" `
     -CertStoreLocation "Cert:\CurrentUser\My"
   ```

### 4.2 — Metadata για Store Listing

Συμπλήρωσε τα παρακάτω στο Partner Center:

| Field | Τιμή |
|-------|------|
| **App name** | MySchool Checks |
| **Description (GR)** | Αυτοματοποιημένη εφαρμογή για ελέγχους δεδομένων στο σύστημα MySchool. Εξοικονομεί χρόνο σε εκπαιδευτικούς και διοικητικούς. |
| **Description (EN)** | Automated data verification tool for the MySchool educational management system. |
| **Category** | Productivity |
| **Age rating** | 3+ (IARC) |
| **Privacy policy URL** | https://github.com/MichalisKat/myschool-checks#readme |
| **Support URL** | https://github.com/MichalisKat/myschool-checks/issues |
| **Website** | https://github.com/MichalisKat/myschool-checks |

### 4.3 — Screenshots για Store

Απαιτούνται τουλάχιστον **4 screenshots** (1366×768 ή 1920×1080):
- [ ] Κύρια οθόνη εφαρμογής
- [ ] Οθόνη αποτελεσμάτων ελέγχου
- [ ] Οθόνη εξαγωγής Excel
- [ ] Οθόνη ρυθμίσεων/config

### 4.4 — Υποβολή

1. [ ] Upload `MySchoolChecksPlus-1.0.msix`
2. [ ] Συμπλήρωσε Store listings (GR + EN)
3. [ ] Upload screenshots
4. [ ] Content ratings quiz
5. [ ] Submit for certification

**Αναμενόμενος χρόνος έγκρισης:** 24-72 ώρες

---

## 5. UPDATES PROCEDURE

Για κάθε νέα έκδοση (π.χ. v1.1):

```bash
# 1. Άλλαξε version σε όλα τα files
# 2. Commit
git commit -m "chore: bump version to 1.1"
git tag v1.1
git push origin v1.1

# 3. Rebuild
build_executable.bat
makensis myschool-checks.nsi
create_msix.bat

# 4. GitHub Release
gh release create v1.1 dist/MySchoolChecksPlus.exe myschool-checks-1.1-setup.exe

# 5. Store: Partner Center → Apps → MySchool Checks → New submission
# Upload MySchoolChecksPlus-1.1.msix
```

**Σημαντικό:** Κάθε Store update χρειάζεται νέα version number στο `AppxManifest.xml`.  
Format: `MAJOR.MINOR.BUILD.REVISION` (π.χ. `1.1.0.0`)

---

## 6. TROUBLESHOOTING

### PyInstaller Issues

| Πρόβλημα | Λύση |
|----------|------|
| `ModuleNotFoundError` κατά το build | `pip install <module>` και πρόσθεσε `--hidden-import <module>` |
| Το .exe κλείνει αμέσως | Αφαίρεσε το `--windowed`, τρέξε από cmd για να δεις error |
| Μεγάλο μέγεθος .exe (>200MB) | Χρησιμοποίησε virtual environment με μόνο τα απαραίτητα packages |
| Antivirus blocks .exe | Αναφορά false positive στον antivirus vendor |
| `selenium` import error | `pip install selenium webdriver-manager` + `--collect-all selenium` |

### NSIS Errors

| Πρόβλημα | Λύση |
|----------|------|
| `File not found: dist\MySchoolChecksPlus.exe` | Τρέξε πρώτα `build_executable.bat` |
| Greek characters broken | Βεβαιώσου ότι έχεις `Unicode True` στο .nsi |
| `Error in script` line X | Άνοιξε το .nsi σε NSIS editor για syntax check |

### Store Submission Rejections

| Λόγος απόρριψης | Λύση |
|-----------------|------|
| "Missing privacy policy" | Πρόσθεσε URL στο README ή φτιάξε privacy policy page |
| "Invalid publisher" | Βεβαιώσου ότι Publisher στο AppxManifest ταιριάζει με το cert |
| "App crashes on launch" | Test με `Add-AppxPackage` πριν την υποβολή |
| "Missing assets" | Αντικατέστησε placeholder PNG με πραγματικά εικονίδια |
| "runFullTrust capability" | Απαιτεί επιπλέον documentation — εξήγησε γιατί χρειάζεσαι Selenium |

### Chrome WebDriver Issues

| Πρόβλημα | Λύση |
|----------|------|
| `chromedriver` not found | Βεβαιώσου ότι `drivers\chromedriver.exe` περιλαμβάνεται στο build |
| Version mismatch | Χρησιμοποίησε `webdriver-manager` για αυτόματη διαχείριση versions |
| Chrome not installed | Ενημέρωσε τον χρήστη να εγκαταστήσει Chrome — είναι prerequisite |
| Headless mode error | Πρόσθεσε `--headless=new` ή `--no-sandbox` στα Chrome options |

---

## Useful Links

| Πόρος | URL |
|-------|-----|
| Microsoft Partner Center | https://partner.microsoft.com/dashboard |
| Windows SDK Download | https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/ |
| MSIX Packaging Tool | https://aka.ms/msixpackagingtool |
| NSIS Download | https://nsis.sourceforge.io/Download |
| PyInstaller Docs | https://pyinstaller.org/en/stable/ |
| MSIX Documentation | https://docs.microsoft.com/en-us/windows/msix/ |
| Store Certification Requirements | https://docs.microsoft.com/en-us/windows/apps/publish/store-policies |
| GitHub Releases | https://github.com/MichalisKat/myschool-checks/releases |

---

*Last updated: 2026-04-23 | Maintained by Michalis Katsirintakis*
