# MySchool Checks — Οδηγός για Claude

## Τι είναι αυτό το project

Python εφαρμογή Windows για αυτοματοποιημένους ελέγχους δεδομένων εκπαιδευτικών στο MySchool (Διεύθυνση Π.Ε. Ανατολικής Θεσσαλονίκης). Κατεβάζει αρχεία μέσω Selenium/Chrome ή Firefox, επεξεργάζεται CSV/Excel με pandas, παράγει αποτελέσματα Excel και στέλνει emails.

**Repo:** https://github.com/MichalisKat/myschool-checks  
**Υπεύθυνος:** Μιχάλης Κατσιρντάκης  
**Τρέχουσα έκδοση:** 0.9.6 (beta)

---

## Δομή project

```
myschool-checks/
├── MySchoolChecks/
│   ├── main.py              # Κύριο αρχείο — UI, splash, settings, load_checks(), EidikotitaDialog, MonadaDialog
│   ├── config.py            # Ρυθμίσεις — φορτώνει JSON + keyring, APP_VERSION
│   ├── encryption.py        # Keyring wrapper — store/get/delete credentials
│   ├── setup_credentials.py # CLI wizard πρώτης ρύθμισης credentials
│   ├── app.ico              # Εικονίδιο εφαρμογής (6 μεγέθη: 16-256px)
│   ├── startup.mp3          # Μουσική splash (Pixabay License — δεν ανεβαίνει στο GitHub)
│   ├── checks/              # Ένα .py ανά έλεγχο
│   └── core/
│       ├── framework.py     # run_check(), get_downloaded_file(), send_email()
│       └── downloader.py    # MySchoolDownloader, REPORTS, get_downloads_dir()
├── gen_multi.py             # Standalone script — batch εξαγωγή πολλών ειδικοτήτων (dev/test)
├── MySchoolChecks_Odigos.pdf  # Οδηγός χρήστη (ReportLab, ελληνικά)
├── MySchoolChecks.spec      # PyInstaller spec — ΜΗΝ το διαγράψεις
├── build_executable.bat     # Φτιάχνει dist\MySchoolChecks.exe
├── compile_installer.bat    # Φτιάχνει myschool-checks-0.9.6-setup.exe
├── myschool-checks.nsi      # NSIS script για τον installer
├── SECURITY.md              # Τεκμηρίωση ασφάλειας credentials
└── CLAUDE.md                # Αυτό το αρχείο
```

---

## Κρίσιμες αρχιτεκτονικές αποφάσεις

### Credentials — Keyring (Windows Credential Manager)
Τα ευαίσθητα credentials (`MYSCHOOL_USER`, `MYSCHOOL_PASS`, `FROM_PASSWORD`) αποθηκεύονται **μόνο** στο Windows Credential Manager μέσω του `keyring` module. Δεν γράφονται σε κανένα αρχείο.

- `encryption.py` — wrapper: `store_credential`, `get_credential`, `delete_credential`
- `config.py` → `_load_local()` — φορτώνει μη-ευαίσθητα από JSON, ευαίσθητα από keyring
- `main.py` → `_save_config()` — αποθηκεύει ευαίσθητα στο keyring, μη-ευαίσθητα στο JSON
- `*.mp3` είναι στο `.gitignore` — δεν ανεβαίνει στο GitHub

### Paths — _app_base()
Βρίσκεται στο `main.py`:
- **Εγκατεστημένο (Program Files)** → `%LOCALAPPDATA%\MySchoolChecks\`
- **Portable (dist\)** → δίπλα στο .exe
- **Development** → φάκελος του κώδικα

Η ίδια λογική υπάρχει και στο `config.py` (`_load_local()`) και στο `framework.py`.

### Downloads — _docs_base()
Τα κατεβασμένα στατιστικά αρχεία πηγαίνουν στο `Documents\MySchoolChecks\downloads\YYYYMMDD\` (ίδιος φάκελος με τα results) — ΟΧΙ στο LOCALAPPDATA.

### Αποθήκευση δεδομένων
| Τι | Πού |
|---|---|
| Μη-ευαίσθητες ρυθμίσεις | `%LOCALAPPDATA%\MySchoolChecks\data\local_settings.json` |
| Ευαίσθητα credentials | Windows Credential Manager (keyring) |
| Downloads | `Documents\MySchoolChecks\downloads\YYYYMMDD\` |
| Αποτελέσματα | `Documents\MySchoolChecks\results_YYYYMMDD\` |

### Browser
Υποστηρίζονται Chrome και Firefox. Επιλογή από Ρυθμίσεις → Σύνδεση → radio button.
- `BROWSER = 'chrome'` ή `'firefox'` στο config
- `downloader.py` → `MySchoolDownloader(browser=...)` — split λογική Chrome/Firefox

### Frozen exe (PyInstaller)
- `sys.frozen = True` → το app τρέχει ως .exe
- `sys._MEIPASS` → temp φάκελος (διαγράφεται μετά) — ΔΕΝ γράφουμε εκεί
- `sys.executable` → το .exe αρχείο
- Τα checks φορτώνονται δυναμικά — στο frozen exe δεν υπάρχουν .py, γι' αυτό το spec τα bundlάρει με `--add-data`

### Checks
Κάθε αρχείο στο `checks/` είναι ανεξάρτητος έλεγχος. Πρέπει να έχει:
- `CHECK_TITLE` — τίτλος
- `REQUIRED_REPORTS` — list με τα απαιτούμενα στατιστικά (εμφανίζεται στο dialog όταν λείπουν αρχεία)
- `ask_inputs()` — ζητά αρχεία/παραμέτρους, επιστρέφει context dict
- `process(ctx)` — λογική ελέγχου, επιστρέφει DataFrame

Αν `ask_inputs()` επιστρέψει `None` σε κλειδί `path`/`*_path`, το `run_check()` καλεί `_missing_file_dialog()` με τα `REQUIRED_REPORTS` και τερματίζει καθαρά.

### EidikotitaDialog (main.py)
Εργαλείο εξαγωγής εκπαιδευτικών ανά ειδικότητα — ανεξάρτητο από το σύστημα checks.
- Κουμπί toolbar: **«📋 Εκπ/κοί ανά Ειδικότητα»**
- Πηγές: Topothetiseis + gridResults (2.1) + stat4_1 + stat4_2 + stat4_16
- **Dropdown Διεύθυνση**: φορτώνει μοναδικές τιμές από col19 (Περιοχή Μετάθεσης Φορέα). Η επιλογή φιλτράρει τις ειδικότητες **και** τα αποτελέσματα. Λειτουργεί τόσο με ΔΠΕ (Π.Ε.) όσο και ΔΔΕ (Δ.Ε.) αρχεία.
- Φίλτρα: εξαιρεί ΠΑΡΗΛΘΕ, Ξένο Σχολείο, Ιδ.Δικαίου, Υπερωριακά, Μερική Διάθεση, Τοποθέτηση Διοικητικού — φιλτράρισμα Διεύθυνσης γίνεται από dropdown (όχι hard-coded)
- Χρήση `str.contains` αντί `isin` γιατί οι τιμές έχουν παρενθετικά suffixes
- AFM join: `.str.zfill(9)` και στις δύο πλευρές (Topothetiseis αποθηκεύει ως float)
- stat4_*/stat4_16: 1-column shift στα headers — χρήση `df.columns[N]` με absolute index
- Επιλογή στηλών εξόδου (checkboxes): Email ΠΣΔ, Email, Κινητό
- Απόντες: κόκκινο font, ΑΠΟΥΣΙΑ/Έως εμφανίζονται μόνο για αυτούς
- Αλφαβητική ταξινόμηση κατά Επώνυμο

### DownloadDialog (main.py)
- Checkboxes: ξεκινούν απενεργοποιημένα by default (ο χρήστης επιλέγει)
- Κουμπί «Όλα» τσεκάρει όλα μαζί
- Αρχεία που υπάρχουν ήδη εμφανίζονται με ✓ και πράσινο χρώμα

### MonadaDialog (main.py)
Εργαλείο εξαγωγής στοιχείων σχολικών μονάδων ανά Δήμο — ανεξάρτητο από το σύστημα checks.
- Κουμπί toolbar: **«🏫 Σχολικές Μονάδες»**
- Πηγές: stat3_1 (primary — κατανομή τάξεων) + CSV/stat2_2 (secondary — επικοινωνία, Διευθυντής)
- Φίλτρα: Δημοτικά + Νηπιαγωγεία, εξαιρεί Ιδιωτικά/Ξένα και Αναστολή=NAI (col48 CSV)
- **CSV 1-column shift**: col10=Είδος, col11=Κωδ., col12=Ονομασία, col16=Τηλ., col18=email, col20=Διεύθ., col48=Αναστολή, col55=Διευθυντής, col58=Κινητό, col59=Email, col60=Email ΠΣΔ
- stat3_1 trailing semicolons: `strip_trailing_sep=True` στο `_read_zip_csv`
- Join: clean_code (lstrip '0', strip `="..."`) στα codes stat3_1 col4 ↔ CSV col11
- **Ανά Τάξη**: flat rows + subtotal μόνο για Δημοτικά (Νηπιαγωγεία χωρίς subtotal)
- **Ανά Μονάδα**: groupby κωδικού στο stat3_1, sum Τμήματα/Αγόρια/Κορίτσια
- `_auto_find_zip(*prefixes)`: δέχεται πολλαπλά prefixes, ψάχνει .zip → .csv → .xlsx
- `_read_zip_csv`: χειρίζεται .zip, plain .csv, και .xlsx
- Δήμοι από stat3_1 col7 (fallback: CSV col6)
- Email body: αυτόματη αντικατάσταση `{dimos}` στο `_on_dimos_change`

### DownloadDialog (main.py)
- Checkboxes: ξεκινούν απενεργοποιημένα by default (ο χρήστης επιλέγει)
- Κουμπί «Όλα» τσεκάρει όλα μαζί
- Αρχεία που υπάρχουν ήδη εμφανίζονται με ✓ και πράσινο χρώμα

### downloader.py — REPORTS tuple
Κάθε entry: `(rid, label, url_path, fname_base, wait_search, wait_dl, direct_export, custom_search?, custom_export?, pre_search_labels?)`
- `custom_search`: CSS selector για κουμπί αναζήτησης (π.χ. `'a.hint_search'` για topoth)
- `custom_export`: CSS selector για κουμπί εξαγωγής
- `pre_search_labels`: list με κείμενα `<label>` checkboxes που τσεκάρονται πριν την αναζήτηση (χρησιμοποιείται για 3.1)
- Grid wait: flexible XPath `//*[contains(@id,"DXDataRow0")]` — λειτουργεί για όλες τις σελίδες
- Αρχεία: 2.2 (stat2_2) = Εκτεταμένα Στοιχεία, 3.1 (stat3_1) = Κατανομή μαθητών

---

## Διαδικασία build & release

### Όταν αλλάζει κώδικας:

```bash
# 1. Πάρε τελευταίες αλλαγές
git pull origin main

# 2. Φτιάξε νέο exe (κρατάει αυτόματα τα credentials)
build_executable.bat

# 3. Φτιάξε νέο installer
compile_installer.bat

# 4. Απεγκατάστησε παλιά έκδοση
# Ρυθμίσεις → Εφαρμογές → MySchool Checks → Απεγκατάσταση

# 5. Εγκατάστησε νέα
# Τρέξε myschool-checks-0.9.0-setup.exe

# 6. Ανέβασε στο GitHub Releases
gh release delete v0.9.0 --yes
gh release create v0.9.0 "myschool-checks-0.9.0-setup.exe" --title "MySchool Checks v0.9.0 (beta)" --notes "Ενημερωμένη έκδοση"
```

### Όταν αλλάζει μόνο ο installer (π.χ. .nsi, LICENSE):
Τρέξε μόνο βήματα 3 → 4 → 5 → 6. Το `build_executable.bat` δεν χρειάζεται.

### Versioning
- Τρέχουσα: `0.9.6` (beta)
- Stable release: `1.0.0` (μετά από testing)
- Αλλαγή version: στο `myschool-checks.nsi` (`APP_VERSION`), στο `compile_installer.bat` **και** στο `MySchoolChecks/config.py` (`APP_VERSION`)

### Backup credentials
Το `build_executable.bat` κρατάει αυτόματα τα credentials:
- Πριν το build: backup από `%LOCALAPPDATA%\MySchoolChecks\data\` ή `dist\data\`
- Μετά το build: επαναφορά και στα δύο μέρη

---

## Εργαλεία που χρειάζονται

| Εργαλείο | Path |
|---|---|
| Python 3.14 | `C:\Users\mkatsirntakis\AppData\Local\Python\pythoncore-3.14-64\` |
| PyInstaller | `pyinstaller` (στο PATH) |
| keyring | `pip install keyring` (απαιτείται για build) |
| NSIS | `C:\Program Files (x86)\NSIS\makensis.exe` |
| Git | Στο PATH |
| GitHub CLI | Στο PATH (`gh auth login` αν δεν είναι συνδεδεμένο) |

> **Σημείωση build:** Το `build_executable.bat` δεν τρέχει σωστά από bash/Claude — χρησιμοποίησε `pyinstaller MySchoolChecks.spec` απευθείας και μετά `makensis myschool-checks.nsi`.

---

## Συνηθισμένα προβλήματα & λύσεις

| Πρόβλημα | Αιτία | Λύση |
|---|---|---|
| "keyring μη διαθέσιμο" στο app | keyring δεν είναι εγκατεστημένο στο Python | `pip install keyring` → ξανabuild |
| Batch file με σκουπίδια | LF αντί CRLF | `$content -replace "(?<!\r)\n", "\`r\`n"` σε PowerShell |
| Settings χάθηκαν μετά rebuild | Backup από λάθος φάκελο | Το build_executable.bat το διορθώνει αυτόματα |
| WinError 5 (Program Files) | Εγγραφή σε Program Files χωρίς admin | `_app_base()` ανακατευθύνει στο LOCALAPPDATA |
| 0 checks φορτώθηκαν | Import error στο frozen exe | Ελέγχα `checks_errors.log` στο Desktop |
| Infinite pip loop | `sys.executable` = το .exe | Skip pip όταν `sys.frozen = True` |
| Spec file not found | `del /q *.spec` στο build script | Το script ΔΕΝ διαγράφει το .spec |
| git index.lock | Crashed git process | `del /f .git\index.lock` στο repo folder |

---

## Pending / Επόμενα βήματα

- [ ] Beta testing → stable release v1.0.0
- [ ] MSIX package για Microsoft Store (αναβλήθηκε — χρειάζεται Windows SDK + Partner Center)
- [ ] Αν χρειαστεί νέα έκδοση: αλλαγή `APP_VERSION` στο `.nsi`, `compile_installer.bat` **και** `config.py`
- [ ] Επαλήθευση λήψης 3.1 με pre_search_labels (DevExpress checkboxes) — πρώτο run in production
- [ ] Ενημέρωση PDF οδηγού χρήστη (χειροκίνητα με ReportLab) — να προστεθεί ενότητα Σχολικών Μονάδων

## Αλλαγές ανά έκδοση

### v0.9.6
- Κωδικός email **προαιρετικός** — αποθηκεύεται αλλά δεν απαιτείται για αποθήκευση ρυθμίσεων
- Κανονική αποστολή **απενεργοποιημένη** στο dialog αν δεν υπάρχει κωδικός (με hint για Ρυθμίσεις)
- `forma_82.py`: **αυτόματη cutoff** — 1η (ημέρες 1-14) ή 15η (ημέρες 15-31) τρέχοντος μήνα
- `framework.py` `get_downloaded_file()`: **διορθώθηκε path** — χρησιμοποιεί `Documents\MySchoolChecks\` (όχι LOCALAPPDATA)
- `framework.py` `_missing_file_dialog()`: **νέα συνάρτηση** — εμφανίζει workflow hint + REQUIRED_REPORTS όταν λείπουν αρχεία
- Κάθε check module έχει πλέον `REQUIRED_REPORTS` list
- EidikotitaDialog + MonadaDialog: χρήση `_missing_file_dialog()` αντί για showwarning
- Toolbar: κουμπί «⬇ Λήψη Δεδομένων» εμφανίζεται **έντονο** (πράσινο) αν δεν υπάρχουν αρχεία σήμερα
- Status bar: hint «💡 Ξεκινήστε με ⬇ Λήψη Δεδομένων» αν δεν υπάρχουν αρχεία σήμερα
- Version εμφανίζεται στο **status bar** (δεξιά) αντί για header
- Διόρθωση `NameError` crash κατά την εκκίνηση (`_has_files` ορίζεται πριν τη χρήση του)
- Διόρθωση `TclError` crash στον auto-updater progressbar (`_safe_after` + `winfo_exists`)
- `framework.py` `send_email()`: επιστρέφει `False` αθόρυβα αν δεν υπάρχει κωδικός
- `8ball.ico` → `app.ico` σε όλες τις αναφορές του `framework.py`

### v0.9.5
- EidikotitaDialog: εξαγωγή εκπαιδευτικών ανά ειδικότητα
- MonadaDialog: εξαγωγή στοιχείων σχολικών μονάδων ανά Δήμο
- DownloadDialog: checkboxes απενεργοποιημένα by default, κουμπί «Όλα»
- downloader.py REPORTS: custom_search, custom_export, pre_search_labels

---

## Git workflow (δύο υπολογιστές)

**Πάντα πριν δουλέψεις:**
```bash
git pull origin main
```

**Μετά από αλλαγές:**
```bash
git add .
git commit -m "περιγραφή"
git push origin main
```
