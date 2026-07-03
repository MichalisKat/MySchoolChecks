# config.py
# ══════════
# Προεπιλεγμένες τιμές — δεν περιέχουν κωδικούς ή προσωπικά στοιχεία.
# Μη-ευαίσθητες ρυθμίσεις: data/local_settings.json (gitignored)
# Ευαίσθητα credentials: Windows Credential Manager μέσω keyring (encryption.py)

# ── Έκδοση εφαρμογής ──────────────────────────────────────────────────────
APP_VERSION = '2.9.5'

# ── MySchool credentials (κενά — συμπληρώνονται από Ρυθμίσεις) ────────────
MYSCHOOL_USER = ''
MYSCHOOL_PASS = ''

# ── Email ──────────────────────────────────────────────────────────────────
SMTP_HOST     = 'mail.sch.gr'
FROM_EMAIL    = ''
FROM_NAME     = ''
FROM_PASSWORD = ''
TEST_EMAIL    = ''
NOTIFY_EMAIL  = ''   # Email για ενημερωτικό μετά από κάθε αποστολή

# ── Υπογραφή email ────────────────────────────────────────────────────────
EMAIL_SIGNATURE = ''

# ── Browser για Selenium (chrome ή firefox) ───────────────────────────────
BROWSER = 'chrome'

# ── Αρχείο Αδυνατούντων υπό έγκριση ──────────────────────────────────────
ADY_XORIS_EGKRISI_PATH = ''

# -- Παραλήπτες ενημερωτικού email -------------------------------------------
INFORM_RECIPIENTS = ''   # emails παραλητών (ένα ανά γραμμή)


def email_signature():
    """Επιστρέφει την υπογραφή email βάσει των ρυθμίσεων."""
    return EMAIL_SIGNATURE


# ── Θέμα & σώμα email ──────────────────────────────────────────────────────
SUBJECT       = 'Παρόντες με ενεργή μακροχρόνια άδεια'
BODY_TEMPLATE = (
    'Καλημέρα,\n\n'
    'Εντοπίστηκαν εκπαιδευτικοί στο σχολείο {school} που εμφανίζονται ως παρόντες '
    'ενώ βρίσκονται σε μακροχρόνια άδεια (επισυνάπτεται αρχείο).\n\n'
    'Παρακαλούμε για τις ενέργειές σας.\n\n'
    'Με εκτίμηση,\n'
    'Για τη Δ/νση ΠΕ ...,\n'
    'Υπεύθυνος MySchool\n'
    'τηλ. ...'
)


def _load_local():
    """
    Φορτώνει ρυθμίσεις από δύο πηγές:
      1. data/local_settings.json  -> μη-ευαίσθητα (FROM_NAME, FROM_EMAIL, SMTP_HOST κ.λπ.)
      2. Windows Credential Manager -> ευαίσθητα (MYSCHOOL_USER, MYSCHOOL_PASS, FROM_PASSWORD)
    Το JSON ΔΕΝ ανεβαίνει στο GitHub (βλ. .gitignore).
    """
    import json, os, sys

    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        pf   = os.environ.get('PROGRAMFILES',      r'C:\Program Files').lower()
        pf86 = os.environ.get('PROGRAMFILES(X86)', r'C:\Program Files (x86)').lower()
        if exe_dir.lower().startswith(pf) or exe_dir.lower().startswith(pf86):
            base = os.path.join(
                os.environ.get('LOCALAPPDATA', os.path.expanduser('~')),
                'MySchoolChecks')
            os.makedirs(base, exist_ok=True)
        else:
            base = exe_dir
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    # 1. Φόρτωσε μη-ευαίσθητα από JSON
    path = os.path.join(base, 'data', 'local_settings.json')
    if os.path.exists(path):
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            g = globals()
            for k, v in data.items():
                if k in g:
                    g[k] = v
        except Exception:
            pass

    # 2. Φόρτωσε ευαίσθητα από Windows Credential Manager
    try:
        import keyring
        _SENSITIVE = ('MYSCHOOL_USER', 'MYSCHOOL_PASS', 'FROM_PASSWORD')
        _SERVICE   = 'MySchoolChecks'
        g = globals()
        for key in _SENSITIVE:
            val = keyring.get_password(_SERVICE, key)
            if val:
                g[key] = val
    except Exception:
        pass


_load_local()
