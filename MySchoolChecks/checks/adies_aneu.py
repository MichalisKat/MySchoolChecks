"""
checks/adies_aneu.py
════════════════════
Σύγκριση 4.20 (Άδειες Άνευ Αποδοχών) vs 4.9 (Παρόντες).
Εντοπίζει εκπαιδευτικούς σε ενεργή άδεια άνευ αποδοχών που εμφανίζονται ως παρόντες.
"""

from datetime import datetime
import pandas as pd
import config
from core.framework import ask_file, get_downloaded_file, ask_date_yyyymmdd, read_csv_fixed

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE       = 'Σύγκριση Αδειών Άνευ Αποδοχών & Παρόντων'
CHECK_DESCRIPTION = 'Εκπαιδευτικοί σε ενεργή άδεια άνευ αποδοχών που εμφανίζονται ως παρόντες'
RESULTS_FOLDER    = 'adies_aneu_apodoxon'
HAS_EMAIL         = True
REQUIRED_REPORTS  = ['4.20 — Άδειες άνευ αποδοχών', '4.9 — Παρόντες εκπαιδευτικοί']
TEST_ONLY         = True   # Μόνο test mode — δεν αφορά σχολεία

COLUMNS = [
    ('Κωδικός Σχολείου',  14),
    ('Ονομασία Σχολείου', 38),
    ('Email Σχολείου',    30),
    ('ΑΜ',                11),
    ('ΑΦΜ',               13),
    ('Επώνυμο',           18),
    ('Αιτιολογία',        48),
    ('ΑΠΟ',               12),
    ('ΕΩΣ',               12),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email Σχολείου'

CENTER_COLS = {'ΑΜ', 'ΑΦΜ', 'ΑΠΟ', 'ΕΩΣ', 'Κωδικός Σχολείου'}

EMAIL_SUBJECT = 'Παρόντες με ενεργή άδεια άνευ αποδοχών'
EMAIL_BODY    = lambda school='': (
    'Καλημέρα,\n\n'
    'Εντοπίστηκαν εκπαιδευτικοί στο σχολείο σας που εμφανίζονται ως παρόντες '
    'ενώ βρίσκονται σε άδεια άνευ αποδοχών (επισυνάπτεται αρχείο).\n\n'
    'Να καταχωρηθεί η σχετική άδεια και ως απουσία στην τοποθέτηση (κοκκίνισμα ονόματος).\n\n'
    'Παρακαλούμε για τις ενέργειές σας.\n\n'
    + config.email_signature()
)

MIN_DAYS = 10


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path_420 = get_downloaded_file('4.20', 'Αρχείο 4.20 (Άδειες Άνευ Αποδοχών) [.csv]:', csv_only=True)
    path_49  = get_downloaded_file('4.9',  'Αρχείο 4.9 (Παρόντες) [.csv]:', csv_only=True)
    if path_420 is None or path_49 is None:
        return {'path_420': path_420, 'path_49': path_49, 'today': None}
    today    = ask_date_yyyymmdd()
    return {'path_420': path_420, 'path_49': path_49, 'today': today}


# ── Λογική ──────────────────────────────────────────────────────────────────
def _parse_date(series):
    return pd.to_datetime(series, format='%d/%m/%Y', errors='coerce')

def _clean(series):
    return series.astype(str).str.replace('="', '').str.replace('"', '').str.strip()

def process(ctx):
    today = ctx['today']
    df20  = read_csv_fixed(ctx['path_420'])
    df9   = read_csv_fixed(ctx['path_49'])

    # Φόρτωση 4.20
    df20['ΑΦΜ_clean']  = _clean(df20['ΑΦΜ'])
    df20['Από_dt']     = _parse_date(df20['Ισχύει από'])
    df20['Έως_dt']     = _parse_date(df20['Ισχύει έως'])
    df20['Διάρκεια']   = (df20['Έως_dt'] - df20['Από_dt']).dt.days + 1
    df20['Σχολείο_20'] = df20['Φορέας'].astype(str).str.strip()

    # Φόρτωση 4.9
    df9['ΑΦΜ_clean'] = _clean(df9['ΑΦΜ'])
    df9['Σχολείο_9'] = df9['Ονομασία Σχολείου'].astype(str).str.strip()
    df9['Κωδικός_9'] = _clean(df9['Κωδικός Σχολείου'])
    df9['Email_9']   = df9['Email'].astype(str).str.strip()

    # Ενεργές άδειες σήμερα με διάρκεια > MIN_DAYS
    mask = (
        (df20['Από_dt'] <= today) &
        (df20['Έως_dt'] >= today) &
        (df20['Διάρκεια'] > MIN_DAYS)
    )
    active = df20[mask].copy()
    print(f'  → {len(active)} ενεργές άδειες άνευ αποδοχών')

    # Τομή με παρόντες — μόνο βάσει ΑΦΜ
    # (Στο 4.20 ο Φορέας είναι η Διεύθυνση, όχι το σχολείο —
    #  οποιαδήποτε παρουσία στο 4.9 για εκπαιδευτικό σε άδεια άνευ είναι πρόβλημα)
    set_afm_49 = set(df9['ΑΦΜ_clean'])
    problem    = active[active['ΑΦΜ_clean'].isin(set_afm_49)].copy()
    print(f'  → {len(problem)} παρόντες ΚΑΙ σε άδεια άνευ αποδοχών')

    if problem.empty:
        return pd.DataFrame()

    # Πληροφορίες σχολείου από 4.9 (με βάση ΑΦΜ)
    afm_to_school = df9.drop_duplicates(subset=['ΑΦΜ_clean']).set_index('ΑΦΜ_clean')[
        ['Σχολείο_9', 'Κωδικός_9', 'Email_9']
    ]
    records = []
    for _, row in problem.iterrows():
        afm  = row['ΑΦΜ_clean']
        info = afm_to_school.loc[afm] if afm in afm_to_school.index else pd.Series()
        records.append({
            'Κωδικός Σχολείου': info.get('Κωδικός_9', ''),
            'Ονομασία Σχολείου': info.get('Σχολείο_9', ''),
            'Email Σχολείου':   info.get('Email_9', ''),
            'ΑΜ':               _clean(pd.Series([row['ΑΜ']])).iloc[0],
            'ΑΦΜ':              afm,
            'Επώνυμο':          row['Επώνυμο'],
            'Αιτιολογία':       row.get('Αιτιολογία Χρονικού Διαστήματος', ''),
            'ΑΠΟ':              row['Από_dt'].strftime('%d/%m/%Y') if pd.notna(row['Από_dt']) else '',
            'ΕΩΣ':              row['Έως_dt'].strftime('%d/%m/%Y') if pd.notna(row['Έως_dt']) else '',
        })

    df_out = pd.DataFrame(records)
    return df_out.sort_values(['Ονομασία Σχολείου', 'Επώνυμο'])


def test_body(df_out, today, schools):
    return (
        f'Σύνοψη ελέγχου αδειών ΑΑ vs Παρόντων — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Παρόντες ΚΑΙ σε άδεια άνευ αποδοχών: {len(df_out)} εκπαιδευτικοί σε {len(schools)} σχολεία\n'
        f'Σχολεία: {", ".join(sorted(str(s) for s in schools))}\n'
    )
