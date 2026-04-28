"""
checks/adies.py
═══════════════
Σύγκριση 4.21 (Άδειες) vs 4.9 (Παρόντες).
Εντοπίζει εκπαιδευτικούς σε ενεργή μακροχρόνια άδεια που εμφανίζονται ως παρόντες.
"""

from datetime import datetime
import pandas as pd
import config
from dateutil.relativedelta import relativedelta
from core.framework import ask_file, get_downloaded_file, ask_date_yyyymmdd, read_csv_fixed

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE    = 'Σύγκριση Αδειών (πλην Άνευ Αποδοχών) & Παρόντων'
CHECK_DESCRIPTION = 'Εκπαιδευτικοί σε ενεργή μακροχρόνια άδεια (πλην άνευ αποδοχών) που εμφανίζονται ως παρόντες'
RESULTS_FOLDER = 'adies_plhn_aa'
HAS_EMAIL      = True
REQUIRED_REPORTS = ['4.21 — Άδειες εκπαιδευτικών', '4.9 — Παρόντες εκπαιδευτικοί']

COLUMNS = [
    ('Κωδικός Σχολείου',  14),
    ('Ονομασία Σχολείου', 38),
    ('Email Σχολείου',    30),
    ('ΑΜ',                11),
    ('ΑΦΜ',               13),
    ('Επώνυμο',           18),
    ('Τύπος άδειας',      42),
    ('Κατάσταση άδειας',  18),
    ('ΑΠΟ',               12),
    ('ΕΩΣ',               12),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email Σχολείου'

EMAIL_SUBJECT = 'Παρόντες με ενεργή μακροχρόνια άδεια'
EMAIL_BODY    = lambda school='': (
    'Καλημέρα,\n\n'
    'Εντοπίστηκαν εκπαιδευτικοί στο σχολείο σας που εμφανίζονται ως παρόντες '
    'ενώ βρίσκονται σε μακροχρόνια άδεια (επισυνάπτεται αρχείο).\n\n'
    'Να καταχωρηθεί η σχετική άδεια και ως απουσία στην τοποθέτηση (κοκκίνισμα ονόματος).\n\n'
    'Παρακαλούμε για τις ενέργειές σας.\n\n'
    + config.email_signature()
)

CENTER_COLS    = {'ΑΜ', 'ΑΦΜ', 'ΑΠΟ', 'ΕΩΣ', 'Κωδικός Σχολείου', 'Κατάσταση άδειας'}
STATUS_COLORS  = {
    '3-Εγκρίθηκε':     ('FFE2E2', 'FFEEEE'),
    '2-Υποβλήθηκε':    ('FFF2CC', 'FFF8E1'),
    '1-Δημιουργήθηκε': ('FFF2CC', 'FFF8E1'),
}
STATUS_COL = 'Κατάσταση άδειας'

# ── Παράμετροι λογικής ──────────────────────────────────────────────────────
EXCLUDED_STATUS = {'5-Ανακλήθηκε', '4-Απορρίφθηκε'}
MIN_DAYS        = 10


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path_421 = get_downloaded_file('4.21', 'Αρχείο 4.21 (Άδειες) [.csv]:', csv_only=True)
    path_49  = get_downloaded_file('4.9',  'Αρχείο 4.9 (Παρόντες) [.csv]:', csv_only=True)
    if path_421 is None or path_49 is None:
        return {'path_421': path_421, 'path_49': path_49, 'today': None}
    today    = ask_date_yyyymmdd()
    return {'path_421': path_421, 'path_49': path_49, 'today': today}


# ── Λογική ──────────────────────────────────────────────────────────────────
def _parse_date(series):
    return pd.to_datetime(series, format='%d/%m/%Y', errors='coerce')

def _calc_end(row):
    apo = row['Από_dt']
    if pd.isna(apo): return pd.NaT
    try:
        days   = int(float(row['Εγκρ. Ημέρες'])) if pd.notna(row.get('Εγκρ. Ημέρες')) else 0
        months = int(float(row['Εγκρ. Μήνες']))  if pd.notna(row.get('Εγκρ. Μήνες'))  else 0
        years  = int(float(row['Εγκρ. Έτη']))    if pd.notna(row.get('Εγκρ. Έτη'))    else 0
    except (ValueError, TypeError):
        return pd.NaT
    if days == 0 and months == 0 and years == 0:
        return pd.NaT
    return apo + relativedelta(years=years, months=months, days=days - 1)

def _clean(series):
    return series.astype(str).str.replace('="', '').str.replace('"', '').str.strip()

def process(ctx):
    today    = ctx['today']
    df21     = read_csv_fixed(ctx['path_421'])
    df9      = read_csv_fixed(ctx['path_49'])

    # Φόρτωση 4.21
    df21['ΑΦΜ_clean']  = _clean(df21['ΑΦΜ'])
    df21['Από_dt']     = _parse_date(df21['Από'])
    df21['Έως_dt']     = df21.apply(_calc_end, axis=1)
    df21['Διάρκεια']   = (df21['Έως_dt'] - df21['Από_dt']).dt.days + 1
    df21['Σχολείο_21'] = df21['Φορέας Δημιουργίας/Υποβολής'].astype(str).str.strip()

    # Φόρτωση 4.9
    df9['ΑΦΜ_clean'] = _clean(df9['ΑΦΜ'])
    df9['Σχολείο_9'] = df9['Ονομασία Σχολείου'].astype(str).str.strip()
    df9['Κωδικός_9'] = _clean(df9['Κωδικός Σχολείου'])
    df9['Email_9']   = df9['Email'].astype(str).str.strip()

    # Ενεργές μακροχρόνιες άδειες
    mask = (
        (df21['Από_dt'] <= today) & (df21['Έως_dt'] >= today) &
        (~df21['Κατάσταση άδειας'].isin(EXCLUDED_STATUS)) &
        (df21['Διάρκεια'] > MIN_DAYS)
    )
    active = df21[mask].copy()
    print(f'  → {len(active)} μακροχρόνιες άδειες ενεργές')

    # Τομή με παρόντες
    set_49 = set(zip(df9['ΑΦΜ_clean'], df9['Σχολείο_9']))
    problem = active[active.apply(
        lambda r: (r['ΑΦΜ_clean'], r['Σχολείο_21']) in set_49, axis=1
    )].copy()
    print(f'  → {len(problem)} παρόντες ΚΑΙ σε άδεια')

    if problem.empty:
        return pd.DataFrame()

    # Κατασκευή εξόδου
    school_info = (
        df9.drop_duplicates(subset=['Σχολείο_9'])
        .set_index('Σχολείο_9')[['Κωδικός_9', 'Email_9']]
    )
    records = []
    for _, row in problem.iterrows():
        sn   = row['Σχολείο_21']
        info = school_info.loc[sn] if sn in school_info.index else pd.Series()
        records.append({
            'Κωδικός Σχολείου': info.get('Κωδικός_9', ''),
            'Ονομασία Σχολείου': sn,
            'Email Σχολείου':   info.get('Email_9', ''),
            'ΑΜ':               _clean(pd.Series([row['ΑΜ']])).iloc[0],
            'ΑΦΜ':              row['ΑΦΜ_clean'],
            'Επώνυμο':          row['Επώνυμο'],
            'Τύπος άδειας':     row['Τύπος άδειας'],
            'Κατάσταση άδειας': row['Κατάσταση άδειας'],
            'ΑΠΟ':              row['Από_dt'].strftime('%d/%m/%Y') if pd.notna(row['Από_dt']) else '',
            'ΕΩΣ':              row['Έως_dt'].strftime('%d/%m/%Y') if pd.notna(row['Έως_dt']) else '',
        })

    df_out = pd.DataFrame(records)
    return df_out.sort_values(['Ονομασία Σχολείου', 'Επώνυμο'])


def test_body(df_out, today, schools):
    egk  = (df_out['Κατάσταση άδειας'] == '3-Εγκρίθηκε').sum() if 'Κατάσταση άδειας' in df_out.columns else 0
    ypo  = (df_out['Κατάσταση άδειας'].isin(['2-Υποβλήθηκε','1-Δημιουργήθηκε'])).sum() if 'Κατάσταση άδειας' in df_out.columns else 0
    return (
        f'Σύνοψη ελέγχου αδειών (πλην ΑΑ) vs Παρόντων — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Παρόντες ΚΑΙ σε άδεια: {len(df_out)} εκπαιδευτικοί σε {len(schools)} σχολεία\n'
        f'  Εγκρίθηκε: {egk}\n'
        f'  Υποβλήθηκε/Δημιουργήθηκε: {ypo}\n'
        f'Σχολεία: {", ".join(sorted(str(s) for s in schools))}\n'
    )
