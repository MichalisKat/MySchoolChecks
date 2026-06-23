"""
checks/tmimata_genikis.py
═════════════════════════
Έλεγχος τμημάτων γενικής παιδείας / δυναμικού.

5.3 Νηπιαγωγεία:
  Έλεγχος 1: Λειτουργικότητα == Υποχρεωτικά Πρωινά Τμήματα
             (έχουν δημιουργηθεί όσα τμήματα πρέπει;)
  Έλεγχος 2: Λειτουργικότητα == Υποχρεωτικά Πρωινά Τμήματα με Μαθητές
             (έχουν καταχωρηθεί μαθητές σε όλα τα τμήματα;)

5.4 Δημοτικά:
  Έλεγχος 1: Λειτουργικότητα == Τμήματα Γενικής Παιδείας
  Έλεγχος 2: Λειτουργικότητα == Τμήματα Γενικής Παιδείας με μαθητές

Αποτέλεσμα: σχολεία με απόκλιση στον 1ο, 2ο ή και στους δύο ελέγχους.
Email σχολείου: αντλείται από stat2_2 (Εκτεταμένα Στοιχεία Σχολ. Μον.).
"""

import pandas as pd
import config
from core.framework import get_downloaded_file, read_csv_fixed, clean_field

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE       = 'Έλεγχος Τμημάτων Γενικής Παιδείας / Δυναμικού (2026-2027)'
CHECK_DESCRIPTION = 'Αποτύπωση αποκλίσεων Λειτουργικότητας vs Τμημάτων και Τμημάτων με Μαθητές (5.3 & 5.4)'
RESULTS_FOLDER    = 'tmimata_genikis'
HAS_EMAIL         = True
REQUIRED_REPORTS  = [
    '5.3 — Αποτύπωση νηπιαγωγείων',
    '5.4 — Αποτύπωση δημοτικών',
    '2.2 — Εκτεταμένα Στοιχεία Σχολ. Μον.',
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email Σχολείου'

EMAIL_SUBJECT = 'Αποτύπωση τμημάτων γενικής παιδείας στο MySchool'
EMAIL_BODY    = lambda school='': (
    'Καλημέρα,\n\n'
    'Κατά τον έλεγχο των στοιχείων στο MySchool διαπιστώθηκε απόκλιση μεταξύ '
    'της Λειτουργικότητας και των καταχωρημένων τμημάτων / μαθητών για το σχολείο σας '
    '(επισυνάπτεται αναλυτικός πίνακας).\n\n'
    'Παρακαλούμε να ελέγξετε και να διορθώσετε τα στοιχεία στο MySchool το συντομότερο δυνατό.\n\n'
    'Παρακαλούμε για τις ενέργειές σας.\n\n'
    + config.email_signature()
)

COLUMNS = [
    ('Τύπος',                14),
    ('Κωδικός Σχολείου',     16),
    ('Ονομασία Σχολείου',    42),
    ('Email Σχολείου',       32),
    ('Λειτουργικότητα',      14),
    ('Τμήματα',              12),
    ('Τμήματα με Μαθητές',   18),
    ('Μαθητές',              10),
    ('Διαφορά Τμημάτων',     16),
    ('Διαφορά Μαθητών',      16),
    ('Απόκλιση',             20),
]

CENTER_COLS = {
    'Τύπος', 'Κωδικός Σχολείου', 'Λειτουργικότητα',
    'Τμήματα', 'Τμήματα με Μαθητές', 'Μαθητές',
    'Διαφορά Τμημάτων', 'Διαφορά Μαθητών', 'Απόκλιση',
}

HIGHLIGHT_COL    = 'Απόκλιση'
HIGHLIGHT_COLORS = ('E74C3C', 'FADBD8', 'FDEDEC')


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path_53 = get_downloaded_file('5.3', 'Αρχείο 5.3 (Αποτύπωση νηπιαγωγείων):', silent=True)
    path_54 = get_downloaded_file('5.4', 'Αρχείο 5.4 (Αποτύπωση δημοτικών):',    silent=True)
    # stat2_2: για emails — προαιρετικό, αθόρυβα
    path_22 = get_downloaded_file('2.2', 'Αρχείο 2.2 (Εκτεταμένα Στοιχεία):', silent=True)
    return {'path_53': path_53, 'path_54': path_54, 'path_22': path_22}


# ── Βοηθητικές ──────────────────────────────────────────────────────────────
def _to_int(series):
    """Μετατρέπει στήλη σε int (NaN → 0)."""
    return pd.to_numeric(series, errors='coerce').fillna(0).astype(int)


def _build_email_map(path_22):
    """
    Χτίζει dict {κωδικός_σχολείου: email} από stat2_2.
    Στήλες (με 1-column shift): col11=Κωδικός, col18=Email.
    """
    if not path_22:
        return {}
    try:
        import os
        ext = os.path.splitext(path_22)[1].lower()
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(path_22)
        else:
            df = read_csv_fixed(path_22)

        if df.shape[1] <= 18:
            print('  ⚠ stat2_2: λίγες στήλες, emails δεν αντλήθηκαν.')
            return {}

        codes  = clean_field(df.iloc[:, 11]).str.lstrip('0')
        emails = df.iloc[:, 18].astype(str).str.strip()
        result = {}
        for code, email in zip(codes, emails):
            if code and email and email.lower() not in ('nan', ''):
                result[code] = email
        print(f'  ✓ stat2_2: {len(result)} emails φορτώθηκαν.')
        return result
    except Exception as e:
        print(f'  ⚠ stat2_2 σφάλμα: {e}')
        return {}


def _check_file(df, label, col_tmimata, col_mathites, email_map):
    """
    Ελέγχει ένα αρχείο (5.3 ή 5.4) και επιστρέφει DataFrame με αποκλίσεις.
    """
    required = ['Κωδ. Υπουργείου Σχολείου', 'Ονομασία Σχολ. Μονάδας',
                'Λειτουργικότητα', col_tmimata, col_mathites, 'Μαθητές']
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f'  ⚠ [{label}] Λείπουν στήλες: {missing}')
        print(f'  Διαθέσιμες: {list(df.columns)}')
        return pd.DataFrame()

    df = df.copy()
    df['_leit']     = _to_int(df['Λειτουργικότητα'])
    df['_tmimata']  = _to_int(df[col_tmimata])
    df['_mathites'] = _to_int(df[col_mathites])

    dev_t = df['_leit'] != df['_tmimata']
    dev_m = df['_leit'] != df['_mathites']
    mask  = dev_t | dev_m

    if not mask.any():
        print(f'  ✓ [{label}] Καμία απόκλιση.')
        return pd.DataFrame()

    def _apoklisi(row):
        t = row['_leit'] != row['_tmimata']
        m = row['_leit'] != row['_mathites']
        if t and m: return 'Τμήματα & Μαθητές'
        if t:       return 'Τμήματα'
        return 'Μαθητές'

    records = []
    for _, row in df[mask].iterrows():
        code   = str(row['Κωδ. Υπουργείου Σχολείου']).strip()
        diff_t = int(row['_tmimata'])  - int(row['_leit'])
        diff_m = int(row['_mathites']) - int(row['_leit'])
        records.append({
            'Τύπος':               label,
            'Κωδικός Σχολείου':    code,
            'Ονομασία Σχολείου':   str(row['Ονομασία Σχολ. Μονάδας']).strip(),
            'Email Σχολείου':      email_map.get(code.lstrip('0'), ''),
            'Λειτουργικότητα':     int(row['_leit']),
            'Τμήματα':             int(row['_tmimata']),
            'Τμήματα με Μαθητές': int(row['_mathites']),
            'Μαθητές':             int(pd.to_numeric(row['Μαθητές'], errors='coerce') or 0),
            'Διαφορά Τμημάτων':   (f'+{diff_t}' if diff_t > 0 else str(diff_t))
                                    if row['_leit'] != row['_tmimata'] else '—',
            'Διαφορά Μαθητών':    (f'+{diff_m}' if diff_m > 0 else str(diff_m))
                                    if row['_leit'] != row['_mathites'] else '—',
            'Απόκλιση':           _apoklisi(row),
        })

    print(f'  → [{label}] {len(records)} σχολεία με απόκλιση')
    return pd.DataFrame(records)


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    email_map = _build_email_map(ctx.get('path_22'))
    frames = []

    # 5.3 — Νηπιαγωγεία
    if ctx.get('path_53'):
        try:
            df53  = pd.read_excel(ctx['path_53'])
            res53 = _check_file(df53, 'Νηπιαγωγείο',
                                 'Υποχρεωτικά Πρωινά Τμήματα',
                                 'Υποχρεωτικά Πρωινά Τμήματα με Μαθητές',
                                 email_map)
            if not res53.empty:
                frames.append(res53)
        except Exception as e:
            print(f'  ✗ Σφάλμα φόρτωσης 5.3: {e}')
    else:
        print('  ⚠ Δεν βρέθηκε αρχείο 5.3')

    # 5.4 — Δημοτικά
    if ctx.get('path_54'):
        try:
            df54  = pd.read_excel(ctx['path_54'])
            res54 = _check_file(df54, 'Δημοτικό',
                                 'Τμήματα Γενικής Παιδείας',
                                 'Τμήματα Γενικής Παιδείας με μαθητές',
                                 email_map)
            if not res54.empty:
                frames.append(res54)
        except Exception as e:
            print(f'  ✗ Σφάλμα φόρτωσης 5.4: {e}')
    else:
        print('  ⚠ Δεν βρέθηκε αρχείο 5.4')

    if not frames:
        return pd.DataFrame()

    df_out = pd.concat(frames, ignore_index=True)
    return df_out.sort_values(['Τύπος', 'Ονομασία Σχολείου']).reset_index(drop=True)
