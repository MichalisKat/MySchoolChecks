"""
checks/orario_diafora.py
════════════════════════
Διαφορά AK-AL στο αρχείο 4.9.
Εντοπίζει εκπαιδευτικούς όπου το υποχρεωτικό ωράριο διαφέρει από το ωράριο στους φορείς.
"""

import sys
import pandas as pd
import config
from core.framework import ask_file, get_downloaded_file, ask_date_yyyymmdd, read_input

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE    = 'Διαφορές AK-AL'
CHECK_DESCRIPTION = 'Διαφορά μεταξύ υποχρεωτικού ωραρίου και ωραρίου στους φορείς'
RESULTS_FOLDER = 'diafores_ak_al'
HAS_EMAIL      = True
REQUIRED_REPORTS = ['4.9 — Παρόντες εκπαιδευτικοί']
TEST_ONLY      = True   # Μόνο test mode — δεν αφορά σχολεία

COLUMNS = [
    ('Κωδικός Σχολείου',  14),
    ('Ονομασία Σχολείου', 38),
    ('Email',             30),
    ('ΑΜ',                11),
    ('ΑΦΜ',               13),
    ('Επώνυμο',           18),
    ('Όνομα',             16),
    ('Κωδ. Ειδικότητας',  14),
    ('Υποχρεωτικό Ωράριο (AK)',   14),
    ('Ωράριο στους Φορείς (AL)',   14),
    ('AK-AL',              9),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email'
CENTER_COLS   = {'ΑΜ', 'ΑΦΜ', 'Κωδικός Σχολείου', 'Κωδ. Ειδικότητας',
                 'Υποχρεωτικό Ωράριο (AK)', 'Ωράριο στους Φορείς (AL)', 'AK-AL'}

EMAIL_SUBJECT = 'Διαφορες AK-AL - 4.9'
EMAIL_BODY    = lambda school='': (
    'Επισυνάπτεται αρχείο με εκπαιδευτικούς για τους οποίους '
    'υπάρχει διαφορά μεταξύ υποχρεωτικού ωραρίου (AK) '
    'και ωραρίου στους φορείς (AL).\n\n'
    + config.email_signature()
)


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path  = get_downloaded_file('4.9', 'Αρχείο 4.9 [csv / xlsx]:')
    if path is None:
        return {'path': None, 'today': None}
    today = ask_date_yyyymmdd()
    return {'path': path, 'today': today}


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    df   = read_input(ctx['path'])
    cols = list(df.columns)
    n    = len(cols)

    idx_AK, idx_AL = 36, 37
    if n <= idx_AL:
        print(f'  ✗ Το αρχείο έχει {n} στήλες — αναμένονται ≥38.')
        return pd.DataFrame()

    col_AK = cols[idx_AK]
    col_AL = cols[idx_AL]
    print(f'  Στήλη AK ({idx_AK}): "{col_AK}"')
    print(f'  Στήλη AL ({idx_AL}): "{col_AL}"')

    def clean_num(s):
        return pd.to_numeric(
            s.astype(str).str.replace('="', '', regex=False).str.replace('"', '', regex=False).str.strip(),
            errors='coerce'
        ).fillna(0)

    df['_AK']   = clean_num(df[col_AK])
    df['_AL']   = clean_num(df[col_AL])
    df['_DIFF'] = df['_AK'] - df['_AL']
    df          = df[df['_DIFF'] != 0].copy()

    if df.empty:
        return pd.DataFrame()

    def g(row, idx):
        val = str(row.iloc[idx]).strip().replace('="', '').replace('"', '')
        return val if val not in ('', 'nan', 'None') else ''

    records = []
    for _, row in df.iterrows():
        diff = int(row['_DIFF'])
        records.append({
            'Κωδικός Σχολείου':          g(row,  7),
            'Ονομασία Σχολείου':         g(row,  8),
            'Email':                     g(row, 11),
            'ΑΜ':                        g(row, 16),
            'ΑΦΜ':                       g(row, 17),
            'Επώνυμο':                   g(row, 19),
            'Όνομα':                     g(row, 20),
            'Κωδ. Ειδικότητας':          g(row, 28),
            'Υποχρεωτικό Ωράριο (AK)':  int(row['_AK']),
            'Ωράριο στους Φορείς (AL)':  int(row['_AL']),
            'AK-AL':                     diff,
        })

    df_out = pd.DataFrame(records)
    return df_out.sort_values(['Ονομασία Σχολείου', 'Επώνυμο'])


def test_body(df_out, today, schools):
    pos = (df_out['AK-AL'] > 0).sum() if 'AK-AL' in df_out.columns else 0
    neg = (df_out['AK-AL'] < 0).sum() if 'AK-AL' in df_out.columns else 0
    return (
        f'Σύνοψη ελέγχου διαφορών AK-AL — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Βρέθηκαν: {len(df_out)} εγγραφές με διαφορά ≠ 0\n'
        f'Θετικές διαφορές (AK > AL): {pos}\n'
        f'Αρνητικές διαφορές (AK < AL): {neg}\n'
        f'Σχολεία που εμφανίζονται ({len(schools)}): {", ".join(sorted(str(s) for s in schools))}\n'
    )
