"""
checks/arnhtika_ypoloipa.py
═══════════════════════════
Αρνητικά Υπόλοιπα Υποχρεωτικού Διδακτικού Ωραρίου (4.8).
"""

import pandas as pd
import config
from core.framework import ask_file, get_downloaded_file, ask_date_yyyymmdd, read_input, clean_field

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE    = 'Αρνητικά Υπόλοιπα Ωραρίου'
CHECK_DESCRIPTION = 'Εκπαιδευτικοί με αρνητικό υπόλοιπο υποχρεωτικού διδακτικού ωραρίου'
RESULTS_FOLDER = 'arnhtika_ypoloipa'
HAS_EMAIL      = True
REQUIRED_REPORTS = ['4.8 — Ωράριο εκπαιδευτικών']

SCHOOL_TYPES = None  # Φιλτράρισμα βάσει αποκλεισμού "Ιδιωτικό"
COL_YPOLOIPO = 'Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου'

COLUMNS = [
    ('Κωδικός Σχολείου',                                     14),
    ('Ονομασία Σχολείου',                                    38),
    ('Τηλέφωνο',                                             14),
    ('Email',                                                30),
    ('Α.Μ.',                                                 11),
    ('Α.Φ.Μ.',                                               13),
    ('Επώνυμο',                                              18),
    ('Όνομα',                                                16),
    ('Κωδικός Κύριας Ειδικότητας',                           14),
    ('Ώρες Υποχ. Διδακτικού Ωραρίου Υπηρέτησης στο Φορέα',  16),
    ('Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου',            14),
    ('Ισχύει από',                                           13),
    ('Ισχύει έως',                                           13),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email'

CENTER_COLS = {
    'Κωδικός Σχολείου', 'Α.Μ.', 'Α.Φ.Μ.',
    'Κωδικός Κύριας Ειδικότητας',
    'Ώρες Υποχ. Διδακτικού Ωραρίου Υπηρέτησης στο Φορέα',
    'Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου',
    'Ισχύει από', 'Ισχύει έως',
}

COL_MAP = {
    'Κωδικός Σχολείου': 'Κωδικός Σχολείου',
    'Κωδικός Φορέα':    'Κωδικός Σχολείου',
    'Ονομασία Σχολείου':'Ονομασία Σχολείου',
    'Ονομασία Φορέα':   'Ονομασία Σχολείου',
    'Τηλέφωνο':         'Τηλέφωνο',
    'Email':            'Email',
    'Α.Μ.':             'Α.Μ.',
    'Α.Φ.Μ.':           'Α.Φ.Μ.',
    'Επώνυμο':          'Επώνυμο',
    'Όνομα':            'Όνομα',
    'Κωδικός Κύριας Ειδικότητας': 'Κωδικός Κύριας Ειδικότητας',
    'Ώρες Υποχ. Διδακτικού Ωραρίου Υπηρέτησης στο Φορέα': 'Ώρες Υποχ. Διδακτικού Ωραρίου Υπηρέτησης στο Φορέα',
    'Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου': 'Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου',
    'Ισχύει από': 'Ισχύει από',
    'Ισχύει έως': 'Ισχύει έως',
}

EMAIL_SUBJECT = 'Διορθώσεις στο Myschool (αρνητικά υπόλοιπα ωραρίου)'
EMAIL_BODY    = lambda school='': (
    'Καλημέρα σας!\n\n'
    'Έχετε εκπαιδευτικό με αρνητικό υπόλοιπο ωραρίου (αναθέσεις περισσότερες από '
    'το διδακτικό ωράριο), επισυνάπτεται σχετικό αρχείο.\n\n'
    'Παρακαλούμε για τις ενέργειες σας.\n\n'
    + config.email_signature()
)


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path  = get_downloaded_file('4.8', 'Αρχείο 4.8 (Ωράριο εκπαιδευτικών) [csv / xlsx]:', silent=True)
    if path is None:
        return {'path': None, 'today': None}
    today = ask_date_yyyymmdd()
    return {'path': path, 'today': today}


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    df = read_input(ctx['path']).copy()

    # Φίλτρο τύπου σχολείου — εξαιρεί Ιδιωτικά
    type_col = next((c for c in ['Είδος Σχολείου', 'Είδος Φορέα', 'Τύπος Σχολείου']
                     if c in df.columns), None)
    if type_col:
        df = df[~df[type_col].str.contains('Ιδιωτικ', na=False)].copy()
        print(f'  → Μετά φίλτρο τύπου: {len(df)} εγγραφές')

    # Φίλτρο αρνητικού υπολοίπου
    yp_col = COL_YPOLOIPO
    if yp_col not in df.columns:
        matches = [c for c in df.columns if 'υπόλοιπο' in c.lower() and 'ωράριο' in c.lower()]
        yp_col  = matches[0] if matches else None
    if not yp_col:
        print('  ✗ Δεν βρέθηκε στήλη υπολοίπου.')
        return pd.DataFrame()

    df[yp_col] = (df[yp_col].astype(str)
                  .str.replace('="', '', regex=False).str.replace('"', '', regex=False)
                  .str.replace(',', '.', regex=False).str.strip())
    df['_num'] = pd.to_numeric(df[yp_col], errors='coerce')
    df = df[df['_num'] < 0].copy()
    print(f'  → Μετά φίλτρο αρνητικού: {len(df)} εγγραφές')

    if df.empty:
        return pd.DataFrame()

    # Κατασκευή εξόδου
    results = []
    for _, row in df.iterrows():
        rec = {}
        for src, dst in COL_MAP.items():
            # Αν έχει ήδη τιμή για αυτό το dst, μην την αντικαταστήσεις με κενό
            if rec.get(dst, ''):
                continue
            val = ''
            if src in df.columns:
                val = row.get(src, '')
            else:
                for c in df.columns:
                    if c.strip().lower() == src.strip().lower():
                        val = row.get(c, '')
                        break
            if isinstance(val, str):
                val = val.replace('="', '').replace('"', '').strip()
            if val:
                rec[dst] = val
        results.append(rec)

    df_out = pd.DataFrame(results)
    for col in [c[0] for c in COLUMNS]:
        if col not in df_out.columns:
            df_out[col] = ''
    df_out = df_out[[c[0] for c in COLUMNS]]
    return df_out.sort_values([SCHOOL_COLUMN, 'Επώνυμο', 'Όνομα']).reset_index(drop=True)


def test_body(df_out, today, schools):
    neg = df_out[df_out['Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου'].apply(
        lambda x: str(x).replace('="','').replace('"','').strip()
    ).str.lstrip('-').str.isnumeric() if 'Υπόλοιπο Υποχρεωτικού Διδακτικού Ωραρίου' in df_out.columns else df_out]
    return (
        f'Σύνοψη ελέγχου αρνητικών υπολοίπων ωραρίου — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Βρέθηκαν: {len(df_out)} εκπαιδευτικοί με αρνητικό υπόλοιπο\n'
        f'Σχολεία που εμφανίζονται ({len(schools)}): {", ".join(sorted(schools))}'
    )