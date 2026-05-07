"""
checks/analipsi.py
══════════════════
Ελλιπή Στοιχεία Πράξης Ανάληψης (4.8).
Εκπαιδευτικοί χωρίς Ημερομηνία Ανάληψης, ενεργής τοποθέτησης.
"""

import pandas as pd
import config
from core.framework import ask_file, get_downloaded_file, ask_date_yyyymmdd, read_input, clean_field

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE    = 'Ελλιπή Στοιχεία Πράξης Ανάληψης'
CHECK_DESCRIPTION = 'Εκπαιδευτικοί χωρίς στοιχεία πράξης ανάληψης στο myschool'
RESULTS_FOLDER = 'elliphi_stoixeia_analipsis'
HAS_EMAIL      = True
REQUIRED_REPORTS = ['4.8 — Ωράριο εκπαιδευτικών']

COLUMNS = [
    ('Κωδικός Σχολείου',           14),
    ('Ονομασία Σχολείου',          38),
    ('Τηλέφωνο',                   14),
    ('Email',                      30),
    ('Α.Μ.',                       11),
    ('Α.Φ.Μ.',                     13),
    ('Επώνυμο',                    18),
    ('Όνομα',                      14),
    ('Κωδικός Κύριας Ειδικότητας', 10),
    ('Σχέση Τοποθέτησης',          28),
    ('Ισχύει από',                 13),
    ('Ισχύει έως',                 13),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email Σχολείου'

CENTER_COLS = {
    'Α.Μ.', 'Α.Φ.Μ.', 'Κωδικός Σχολείου',
    'Κωδικός Κύριας Ειδικότητας', 'Ισχύει από', 'Ισχύει έως', 'Τηλέφωνο'
}

EMAIL_SUBJECT = 'Διορθώσεις myschool — Εκπαιδευτικοί χωρίς στοιχεία πράξης ανάληψης'
EMAIL_BODY    = lambda school='': (
    'Καλή σας μέρα,\n\n'
    'Υπάρχουν εκπαιδευτικοί (επισυνάπτεται αρχείο) για τους οποίους δεν έχουν '
    'συμπληρωθεί στο myschool τα πεδία "Στοιχεία Πράξης Ανάληψης" στα Στοιχεία '
    'υπηρέτησης της καρτέλας τοποθέτησης.\n\n'
    'Παρακαλούμε να ελεγχθούν και να συμπληρωθούν τα παραπάνω πεδία αν ανέλαβαν υπηρεσία.\n\n'
    + config.email_signature()
)


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path  = get_downloaded_file('4.8', 'Αρχείο 4.8 [csv / xlsx]:', silent=True)
    if path is None:
        return {'path': None, 'today': None}
    today = ask_date_yyyymmdd()
    return {'path': path, 'today': today}


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    today = ctx['today']
    df    = read_input(ctx['path'])
    df.columns = [c.strip() for c in df.columns]

    df = df[~df['Είδος Σχολείου'].str.contains('Ιδιωτικ', na=False)].copy()
    df = df[df['Διευθυντής Σχολείου'] == 'Όχι'].copy()
    df = df[df['Σχέση Τοποθέτησης'] != 'Υπερωριακά'].copy()

    df['Ημερομηνία Ανάληψης'] = clean_field(df['Ημερομηνία Ανάληψης'])
    df = df[df['Ημερομηνία Ανάληψης'].isin(['', 'nan', 'None', '-'])].copy()

    df['_eos'] = pd.to_datetime(clean_field(df['Ισχύει έως']), format='%d/%m/%Y', errors='coerce')
    df = df[df['_eos'] > today].copy()

    if df.empty:
        return pd.DataFrame()

    df = df.sort_values(['Κωδικός Σχολείου', 'Επώνυμο'])

    out = pd.DataFrame({
        'Κωδικός Σχολείου':           clean_field(df['Κωδικός Σχολείου']),
        'Ονομασία Σχολείου':          clean_field(df['Ονομασία Σχολείου']),
        'Τηλέφωνο':                   clean_field(df['Τηλέφωνο']),
        'Email':                      clean_field(df['Email']),
        'Email Σχολείου':             clean_field(df['Email']),
        'Α.Μ.':                       clean_field(df['Α.Μ.']),
        'Α.Φ.Μ.':                     clean_field(df['Α.Φ.Μ.']),
        'Επώνυμο':                    clean_field(df['Επώνυμο']),
        'Όνομα':                      clean_field(df['Όνομα']),
        'Κωδικός Κύριας Ειδικότητας': clean_field(df['Κωδικός Κύριας Ειδικότητας']),
        'Σχέση Τοποθέτησης':          clean_field(df['Σχέση Τοποθέτησης']),
        'Ισχύει από':                 clean_field(df['Ισχύει από']),
        'Ισχύει έως':                 clean_field(df['Ισχύει έως']),
    })
    return out


def test_body(df_out, today, schools):
    return (
        f'Σύνοψη ελέγχου πράξεων ανάληψης — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Βρέθηκαν: {len(df_out)} εκπαιδευτικοί χωρίς στοιχεία ανάληψης\n'
        f'Σχολεία που εμφανίζονται ({len(schools)}): {", ".join(sorted(schools))}'
    )