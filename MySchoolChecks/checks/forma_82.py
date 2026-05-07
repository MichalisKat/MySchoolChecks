"""
checks/forma_82.py
══════════════════
Επιβεβαίωση Δεδομένων Σχολείων (8.2).
Εντοπίζει σχολεία που δεν έχουν επιβεβαιώσει τα δεδομένα τους πριν από cutoff ημερομηνία.
"""

import pandas as pd
from datetime import datetime, date
import config
from core.framework import ask_file, get_downloaded_file

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE    = 'Επιβεβαίωση Δεδομένων Σχολείων'
CHECK_DESCRIPTION = 'Σχολεία που δεν έχουν ολοκληρώσει επιβεβαίωση δεδομένων'
RESULTS_FOLDER = 'epivevaiosi_dedomenon'
HAS_EMAIL      = True
REQUIRED_REPORTS = ['8.2 — Επιβεβαίωση δεδομένων σχολείων']

EIDH        = [
    'Δημοτικά Σχολεία',
    'Νηπιαγωγεία',
    'Γυμνάσια',
    'Ειδικής Επαγγελματικής Εκπαίδευσης και Κατάρτισης',
    'Επαγγελματικά Λύκεια',
    'Λύκεια',
    'Σχολικό Εργαστηριακό Κέντρο',
]
SRC_DATE_COL = 'Τελευταία Ενημέρωση Φόρμας Επιβεβαίωση Δεδομένων'

COLUMNS = [
    ('Κωδικός Υπ. Σχολ.',              16),
    ('Ονομασία Σχολείου',               40),
    ('Τηλέφωνο',                        16),
    ('Email',                           32),
    ('Τελευταία Επιβεβαίωση Δεδομένων', 30),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email'

CENTER_COLS = {'Κωδικός Υπ. Σχολ.', 'Τηλέφωνο', 'Τελευταία Επιβεβαίωση Δεδομένων'}

# ── Υπολογισμός cutoff ──────────────────────────────────────────────────────
def _auto_cutoff():
    """Επιστρέφει την πλησιέστερη προηγούμενη 1η ή 15η του τρέχοντος μήνα.
    Από 2-14: επιστρέφει 1η του τρέχοντος μήνα.
    Από 16-31: επιστρέφει 15η του τρέχοντος μήνα.
    Την 1η και 15η: επιστρέφει την ίδια ημέρα.
    """
    today = date.today()
    if today.day <= 14:
        return datetime(today.year, today.month, 1)
    else:
        return datetime(today.year, today.month, 15)

def _cutoff_str():
    """Επιστρέφει το cutoff ως string π.χ. '1/4/2026' ή '15/4/2026'."""
    c = _auto_cutoff()
    return f'{c.day}/{c.month}/{c.year}'

EMAIL_SUBJECT = f'Υπενθύμιση: Επιβεβαίωση δεδομένων στο myschool ({_cutoff_str()})'
EMAIL_BODY    = lambda school='': (
    f'Καλημέρα σας!\n\n'
    f'Για τους παραλήπτες του παρόντος δεν έχει γίνει επιβεβαίωση δεδομένων '
    f'στο myschool για υπό εξέταση ημερομηνία {_cutoff_str()}. '
    f'Παρακαλούμε για τις ενέργειες σας.\n\n'
    + config.email_signature()
)


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path   = get_downloaded_file('8.2', 'Αρχείο 8.2 [xls / xlsx]:', silent=True)
    cutoff = _auto_cutoff()
    print(f'  ✓ Cutoff: {cutoff.strftime("%d/%m/%Y")} (υπολογίστηκε αυτόματα)')
    return {'path': path, 'today': cutoff, 'cutoff': cutoff}


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    cutoff = ctx['cutoff']
    df     = pd.read_excel(ctx['path'])
    print(f'  ✓ {len(df)} γραμμές φορτώθηκαν')

    df = df[df['Είδος Σχολείου'].isin(EIDH)].copy()
    print(f'  → {len(df)} μετά φίλτρο Είδους Σχολείου')

    df = df[df['Αναστολή'].str.strip() == 'Όχι'].copy()
    print(f'  → {len(df)} μετά φίλτρο Αναστολής')

    df[SRC_DATE_COL] = pd.to_datetime(df[SRC_DATE_COL])
    df = df[df[SRC_DATE_COL] < cutoff].copy()
    print(f'  → {len(df)} με τελευταία επιβεβαίωση πριν {cutoff.strftime("%d/%m/%Y")}')

    if df.empty:
        return pd.DataFrame()

    df_out = pd.DataFrame({
        'Κωδικός Υπ. Σχολ.'             : df['Κωδικός Υπ. Σχολ.'].astype(str),
        'Ονομασία Σχολείου'              : df['Ονομασία Σχολείου'],
        'Τηλέφωνο'                       : df['Τηλέφωνο'].astype(str),
        'Email'                          : df['Email'],
        'Τελευταία Επιβεβαίωση Δεδομένων': df[SRC_DATE_COL].dt.strftime('%d/%m/%Y %H:%M'),
    }).sort_values('Ονομασία Σχολείου').reset_index(drop=True)

    return df_out


def test_body(df_out, today, schools):
    date_col = 'Τελευταία Επιβεβαίωση Δεδομένων'
    oldest = newest = ''
    if date_col in df_out.columns and len(df_out) > 0:
        dates = df_out[date_col].dropna()
        if len(dates):
            oldest = dates.min()
            newest = dates.max()
    return (
        f'Σύνοψη ελέγχου επιβεβαίωσης δεδομένων — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Βρέθηκαν: {len(df_out)} σχολεία χωρίς επιβεβαίωση\n'
        f'Παλαιότερη: {oldest}\n'
        f'Νεότερη: {newest}\n'
        f'Σχολεία που εμφανίζονται ({len(schools)}): {", ".join(sorted(schools))}'
    )