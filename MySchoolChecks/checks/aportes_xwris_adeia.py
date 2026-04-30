"""
checks/aportes_xwris_adeia.py
══════════════════════════════
Απουσία στην Τοποθέτηση χωρίς Δήλωση Άδειας (4.16 + 4.21).
Εντοπίζει εκπαιδευτικούς που έχουν καταχωρηθεί ως Μακροχρόνια Απόντες
στο 4.16 (εξαιρουμένων Άνευ Αποδοχών) χωρίς ενεργή άδεια στο 4.21.
"""

from datetime import datetime
import pandas as pd
import config
from dateutil.relativedelta import relativedelta
from core.framework import get_downloaded_file, ask_date_yyyymmdd, read_csv_fixed, clean_field

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE       = 'Απουσία χωρίς Δήλωση Άδειας'
CHECK_DESCRIPTION = 'Εκπαιδευτικοί με Μακροχρόνια Απουσία χωρίς ενεργή άδεια στο myschool'
RESULTS_FOLDER    = 'aportes_xwris_adeia'
HAS_EMAIL         = True
REQUIRED_REPORTS  = ['4.16 — Αιτιολόγηση απουσίας', '4.21 — Άδειες (πλην ΑΑ)']

COLUMNS = [
    ('Κωδικός Σχολείου',           14),
    ('Ονομασία Σχολείου',          38),
    ('Τηλέφωνο',                   14),
    ('Email',                      30),
    ('Α.Μ.',                       11),
    ('Α.Φ.Μ.',                     13),
    ('Επώνυμο',                    18),
    ('Όνομα',                      16),
    ('Κωδικός Κύριας Ειδικότητας', 14),
    ('Εξειδ. Αιτιολόγησης',        40),
    ('Από',                        12),
    ('Έως',                        12),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email Σχολείου'

CENTER_COLS = {
    'Α.Μ.', 'Α.Φ.Μ.', 'Κωδικός Σχολείου',
    'Κωδικός Κύριας Ειδικότητας', 'Τηλέφωνο', 'Από', 'Έως'
}

EMAIL_SUBJECT = 'Διορθώσεις myschool — Εκπαιδευτικοί με απουσία χωρίς καταχωρημένη άδεια'
EMAIL_BODY    = lambda school='': (
    'Καλή σας μέρα,\n\n'
    'Εντοπίστηκαν εκπαιδευτικοί στο σχολείο σας (επισυνάπτεται αρχείο) για τους '
    'οποίους έχει καταχωρηθεί Μακροχρόνια Απουσία στο myschool, χωρίς να υπάρχει '
    'αντίστοιχη ενεργή άδεια.\n\n'
    'Παρακαλούμε να ελεγχθεί αν η άδεια έχει καταχωρηθεί ή αν χρειάζεται '
    'διόρθωση της απουσίας.\n\n'
    'Για λεπτομέρειες επικοινωνείτε με τα τμήματα αδειών της Δ/νσης '
    '(μονίμων ή αναπληρωτών κατά περίπτωση).\n\n'
    + config.email_signature()
)

EXCLUDED_STATUS      = {'5-Ανακλήθηκε', '4-Απορρίφθηκε'}
EXCLUDED_SPECIALTIES = {'ΠΕ23-ΣΔΕΥ', 'ΠΕ30-ΣΔΕΥ'}
PERIOCHI             = 'Α΄ ΘΕΣΣΑΛΟΝΙΚΗΣ (Π.Ε.) 2018'


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path_416 = get_downloaded_file('4.16', 'Αρχείο 4.16 (Αιτιολόγηση Απουσίας) [.csv]:', csv_only=True, silent=True)
    path_421 = get_downloaded_file('4.21', 'Αρχείο 4.21 (Άδειες) [.csv]:', csv_only=True, silent=True)
    if path_416 is None or path_421 is None:
        return {'path_416': path_416, 'path_421': path_421, 'today': None}
    today = ask_date_yyyymmdd()
    return {'path_416': path_416, 'path_421': path_421, 'today': today}


# ── Βοηθητικές ──────────────────────────────────────────────────────────────
def _parse_date(series):
    return pd.to_datetime(series, format='%d/%m/%Y', errors='coerce')

def _clean(series):
    return series.astype(str).str.replace('="', '').str.replace('"', '').str.strip()

def _calc_end(row):
    apo = row['Από_dt']
    if pd.isna(apo):
        return pd.NaT
    try:
        days   = int(float(row['Εγκρ. Ημέρες'])) if pd.notna(row.get('Εγκρ. Ημέρες')) else 0
        months = int(float(row['Εγκρ. Μήνες']))  if pd.notna(row.get('Εγκρ. Μήνες'))  else 0
        years  = int(float(row['Εγκρ. Έτη']))    if pd.notna(row.get('Εγκρ. Έτη'))    else 0
    except (ValueError, TypeError):
        return pd.NaT
    if days == 0 and months == 0 and years == 0:
        return pd.NaT
    return apo + relativedelta(years=years, months=months, days=days - 1)


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    today = ctx['today']
    df16  = read_csv_fixed(ctx['path_416'])
    df21  = read_csv_fixed(ctx['path_421'])

    df16.columns = [c.strip() for c in df16.columns]
    df21.columns = [c.strip() for c in df21.columns]

    ait_col   = next((c for c in df16.columns if 'Αιτιολόγηση Απουσίας' in c), None)
    exeid_col = next((c for c in df16.columns if 'Εξειδ' in c), None)
    sxol_col  = next((c for c in df16.columns if 'Ονομασία Σχολείου' in c), None)
    kod_col   = next((c for c in df16.columns if 'Κωδικός Σχολείου' in c or 'Kωδικός Σχολείου' in c), None)
    eid_col   = next((c for c in df16.columns if 'Κωδικός Κύριας Ειδικότητας' in c), None)
    per_col   = next((c for c in df16.columns if 'Περιοχή Μετάθεσης Εκπαιδευτικού' in c), None)
    afm16_col = next((c for c in df16.columns if c.strip() in ('Α.Φ.Μ.', 'ΑΦΜ')), None)
    am_col    = next((c for c in df16.columns if c.strip() in ('Α.Μ.', 'ΑΜ')), None)
    tel_col   = next((c for c in df16.columns if c.strip() == 'Τηλέφωνο'), None)
    email_col = next((c for c in df16.columns if c.strip() == 'Email'), None)
    epwn_col  = next((c for c in df16.columns if 'Επώνυμο' in c), None)
    onoma_col = next((c for c in df16.columns if c.strip() == 'Όνομα'), None)
    apo_col   = next((c for c in df16.columns if c.strip() == 'Από'), None)
    eos_col   = next((c for c in df16.columns if c.strip() == 'Έως'), None)

    if ait_col is None or afm16_col is None:
        print('  ✗ Δεν βρέθηκαν απαραίτητες στήλες στο 4.16')
        return pd.DataFrame()

    # Μακροχρόνιες
    df16 = df16[df16[ait_col].astype(str).str.contains('ΜΑΚΡΟΧΡΟΝ', case=False, na=False)].copy()
    print(f'  → {len(df16)} μακροχρόνιες απουσίες (4.16)')

    # Εξαίρεση Άνευ Αποδοχών
    if exeid_col:
        df16 = df16[~df16[exeid_col].astype(str).str.contains('νευ', case=False, na=False)].copy()
    print(f'  → {len(df16)} μετά εξαίρεση Άνευ Αποδοχών')

    # Εξαίρεση Ιδιωτικών (κωδικός αρχίζει με 9)
    if sxol_col:
        df16 = df16[~df16[sxol_col].astype(str).str.contains('ΙΔΙΩΤ', case=False, na=False)].copy()
    if kod_col:
        df16['_kod'] = df16[kod_col].astype(str).str.replace('="', '').str.replace('"', '').str.strip()
        df16 = df16[df16['_kod'].str.startswith('9')].copy()
    print(f'  → {len(df16)} μετά εξαίρεση Ιδιωτικών')

    # Εξαίρεση ΣΔΕΥ
    if eid_col:
        df16 = df16[~df16[eid_col].astype(str).str.strip().isin(EXCLUDED_SPECIALTIES)].copy()
    print(f'  → {len(df16)} μετά εξαίρεση ΣΔΕΥ')

    # Φίλτρο περιοχής μετάθεσης εκπαιδευτικού
    if per_col:
        df16 = df16[df16[per_col].astype(str).str.strip() == PERIOCHI].copy()
    print(f'  → {len(df16)} μετά φίλτρο Περιοχής Μετάθεσης')

    if df16.empty:
        return pd.DataFrame()

    df16['ΑΦΜ_clean'] = _clean(df16[afm16_col])
    df16_uniq = df16.drop_duplicates(subset='ΑΦΜ_clean', keep='first').copy()
    print(f'  → {len(df16_uniq)} μοναδικά ΑΦΜ')

    # ── 4.21: ενεργές άδειες ────────────────────────────────────────────────
    afm21_col = next((c for c in df21.columns if c.strip() in ('ΑΦΜ', 'Α.Φ.Μ.')), None)
    if afm21_col is None:
        print('  ✗ Δεν βρέθηκε στήλη ΑΦΜ στο 4.21')
        return pd.DataFrame()

    df21['ΑΦΜ_clean'] = _clean(df21[afm21_col])
    df21['Από_dt']    = _parse_date(df21['Από'])
    df21['Έως_dt']    = df21.apply(_calc_end, axis=1)

    mask_active = (
        (df21['Από_dt'] <= today) &
        (df21['Έως_dt'] >= today) &
        (~df21['Κατάσταση άδειας'].isin(EXCLUDED_STATUS))
    )
    set_adeia = set(df21[mask_active]['ΑΦΜ_clean'])
    print(f'  → {len(set_adeia)} ΑΦΜ με ενεργή άδεια (4.21)')

    result = df16_uniq[~df16_uniq['ΑΦΜ_clean'].isin(set_adeia)].copy()
    print(f'  → {len(result)} απόντες ΧΩΡΙΣ ενεργή άδεια')

    if result.empty:
        return pd.DataFrame()

    def _g(col):
        if col is None or col not in result.columns:
            return pd.Series([''] * len(result), index=result.index)
        return clean_field(result[col])

    out = pd.DataFrame({
        'Κωδικός Σχολείου':           _g(kod_col),
        'Ονομασία Σχολείου':          _g(sxol_col),
        'Τηλέφωνο':                   _g(tel_col),
        'Email':                      _g(email_col),
        'Email Σχολείου':             _g(email_col),
        'Α.Μ.':                       _g(am_col),
        'Α.Φ.Μ.':                     result['ΑΦΜ_clean'].values,
        'Επώνυμο':                    _g(epwn_col),
        'Όνομα':                      _g(onoma_col),
        'Κωδικός Κύριας Ειδικότητας': _g(eid_col),
        'Εξειδ. Αιτιολόγησης':        _g(exeid_col),
        'Από':                        _g(apo_col),
        'Έως':                        _g(eos_col),
    })

    return out.sort_values([SCHOOL_COLUMN, 'Επώνυμο', 'Όνομα']).reset_index(drop=True)


def test_body(df_out, today, schools):
    return (
        f'Σύνοψη ελέγχου απουσίας χωρίς δήλωση άδειας — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Βρέθηκαν: {len(df_out)} εκπαιδευτικοί με μακροχρόνια απουσία χωρίς ενεργή άδεια\n'
        f'Σχολεία που εμφανίζονται ({len(schools)}): {", ".join(sorted(schools))}'
    )
