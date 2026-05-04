"""
checks/adies_missing_421.py
════════════════════════════
Εντοπίζει εκπαιδευτικούς με καταχωρημένη απουσία ΜΑΚΡΟΧΡΟΝΙΑΣ ΑΔΕΙΑΣ στο 4.16
χωρίς αντίστοιχη ενεργή άδεια στο 4.21 — η άδεια πρέπει να καλύπτει τη σημερινή
ημερομηνία, να έχει ίδιο τύπο με την αιτιολόγηση, και να μην είναι Ανακλήθηκε.
Εξαιρεί Άνευ Αποδοχών (ελέγχονται από adies_aneu.py), ΣΔΕΥ και Ιδιωτικά.
"""

from datetime import timedelta
import pandas as pd
import config
from core.framework import get_downloaded_file, read_csv_fixed, ask_date_yyyymmdd

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE       = 'Απουσία Μακροχρόνιας Άδειας χωρίς Εγγραφή στο 4.21'
CHECK_DESCRIPTION = 'Εκπαιδευτικοί με ΜΑΚΡΟΧΡΟΝΙΑ ΑΔΕΙΑ στο 4.16 χωρίς ενεργή άδεια στο 4.21 (εξαιρεί ΑΑ, ΣΔΕΥ, Ιδιωτικά)'
RESULTS_FOLDER    = 'adies_missing_421'
HAS_EMAIL         = True
REQUIRED_REPORTS  = ['4.16 — Αιτιολόγηση απουσίας εκπαιδευτικών', '4.21 — Άδειες εκπαιδευτικών']

COLUMNS = [
    ('Κωδικός Σχολείου',    14),
    ('Ονομασία Σχολείου',   38),
    ('Email Σχολείου',      30),
    ('ΑΜ',                  11),
    ('ΑΦΜ',                 13),
    ('Επώνυμο',             18),
    ('Εξειδ. Αιτιολόγησης', 60),
    ('Από',                 12),
    ('Έως',                 12),
]

SCHOOL_COLUMN = 'Ονομασία Σχολείου'
EMAIL_COLUMN  = 'Email Σχολείου'

EMAIL_SUBJECT = 'Απουσία μακροχρόνιας άδειας χωρίς ενεργή καταχώρηση στο 4.21'
EMAIL_BODY    = lambda school='': (
    'Καλημέρα,\n\n'
    'Εντοπίστηκαν εκπαιδευτικοί στο σχολείο σας που εμφανίζονται ως απόντες '
    'λόγω μακροχρόνιας άδειας (4.16), ενώ δεν υπάρχει αντίστοιχη ενεργή άδεια '
    'στο 4.21 — Άδειες που να καλύπτει τη σημερινή ημερομηνία '
    '(επισυνάπτεται αρχείο).\n\n'
    'Παρακαλούμε να ελεγχθεί η υποβολή/έγκριση της άδειας στο σύστημα.\n\n'
    + config.email_signature()
)

CENTER_COLS = {'ΑΜ', 'ΑΦΜ', 'Από', 'Έως', 'Κωδικός Σχολείου'}


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path_416 = get_downloaded_file('stat4_16', 'Αρχείο 4.16 (Αιτιολόγηση απουσίας) [.csv]:', csv_only=True, silent=True)
    path_421 = get_downloaded_file('4.21',     'Αρχείο 4.21 (Άδειες) [.csv]:', csv_only=True, silent=True)
    if path_416 is None or path_421 is None:
        return {'path_416': path_416, 'path_421': path_421, 'today': None}
    today = ask_date_yyyymmdd()
    return {'path_416': path_416, 'path_421': path_421, 'today': today}


# ── Βοηθητικές ──────────────────────────────────────────────────────────────
def _clean(series):
    return series.astype(str).str.replace('="', '').str.replace('"', '').str.strip()

def _parse_date(series):
    return pd.to_datetime(series, format='%d/%m/%Y', errors='coerce')


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(ctx):
    today    = ctx['today']
    df16     = read_csv_fixed(ctx['path_416'])
    df21     = read_csv_fixed(ctx['path_421'])

    # Φίλτρο 4.16: μακροχρόνια άδεια, εξαιρώ ΑΑ, ΣΔΕΥ, Ιδιωτικά
    mask = (
        (df16['Αιτιολόγηση Απουσίας'] == 'ΜΑΚΡΟΧΡΟΝΙΑ ΑΔΕΙΑ (>10 ημέρες)') &
        (~df16['Εξειδ. Αιτιολόγησης'].str.contains('Ανευ Αποδοχών', na=False)) &
        (~df16['Κωδικός Κύριας Ειδικότητας'].str.contains('ΣΔΕΥ', na=False)) &
        (~df16['Είδος Σχολείου'].str.contains('Ιδιωτ', case=False, na=False))
    )
    makro = df16[mask].copy()
    makro['ΑΦΜ_clean'] = _clean(makro['Α.Φ.Μ.']).str.zfill(9)
    makro['Από_dt']    = _parse_date(makro['Από'])
    makro['Έως_dt']    = _parse_date(makro['Έως'])

    # Φόρτωση 4.21
    df21['ΑΦΜ_clean'] = _clean(df21['ΑΦΜ']).str.zfill(9)
    df21['Από_dt']    = _parse_date(df21['Από'])
    df21['Αιτ_Ημ']   = pd.to_numeric(df21['Αιτ. Ημέρες'], errors='coerce').fillna(0).astype(int)
    df21['Έως_dt']    = df21.apply(
        lambda r: r['Από_dt'] + timedelta(days=int(r['Αιτ_Ημ']) - 1)
        if r['Αιτ_Ημ'] > 0 and pd.notna(r['Από_dt']) else pd.NaT, axis=1
    )

    print(f'  → {len(makro)} γραμμές ΜΑΚΡΟΧΡΟΝΙΑ ΑΔΕΙΑ (εξαιρώ ΑΑ/ΣΔΕΥ/Ιδιωτικά)')

    def has_match(afm, exeid, apo16_dt, eos16_dt):
        # Η απουσία πρέπει να είναι ενεργή σήμερα
        if pd.isna(apo16_dt) or pd.isna(eos16_dt):
            return False
        if not (apo16_dt <= today <= eos16_dt):
            return False
        rows = df21[
            (df21['ΑΦΜ_clean'] == afm) &
            (df21['Κατάσταση άδειας'] != '5-Ανακλήθηκε')
        ]
        for _, r in rows.iterrows():
            type_21 = str(r['Τύπος άδειας']).strip()
            # Ίδιος τύπος: το 4.21 είναι substring/prefix του 4.16
            if not (type_21 in exeid or exeid.startswith(type_21)):
                continue
            # Η άδεια καλύπτει σήμερα
            if pd.notna(r['Από_dt']) and pd.notna(r['Έως_dt']):
                if r['Από_dt'] <= today <= r['Έως_dt']:
                    return True
        return False

    records = []
    seen_afm = set()
    for _, row in makro.iterrows():
        afm   = row['ΑΦΜ_clean']
        if afm in seen_afm:
            continue
        exeid = str(row['Εξειδ. Αιτιολόγησης']).strip()
        if not has_match(afm, exeid, row['Από_dt'], row['Έως_dt']):
            seen_afm.add(afm)
            records.append({
                'Κωδικός Σχολείου':    _clean(pd.Series([row['Kωδικός Σχολείου']])).iloc[0],
                'Ονομασία Σχολείου':   str(row['Ονομασία Σχολείου']).strip(),
                'Email Σχολείου':      str(row['Email']).strip(),
                'ΑΜ':                  _clean(pd.Series([row['Α.Μ.']])).iloc[0],
                'ΑΦΜ':                 afm,
                'Επώνυμο':             str(row['Επώνυμο']).strip(),
                'Εξειδ. Αιτιολόγησης': exeid,
                'Από':                 str(row['Από']).strip(),
                'Έως':                 str(row['Έως']).strip(),
            })

    print(f'  → {len(records)} με πρόβλημα')
    if not records:
        return pd.DataFrame()

    df_out = pd.DataFrame(records)
    return df_out.sort_values(['Ονομασία Σχολείου', 'Επώνυμο'])


def test_body(df_out, today, schools):
    return (
        f'Σύνοψη ελέγχου ΜΑΚΡΟΧΡΟΝΙΑ ΑΔΕΙΑ (4.16) vs 4.21 — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Σύνολο με πρόβλημα: {len(df_out)}\n'
        f'Σχολεία ({len(schools)}): {", ".join(sorted(schools))}'
    )
