"""
checks/tmimata_genikis.py
═════════════════════════
Έλεγχος τμημάτων γενικής παιδείας / δυναμικού — πλήρης αποτύπωση (2026-2027).

Πηγές (κατεβαίνουν ΑΠΕΥΘΕΙΑΣ μέσα στον έλεγχο — δες _download_inputs):
  3.1 (stat3_1) — Κατανομή μαθητών ανά τάξη. 2 γραμμές header (η 2η μόνο για τις
                   στήλες Αγόρια/Κορίτσια/Σύνολο, λόγω merged cell στο excel).
                   Στήλες (θέση): 2=Είδος Σχολείου, 4=Κωδικός Υπ., 10=Τάξη,
                   11=Αριθμός Τμημάτων, 14=Σύνολο μαθητών.
  5.3 (stat5_3) — Αποτύπωση νηπιαγωγείων: Τύπος Σχολείου, Κωδ. Υπουργείου Σχολείου,
                   Ονομασία Σχολ. Μονάδας, Λειτουργικότητα, Οργανικότητα, Μαθητές.
  5.4 (stat5_4) — Αποτύπωση δημοτικών: ίδιες στήλες με το 5.3.

Γιατί η λήψη γίνεται μέσα στον έλεγχο (και όχι μέσω του γενικού «⬇ Λήψη Δεδομένων»):
  Και τα 3 στατιστικά πρέπει να αντληθούν με το σχολικό έτος 2026-2027 ενεργό στο
  MySchool (το 3.1 κανονικά αντλείται με το τρέχον έτος — το χρησιμοποιούν και άλλα
  εργαλεία, π.χ. το «Σχολικές Μονάδες» — γι' αυτό εδώ γίνεται ένα τοπικό override
  μόνο για αυτή τη λήψη, χωρίς να πειράζεται η καθολική ρύθμιση στο downloader.py).
  Η αλλαγή έτους γίνεται αυτόματα στην αρχή (Default.aspx → cmbActiveAcadYear).

Λογική:
  Για κάθε σχολείο (πηγή: 5.3/5.4) υπολογίζεται, ανά τάξη, το άθροισμα Τμημάτων
  και Μαθητών από το 3.1 (Νηπιαγωγεία: μία ενιαία στήλη «ΠΡΟΝΗΠΙΑ-ΝΗΠΙΑ» — αθροίζει
  όλες τις σχετικές τάξεις π.χ. ΝΗΠΙΑ/ΠΡΟΝΗΠΙΑ/ΠΡΟΝΗΠΙΑ-ΝΗΠΙΑ. Δημοτικά: Α-ΣΤ).

  Άθροισμα Τμήματα  = Σ(Τμήματα ανά τάξη)
  Άθροισμα Μαθητών  = Σ(Μαθητές ανά τάξη)
  Διαφορά Τμήματα   = Λειτουργικότητα − Άθροισμα Τμήματα
  Διαφορά Μαθητές   = Ενεργοί Μαθητές (5.3/5.4 «Μαθητές») − Άθροισμα Μαθητών

Αποτέλεσμα: 1 αρχείο Excel με 2 φύλλα (Δημοτικά, Νηπιαγωγεία) — ΟΛΑ τα σχολεία,
με πράσινο/κόκκινο χρωματισμό στις στήλες Διαφορά (0 = πράσινο, ≠0 = κόκκινο).
Καμία αποστολή email.
"""

import glob as _glob
import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.framework import _show_results_popup

# ── Μεταδεδομένα ────────────────────────────────────────────────────────────
CHECK_TITLE       = 'Έλεγχος Τμημάτων Γενικής Παιδείας / Δυναμικού (2026-2027)'
CHECK_DESCRIPTION = 'Πλήρης αποτύπωση Λειτουργικότητας vs Τμημάτων/Μαθητών ανά τάξη για νηπιαγωγεία (5.3) και δημοτικά (5.4), με στοιχεία κατανομής από το 3.1 — η λήψη γίνεται αυτόματα μέσα στον έλεγχο (αλλαγή σχολικού έτους σε 2026-2027)'
RESULTS_FOLDER    = 'tmimata_genikis'
HAS_EMAIL         = False
CUSTOM_RUN        = True
SCHOOL_YEAR       = '2026-2027'
REQUIRED_REPORTS  = [
    '3.1 — Κατανομή μαθητών ανά τάξη (κατεβαίνει αυτόματα)',
    '5.3 — Αποτύπωση νηπιαγωγείων (κατεβαίνει αυτόματα)',
    '5.4 — Αποτύπωση δημοτικών (κατεβαίνει αυτόματα)',
]

DEFAULT_EMAIL_SUBJECT = (
    'Τμήματα Γενικής Παιδείας {school_year} — αποκλίσεις MySchool'
)
DEFAULT_EMAIL_BODY = (
    'Καλημέρα,\n\n'
    'Κατά τον έλεγχο των στοιχείων τμημάτων γενικής παιδείας / δυναμικού '
    'για το σχολικό έτος {school_year} στο MySchool, εντοπίστηκαν αποκλίσεις '
    'μεταξύ Λειτουργικότητας και καταχωρημένων τμημάτων/μαθητών '
    'για το σχολείο σας.\n\n'
    'Παρακαλούμε ελέγξτε το συνημμένο αρχείο και προβείτε στις απαραίτητες '
    'διορθώσεις στο MySchool.\n\n'
    'Παρακαλούμε για τις ενέργειές σας.'
)

GRADES_DS  = ['Α', 'Β', 'Γ', 'Δ', 'Ε', 'ΣΤ']          # Δημοτικά
NIP_GROUP  = 'ΠΡΟΝΗΠΙΑ-ΝΗΠΙΑ'                          # Νηπιαγωγεία — ενιαία στήλη

COLOR_HDR   = '1F4E79'
COLOR_SUB   = 'D6E4F0'
COLOR_ALT   = 'EBF3FB'
COLOR_OK    = '92D050'   # πράσινο — διαφορά 0
COLOR_DEV   = 'FF0000'   # κόκκινο — διαφορά ≠ 0
DEFAULT_PERIFEREIA = 'ΚΕΝΤΡΙΚΗΣ ΜΑΚΕΔΟΝΙΑΣ'

_LEFT_ALIGN_COLS = {'Τύπος', 'Ονομασία', 'Δήμος', 'Email Σχολείου'}


# ── Φόρτωση 3.1 ─────────────────────────────────────────────────────────────
def _find_col(header, *names):
    """Επιστρέφει το index της πρώτης στήλης του header που ταιριάζει με κάποιο name."""
    for i, v in enumerate(header):
        if str(v).strip() in names:
            return i
    return None


def _load_stat31(path):
    """
    Φορτώνει το 3.1 και επιστρέφει DataFrame με στήλες
    _eidos, _code (int), _taxi, _tmimata (int), _mathites (int).

    Το αρχείο έχει 2 γραμμές header (η 2η μόνο για Αγόρια/Κορίτσια/Σύνολο,
    λόγω merged cell) — εντοπίζεται δυναμικά η γραμμή με «Τάξη». Οι στήλες
    εντοπίζονται με βάση το ΟΝΟΜΑ τους (όχι σταθερή θέση), γιατί ο αριθμός/η
    σειρά των στηλών αλλάζει ανάλογα με τις επιλογές ομαδοποίησης που
    τσεκαρίστηκαν κατά τη λήψη (π.χ. αν λείπει η «Τύπος Σχολείου»).
    """
    raw = pd.read_excel(path, header=None)

    hdr_row = 0
    for i in range(min(5, len(raw))):
        if raw.iloc[i].astype(str).str.strip().eq('Τάξη').any():
            hdr_row = i
            break

    header = raw.iloc[hdr_row].astype(str).str.strip().tolist()

    eidos_idx   = _find_col(header, 'Είδος Σχολείου')
    code_idx    = _find_col(header, 'Κωδικός Υπ.', 'Κωδικός Υπουργείου Σχολείου', 'Κωδικός Υπουργείου')
    taxi_idx    = _find_col(header, 'Τάξη')
    tmim_idx    = _find_col(header, 'Αριθμός Τμημάτων')
    plithos_idx = _find_col(header, 'Πλήθος Ενεργών Μαθητών')

    missing = [n for n, v in (
        ('Είδος Σχολείου', eidos_idx), ('Κωδικός Υπ.', code_idx), ('Τάξη', taxi_idx),
        ('Αριθμός Τμημάτων', tmim_idx), ('Πλήθος Ενεργών Μαθητών', plithos_idx),
    ) if v is None]
    if missing:
        raise ValueError(
            f'Το 3.1 δεν έχει τις αναμενόμενες στήλες ({", ".join(missing)}).\n'
            f'Στήλες αρχείου: {header}'
        )

    # «Πλήθος Ενεργών Μαθητών» είναι merged header πάνω από 3 υποστήλες
    # (Αγόρια, Κορίτσια, Σύνολο) — το Σύνολο είναι η 3η από αυτές.
    sum_idx = plithos_idx + 2

    data = raw.iloc[hdr_row + 2:].reset_index(drop=True)

    dimos_idx = _find_col(header, 'Δήμος')
    if dimos_idx is None and len(header) > 7:
        dimos_idx = 7  # fallback: absolute position

    df = pd.DataFrame({
        '_eidos':    data[eidos_idx].astype(str).str.strip(),
        '_code':     pd.to_numeric(data[code_idx], errors='coerce'),
        '_taxi':     data[taxi_idx].astype(str).str.strip(),
        '_tmimata':  pd.to_numeric(data[tmim_idx], errors='coerce').fillna(0).astype(int),
        '_mathites': pd.to_numeric(data[sum_idx], errors='coerce').fillna(0).astype(int),
        '_dimos':    data[dimos_idx].astype(str).str.strip() if dimos_idx is not None else '',
    })
    df = df.dropna(subset=['_code'])
    df['_code'] = df['_code'].astype(int)
    return df


def _find_stat22():
    """Αυτόματη εύρεση stat2_2 από τον φάκελο λήψεων της εφαρμογής ή ~/Downloads."""
    docs = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks')
    folders = []
    dl_base = os.path.join(docs, 'downloads')
    if os.path.isdir(dl_base):
        folders += sorted(
            [os.path.join(dl_base, d) for d in os.listdir(dl_base)
             if os.path.isdir(os.path.join(dl_base, d))],
            reverse=True
        )
    folders.append(os.path.join(os.path.expanduser('~'), 'Downloads'))
    for folder in folders:
        for pattern in ('stat2_2*.csv', 'stat2_2*.xlsx', 'stat2_2*.xls', 'CSV_*.zip'):
            matches = [f for f in _glob.glob(os.path.join(folder, pattern))
                       if not f.endswith('.tmp') and not f.endswith('.crdownload')]
            if matches:
                return sorted(matches)[-1]
    return None


def _load_email_lookup(path_22):
    """
    Διαβάζει stat2_2 (CSV με 1-column shift) και επιστρέφει {code_int: email}.
    col11 = Κωδ. ΥΠΠΘ, col18 = e-mail σχολείου.
    """
    import re
    lower = path_22.lower()
    if lower.endswith('.xlsx') or lower.endswith('.xls'):
        df = pd.read_excel(path_22, dtype=str)
    else:
        import zipfile, io
        if lower.endswith('.zip'):
            with zipfile.ZipFile(path_22) as z:
                raw = z.read(z.namelist()[0])
        else:
            with open(path_22, 'rb') as f:
                raw = f.read()
        text = raw.decode('cp1253')
        df = pd.read_csv(io.StringIO(text), sep=';', dtype=str, header=0)

    if len(df.columns) <= 18:
        return {}

    c_code  = df.columns[11]
    c_email = df.columns[18]

    def _clean_code(val):
        s = str(val).strip().strip('"').lstrip('=').strip('"').strip()
        s = re.sub(r'\.0$', '', s)
        return s.lstrip('0') or s

    lookup = {}
    for _, row in df.iterrows():
        raw_code = _clean_code(row[c_code])
        email    = str(row[c_email]).strip() if pd.notna(row[c_email]) else ''
        if email.lower() in ('nan', 'none', ''):
            email = ''
        if raw_code.isdigit():
            lookup[int(raw_code)] = email
    return lookup


def _dimos_lookup(df31):
    """dict {code: dimos} — ένας Δήμος ανά κωδικό σχολείου."""
    sub = df31.drop_duplicates(subset=['_code'])[['_code', '_dimos']]
    return {int(r['_code']): str(r['_dimos']).strip() for _, r in sub.iterrows()}


def _grade_lookup_ds(df31):
    """dict {code: {grade: (τμήματα, μαθητές)}} — Δημοτικά (Α-ΣΤ)."""
    sub = df31[df31['_eidos'] == 'Δημοτικά Σχολεία']
    g = sub.groupby(['_code', '_taxi'])[['_tmimata', '_mathites']].sum().reset_index()
    lookup = {}
    for _, r in g.iterrows():
        lookup.setdefault(int(r['_code']), {})[r['_taxi']] = (int(r['_tmimata']), int(r['_mathites']))
    return lookup


def _grade_lookup_nip(df31):
    """dict {code: (τμήματα, μαθητές)} — Νηπιαγωγεία, όλες οι τάξεις αθροισμένες."""
    sub = df31[df31['_eidos'] == 'Νηπιαγωγεία']
    g = sub.groupby('_code')[['_tmimata', '_mathites']].sum().reset_index()
    return {int(r['_code']): (int(r['_tmimata']), int(r['_mathites'])) for _, r in g.iterrows()}


# ── Στήλες εξόδου ─────────────────────────────────────────────────────────
def _base_cols():
    return [
        ('Τύπος',             42),
        ('Κωδ. ΥΠΠΘ',         14),
        ('Ονομασία',          40),
        ('Δήμος',             20),
        ('Email Σχολείου',    32),
        ('Λειτουργικότητα',   14),
        ('Οργανικότητα',      14),
        ('Ενεργοί Μαθητές',   14),
    ]

def _tail_cols():
    return [
        ('Άθροισμα Μαθητών',  16),
        ('Άθροισμα Τμήματα',  16),
        ('Διαφορά Τμήματα',   16),
        ('Διαφορά Μαθητές',   16),
    ]

def columns_ds():
    grade_cols = []
    for g in GRADES_DS:
        grade_cols += [(f'{g} Μαθητές', 10), (f'{g} Τμήματα', 10)]
    return _base_cols() + grade_cols + _tail_cols()

def columns_nip():
    grade_cols = [(f'{NIP_GROUP} Μαθητές', 18), (f'{NIP_GROUP} Τμήματα', 18)]
    return _base_cols() + grade_cols + _tail_cols()


# ── Χτίσιμο DataFrame ανά τύπο σχολείου ─────────────────────────────────────
def _build_records(df_src, grade_lookup, multi_grade, dimos_lookup=None, email_lookup=None):
    """
    df_src        : DataFrame από 5.3 ή 5.4
    grade_lookup  : dict {code: {grade: (tm, ma)}}  (Δημοτικά) ή
                     dict {code: (tm, ma)}           (Νηπιαγωγεία)
    multi_grade   : True → Δημοτικά (πολλαπλές τάξεις Α-ΣΤ)
                     False → Νηπιαγωγεία (μία ενιαία στήλη)
    dimos_lookup  : dict {code: dimos} από το 3.1 (προαιρετικό)
    email_lookup  : dict {code: email} από το 2.2 (προαιρετικό)
    """
    if dimos_lookup is None:
        dimos_lookup = {}
    if email_lookup is None:
        email_lookup = {}
    records = []
    for _, row in df_src.iterrows():
        code  = int(row['Κωδ. Υπουργείου Σχολείου'])
        leit  = int(row['Λειτουργικότητα']) if pd.notna(row['Λειτουργικότητα']) else 0
        org   = int(row['Οργανικότητα'])    if pd.notna(row['Οργανικότητα'])    else 0
        energ = int(row['Μαθητές'])         if pd.notna(row['Μαθητές'])         else 0

        rec = {
            'Τύπος':             str(row['Τύπος Σχολείου']).strip(),
            'Κωδ. ΥΠΠΘ':         code,
            'Ονομασία':          str(row['Ονομασία Σχολ. Μονάδας']).strip(),
            'Δήμος':             dimos_lookup.get(code, ''),
            'Email Σχολείου':    email_lookup.get(code, ''),
            'Λειτουργικότητα':   leit,
            'Οργανικότητα':      org,
            'Ενεργοί Μαθητές':   energ,
        }

        sum_tm = sum_ma = 0
        if multi_grade:
            per_school = grade_lookup.get(code, {})
            for g in GRADES_DS:
                tm, ma = per_school.get(g, (0, 0))
                rec[f'{g} Μαθητές'] = ma
                rec[f'{g} Τμήματα'] = tm
                sum_tm += tm
                sum_ma += ma
        else:
            tm, ma = grade_lookup.get(code, (0, 0))
            rec[f'{NIP_GROUP} Μαθητές'] = ma
            rec[f'{NIP_GROUP} Τμήματα'] = tm
            sum_tm += tm
            sum_ma += ma

        rec['Άθροισμα Μαθητών'] = sum_ma
        rec['Άθροισμα Τμήματα'] = sum_tm
        rec['Διαφορά Τμήματα']  = leit  - sum_tm
        rec['Διαφορά Μαθητές']  = energ - sum_ma
        records.append(rec)

    df_out = pd.DataFrame(records)
    if not df_out.empty:
        df_out = df_out.sort_values('Ονομασία').reset_index(drop=True)
    return df_out


def _periferia_name(df_src):
    try:
        raw = str(df_src['Περιφέρεια'].dropna().iloc[0])
        if 'ΕΚΠ/ΣΗΣ ' in raw:
            return raw.split('ΕΚΠ/ΣΗΣ ')[-1].strip()
    except Exception:
        pass
    return DEFAULT_PERIFEREIA


# ── Λογική ──────────────────────────────────────────────────────────────────
def process(path_31, path_53, path_54, path_22=None):
    """Επιστρέφει (df_ds, df_nip, periфereia_name)."""
    df31 = _load_stat31(path_31)
    df53 = pd.read_excel(path_53)
    df54 = pd.read_excel(path_54)

    dimos = _dimos_lookup(df31)
    email = _load_email_lookup(path_22) if path_22 else {}
    df_ds  = _build_records(df54, _grade_lookup_ds(df31),  multi_grade=True,  dimos_lookup=dimos, email_lookup=email)
    df_nip = _build_records(df53, _grade_lookup_nip(df31), multi_grade=False, dimos_lookup=dimos, email_lookup=email)

    perif = _periferia_name(df54) if not df54.empty else _periferia_name(df53)
    return df_ds, df_nip, perif


# ── Excel ─────────────────────────────────────────────────────────────────
def _brd():
    t = Side(style='thin', color='CCCCCC')
    return Border(left=t, right=t, top=t, bottom=t)


def _write_sheet(ws, df, title, columns, today, subtitle_extra=''):
    brd = _brd()
    ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
    lft = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    ncols = len(columns)

    ws.merge_cells(f'A1:{get_column_letter(ncols)}1')
    ws['A1'] = f'{title}  —  {today.strftime("%d/%m/%Y")}{subtitle_extra}'
    ws['A1'].font      = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    ws['A1'].fill      = PatternFill('solid', start_color=COLOR_HDR)
    ws['A1'].alignment = ctr
    ws.row_dimensions[1].height = 24

    ws.merge_cells(f'A2:{get_column_letter(ncols)}2')
    ws['A2'] = f'Σύνολο σχολείων: {len(df)}'
    ws['A2'].font      = Font(name='Arial', italic=True, size=9)
    ws['A2'].fill      = PatternFill('solid', start_color=COLOR_SUB)
    ws['A2'].alignment = ctr
    ws.row_dimensions[2].height = 16

    for ci, (name, width) in enumerate(columns, 1):
        c = ws.cell(row=3, column=ci, value=name)
        c.font      = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill      = PatternFill('solid', start_color=COLOR_HDR)
        c.alignment = ctr
        c.border    = brd
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[3].height = 28

    diff_cols = {'Διαφορά Τμήματα', 'Διαφορά Μαθητές'}
    col_keys  = [c[0] for c in columns]

    for ri, (_, row) in enumerate(df.iterrows(), start=4):
        base_fill = PatternFill('solid', start_color=COLOR_ALT) if ri % 2 == 0 else PatternFill()
        for ci, key in enumerate(col_keys, 1):
            val = row.get(key, '')
            if hasattr(val, 'item'):
                val = val.item()
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = brd
            if key in diff_cols:
                is_dev = val != 0
                c.font = Font(name='Arial', size=9, bold=True,
                               color='FFFFFF' if is_dev else '000000')
                c.fill = PatternFill('solid', start_color=COLOR_DEV if is_dev else COLOR_OK)
            else:
                c.font = Font(name='Arial', size=9)
                c.fill = base_fill
            c.alignment = lft if key in _LEFT_ALIGN_COLS else ctr
        ws.row_dimensions[ri].height = 16

    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f'A3:{get_column_letter(ncols)}3'


def build_workbook(df_ds, df_nip, perif, today, out_path):
    wb = Workbook()

    ws1 = wb.active
    ws1.title = f'ΔΣ-{perif}'[:31]
    _write_sheet(ws1, df_ds, CHECK_TITLE, columns_ds(), today, subtitle_extra='  |  Δημοτικά')

    ws2 = wb.create_sheet(f'ΝΗΠ-{perif}'[:31])
    _write_sheet(ws2, df_nip, CHECK_TITLE, columns_nip(), today, subtitle_extra='  |  Νηπιαγωγεία')

    wb.save(out_path)


# ── Email — Βοηθητικά ─────────────────────────────────────────────────────────
def _app_base():
    import sys
    if getattr(sys, 'frozen', False):
        lad = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'MySchoolChecks')
        return lad if os.path.isdir(lad) else os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_tmimata_email_template():
    import json
    spath = os.path.join(_app_base(), 'data', 'local_settings.json')
    try:
        with open(spath, encoding='utf-8') as f:
            s = json.load(f)
        t = s.get('tmimata_email', {})
        return t.get('subject', DEFAULT_EMAIL_SUBJECT), t.get('body', DEFAULT_EMAIL_BODY)
    except Exception:
        return DEFAULT_EMAIL_SUBJECT, DEFAULT_EMAIL_BODY


def _save_tmimata_email_template(subject, body):
    import json
    spath = os.path.join(_app_base(), 'data', 'local_settings.json')
    os.makedirs(os.path.dirname(spath), exist_ok=True)
    try:
        with open(spath, encoding='utf-8') as f:
            s = json.load(f)
    except Exception:
        s = {}
    s['tmimata_email'] = {'subject': subject, 'body': body}
    with open(spath, 'w', encoding='utf-8') as f:
        json.dump(s, f, ensure_ascii=False, indent=2)


def _build_mini_excel(row_df, col_defs, label, today):
    """Mini Excel (bytes) για ένα σχολείο."""
    import io as _io
    buf = _io.BytesIO()
    wb  = Workbook()
    ws  = wb.active
    ws.title = label[:31]
    _write_sheet(ws, row_df, CHECK_TITLE, col_defs, today,
                 subtitle_extra=f'  |  {label}')
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _qualifying_schools(df_ds, df_nip, threshold):
    """
    Επιστρέφει list of (code, name, email, row_df, col_defs, label)
    για σχολεία με |Διαφορά Μαθητές| > threshold ή Διαφορά Τμήματα ≠ 0.
    """
    result = []
    for df, col_defs, label in (
        (df_ds,  columns_ds(),  'Δημοτικά'),
        (df_nip, columns_nip(), 'Νηπιαγωγεία'),
    ):
        if df.empty:
            continue
        mask = (df['Διαφορά Τμήματα'] != 0) | (df['Διαφορά Μαθητές'].abs() > threshold)
        for _, row in df[mask].iterrows():
            result.append((
                row['Κωδ. ΥΠΠΘ'],
                row['Ονομασία'],
                row.get('Email Σχολείου', ''),
                pd.DataFrame([row]),
                col_defs,
                label,
            ))
    return result


def _send_summary_tmimata(config, sent, failed, today, log):
    """Summary email στο FROM_EMAIL μετά την ολοκλήρωση."""
    import ssl, smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import formataddr, formatdate
    from email.header import Header

    ef = getattr(config, 'FROM_EMAIL',    '')
    pw = getattr(config, 'FROM_PASSWORD', '')
    sh = getattr(config, 'SMTP_HOST',     'mail.sch.gr')
    if not ef or not pw:
        return

    sent_lines   = '\n'.join(f'  ✓  {n}  ({e})' for n, e in sorted(sent))
    failed_lines = ('\n'.join(f'  ✗  {n}  ({e})' for n, e in sorted(failed))
                    if failed else '  —')
    body = (
        f'Αποστολή email Τμημάτων Γενικής Παιδείας — {today.strftime("%d/%m/%Y")}\n\n'
        f'Εστάλησαν: {len(sent)}\n{sent_lines}\n\n'
        + (f'Αποτυχίες ({len(failed)}):\n{failed_lines}' if failed else '')
    )
    subject = (
        f'[Τμήματα {SCHOOL_YEAR}] Αποστολή ολοκληρώθηκε — {len(sent)} ✓'
        + (f', {len(failed)} ✗' if failed else '')
    )
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From']    = formataddr((str(Header('MySchool Checks', 'utf-8')), ef))
    msg['To']      = ef
    msg['Date']    = formatdate(localtime=True)
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    log(f'\n  ✉ Summary → {ef}')
    try:
        ctx = ssl.create_default_context()
        srv = smtplib.SMTP(sh, 587)
        srv.starttls(context=ctx)
        r = srv.login(ef, pw)
        if r[0] == 235:
            srv.sendmail(ef, ef, msg.as_string())
            log('  ✓ Summary εστάλη.')
        else:
            log('  ✗ Summary: login απέτυχε.')
        srv.quit()
    except Exception as e:
        log(f'  ✗ Summary: {e}')


def _do_send_emails(df_ds, df_nip, threshold, dry_run, subj_tpl, body_tpl,
                    config, today, log):
    """Αποστολή email ανά σχολείο. Επιστρέφει (sent, failed)."""
    import ssl, smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.utils import formataddr, formatdate
    from email.header import Header
    from email import encoders

    ef = getattr(config, 'FROM_EMAIL',    '')
    pw = getattr(config, 'FROM_PASSWORD', '')
    sh = getattr(config, 'SMTP_HOST',     'mail.sch.gr')
    if not ef:
        log('  ✗ FROM_EMAIL δεν έχει οριστεί στις Ρυθμίσεις.')
        return [], []
    if not pw and not dry_run:
        log('  ✗ Κωδικός email δεν έχει οριστεί στις Ρυθμίσεις.')
        return [], []

    schools = _qualifying_schools(df_ds, df_nip, threshold)
    if not schools:
        log('  ℹ Κανένα σχολείο δεν πληροί τα κριτήρια αποστολής.')
        return [], []

    log(f'  → {len(schools)} σχολεία για αποστολή (όριο μαθητών: {threshold})')
    sent, failed = [], []

    for code, name, email_to, row_df, col_defs, label in schools:
        if not email_to:
            log(f'  ⚠ {name}: κενό email — παράλειψη')
            failed.append((name, '—'))
            continue

        subj     = subj_tpl.replace('{school_year}', SCHOOL_YEAR).replace('{school_name}', name)
        body_txt = body_tpl.replace('{school_year}', SCHOOL_YEAR).replace('{school_name}', name)
        xls_bytes = _build_mini_excel(row_df, col_defs, label, today)
        fname_att = f'{today.strftime("%Y%m%d")}_{code}_{label}.xlsx'
        dest      = ef if dry_run else email_to

        msg = MIMEMultipart()
        msg['Subject'] = subj
        msg['From']    = formataddr((str(Header('ΔΙ.Π.Ε. Αν. Θεσ/νίκης', 'utf-8')), ef))
        msg['To']      = formataddr((str(Header(name, 'utf-8')), dest))
        msg['Date']    = formatdate(localtime=True)
        msg.attach(MIMEText(body_txt, 'plain', 'utf-8'))
        part = MIMEBase('application',
                        'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(xls_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment',
                        filename=Header(fname_att, 'utf-8').encode())
        msg.attach(part)

        if dry_run:
            log(f'  (test) → {ef}  [{name}]')
            sent.append((name, email_to))
            continue

        try:
            ctx = ssl.create_default_context()
            srv = smtplib.SMTP(sh, 587)
            srv.starttls(context=ctx)
            r = srv.login(ef, pw)
            if r[0] == 235:
                srv.sendmail(ef, dest, msg.as_string())
                log(f'  ✓ → {email_to}  [{name}]')
                sent.append((name, email_to))
            else:
                log(f'  ✗ Login απέτυχε → {name}')
                failed.append((name, email_to))
            srv.quit()
        except Exception as e:
            log(f'  ✗ {name}: {e}')
            failed.append((name, email_to))

    if not dry_run and sent:
        _send_summary_tmimata(config, sent, failed, today, log)

    return sent, failed


def _open_tmimata_template_editor(parent, C):
    """Dialog επεξεργασίας θέματος & κειμένου email."""
    import tkinter as tk
    cur_subj, cur_body = _load_tmimata_email_template()
    if cur_body == DEFAULT_EMAIL_BODY:
        try:
            import config as _cfg
            sig = getattr(_cfg, 'EMAIL_SIGNATURE', '').strip()
            if sig:
                cur_body = cur_body + '\n\n' + sig
        except Exception:
            pass

    ed = tk.Toplevel(parent)
    ed.title('Πρότυπο Email — Τμήματα Γενικής Παιδείας')
    ed.configure(bg=C['bg'])
    ed.resizable(True, False)
    ed.grab_set()
    ed.transient(parent)

    epad = dict(padx=14, pady=5)
    tk.Label(ed, text='Θέμα:', bg=C['bg'], fg=C['hdr_bg'],
             font=('Arial', 9, 'bold'), anchor='w').pack(fill='x', **epad)
    subj_var = tk.StringVar(value=cur_subj)
    tk.Entry(ed, textvariable=subj_var, font=('Arial', 9),
             width=70).pack(fill='x', padx=14, pady=(0, 8))

    tk.Label(ed, text='Κείμενο email:', bg=C['bg'], fg=C['hdr_bg'],
             font=('Arial', 9, 'bold'), anchor='w').pack(fill='x', **epad)
    txt = tk.Text(ed, font=('Arial', 9), width=70, height=14,
                  wrap='word', relief='solid', bd=1)
    txt.pack(fill='x', padx=14, pady=(0, 4))
    txt.insert('1.0', cur_body)

    tk.Label(ed,
             text='Χρησιμοποιήστε {school_year} και {school_name} (αντικαθίστανται αυτόματα).',
             bg=C['bg'], fg=C['desc'], font=('Arial', 8),
             anchor='w').pack(fill='x', padx=14, pady=(0, 10))

    def _save():
        _save_tmimata_email_template(subj_var.get().strip(), txt.get('1.0', 'end-1c'))
        ed.destroy()
        import tkinter.messagebox as _mb
        _mb.showinfo('Αποθήκευση', 'Το πρότυπο αποθηκεύτηκε.', parent=parent)

    def _reset():
        import tkinter.messagebox as _mb
        if _mb.askyesno('Επαναφορά', 'Να επανέλθει το προεπιλεγμένο κείμενο;', parent=ed):
            _save_tmimata_email_template(DEFAULT_EMAIL_SUBJECT, DEFAULT_EMAIL_BODY)
            ed.destroy()

    br = tk.Frame(ed, bg=C['bg'])
    br.pack(pady=(0, 12))
    tk.Button(br, text='Αποθήκευση', bg=C['btn_bg'], fg=C['btn_fg'],
              font=('Arial', 9, 'bold'), relief='flat', padx=14, pady=5,
              cursor='hand2', command=_save).pack(side='left', padx=4)
    tk.Button(br, text='Επαναφορά προεπιλογής', bg=C['bg2'], fg=C['hdr_bg'],
              font=('Arial', 9), relief='flat', padx=14, pady=5,
              cursor='hand2', command=_reset).pack(side='left', padx=4)
    tk.Button(br, text='Άκυρο', bg=C['bg2'], fg=C['desc'],
              font=('Arial', 9), relief='flat', padx=14, pady=5,
              cursor='hand2', command=ed.destroy).pack(side='left', padx=4)

    ed.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width()  - ed.winfo_width())  // 2
    y = parent.winfo_y() + (parent.winfo_height() - ed.winfo_height()) // 2
    ed.geometry(f'+{x}+{y}')


def _show_email_dialog(config, df_ds, df_nip, today):
    """Dialog αποστολής email ανά σχολείο με απόκλιση."""
    import tkinter as tk
    from tkinter import scrolledtext
    import threading

    root = tk._default_root
    if root is None:
        return

    C = {
        'bg': '#F5F7FA', 'bg2': '#E8EDF3', 'hdr_bg': '#1F4E79',
        'btn_bg': '#1F4E79', 'btn_fg': '#FFFFFF',
        'desc': '#666666',   'sel_bg': '#D6E4F0',
    }
    # Αν δεν υπάρχουν καθόλου αποκλίσεις, δεν εμφανίζουμε το dialog
    all_dev = (
        (not df_ds.empty  and ((df_ds['Διαφορά Τμήματα']  != 0) | (df_ds['Διαφορά Μαθητές']  != 0)).any()) or
        (not df_nip.empty and ((df_nip['Διαφορά Τμήματα'] != 0) | (df_nip['Διαφορά Μαθητές'] != 0)).any())
    )
    if not all_dev:
        return

    dlg = tk.Toplevel(root)
    dlg.title('Αποστολή Email — Τμήματα Γενικής Παιδείας')
    dlg.configure(bg=C['bg'])
    dlg.resizable(False, False)

    pad = dict(padx=14, pady=5)

    # Header
    hdr = tk.Frame(dlg, bg=COLOR_HDR)
    hdr.pack(fill='x')
    tk.Label(hdr, text='Αποστολή Email ανά Σχολείο',
             bg=COLOR_HDR, fg='white', font=('Arial', 11, 'bold'),
             padx=14, pady=8).pack(side='left')

    body_f = tk.Frame(dlg, bg=C['bg'])
    body_f.pack(fill='both', expand=True)

    # Κουμπί πρότυπου
    tk.Button(body_f, text='✉  Πρότυπο Email (Θέμα & Κείμενο)',
              bg=C['bg2'], fg=C['hdr_bg'], font=('Arial', 9),
              relief='flat', padx=10, pady=4, cursor='hand2',
              command=lambda: _open_tmimata_template_editor(dlg, C)
              ).pack(anchor='w', **pad)

    # Όριο απόκλισης μαθητών
    thr_f = tk.Frame(body_f, bg=C['bg'])
    thr_f.pack(fill='x', **pad)
    tk.Label(thr_f, text='Ελάχιστη απόκλιση μαθητών για αποστολή:',
             bg=C['bg'], fg=C['hdr_bg'], font=('Arial', 9)).pack(side='left')
    thr_var = tk.IntVar(value=0)
    tk.Spinbox(thr_f, from_=0, to=99, textvariable=thr_var,
               width=5, font=('Arial', 9)).pack(side='left', padx=6)
    count_lbl = tk.Label(thr_f, text='', bg=C['bg'], fg=C['desc'], font=('Arial', 8))
    count_lbl.pack(side='left', padx=6)

    def _update_count(*args):
        try:
            t = thr_var.get()
        except Exception:
            return
        try:
            schools   = _qualifying_schools(df_ds, df_nip, t)
            with_e    = sum(1 for s in schools if s[2])
            without_e = len(schools) - with_e
            txt = f'→ {len(schools)} σχολεία'
            if without_e:
                txt += f'  ({without_e} χωρίς email)'
            count_lbl.config(text=txt)
        except Exception:
            pass

    thr_var.trace_add('write', _update_count)
    _update_count()

    # Τρόπος αποστολής
    mode_f = tk.Frame(body_f, bg=C['bg'])
    mode_f.pack(fill='x', **pad)
    mode_var = tk.StringVar(value='schools')
    fe = getattr(config, 'FROM_EMAIL', '') or '...'
    tk.Radiobutton(mode_f, text='Αποστολή σε σχολεία (ένα email ανά σχολείο)',
                   variable=mode_var, value='schools',
                   bg=C['bg'], selectcolor=C['sel_bg'],
                   activebackground=C['bg'], font=('Arial', 9)).pack(anchor='w')
    tk.Radiobutton(mode_f, text=f'Test mode — όλα στο: {fe}',
                   variable=mode_var, value='test',
                   bg=C['bg'], selectcolor=C['sel_bg'],
                   activebackground=C['bg'], font=('Arial', 9)).pack(anchor='w')

    # Log
    log_w = scrolledtext.ScrolledText(body_f, height=10, font=('Courier', 8),
                                       wrap='word', relief='solid', bd=1,
                                       state='disabled')
    log_w.pack(fill='both', expand=True, padx=14, pady=8)

    def _log(msg):
        def _append():
            log_w.config(state='normal')
            log_w.insert('end', msg + '\n')
            log_w.see('end')
            log_w.config(state='disabled')
        try:
            dlg.after(0, _append)
        except Exception:
            pass

    # Κουμπιά
    btn_f = tk.Frame(body_f, bg=C['bg'])
    btn_f.pack(fill='x', padx=14, pady=(0, 12))
    send_btn = tk.Button(btn_f, text='✉  Αποστολή',
                         bg=C['btn_bg'], fg=C['btn_fg'],
                         font=('Arial', 9, 'bold'), relief='flat',
                         padx=14, pady=5, cursor='hand2')
    send_btn.pack(side='left', padx=(0, 8))
    tk.Button(btn_f, text='Κλείσιμο',
              bg=C['bg2'], fg=C['desc'],
              font=('Arial', 9), relief='flat', padx=14, pady=5,
              cursor='hand2', command=dlg.destroy).pack(side='left')

    def _start():
        send_btn.config(state='disabled')
        try:
            threshold = thr_var.get()
        except Exception:
            threshold = 0
        dry_run       = (mode_var.get() == 'test')
        subj, body_tpl = _load_tmimata_email_template()

        def _worker():
            try:
                sent, failed = _do_send_emails(
                    df_ds, df_nip, threshold, dry_run,
                    subj, body_tpl, config, today, _log)
                _log(f'\n{"─"*40}')
                _log(f'✓ Εστάλησαν: {len(sent)}   ✗ Αποτυχίες: {len(failed)}')
            except Exception as e:
                _log(f'✗ Σφάλμα: {e}')
            finally:
                try:
                    dlg.after(0, lambda: send_btn.config(state='normal'))
                except Exception:
                    pass

        threading.Thread(target=_worker, daemon=True).start()

    send_btn.config(command=_start)

    dlg.update_idletasks()
    x = root.winfo_x() + (root.winfo_width()  - 560) // 2
    y = root.winfo_y() + (root.winfo_height() - 480) // 2
    dlg.geometry(f'560x480+{x}+{y}')
    dlg.lift()
    dlg.focus_force()
    dlg.wait_window()


# ── Λήψη δεδομένων (μέσα στον έλεγχο — απαιτεί αλλαγή σχολικού έτους) ───────
def _with_year(entry, year, idx=11):
    """Επιστρέφει αντίγραφο ενός REPORTS tuple με το report_year ορισμένο στη θέση idx."""
    entry = list(entry)
    while len(entry) <= idx:
        entry.append(None)
    entry[idx] = year
    return tuple(entry)


def _download_inputs(config, log=print):
    """
    Κατεβάζει 3.1, 5.3, 5.4 απευθείας, αφού πρώτα οριστεί το σχολικό έτος
    SCHOOL_YEAR (2026-2027) στο MySchool. Το 3.1 δεν έχει προεπιλεγμένο
    report_year στο core/downloader.py (το χρησιμοποιούν και άλλα εργαλεία
    με το τρέχον έτος) — εδώ γίνεται τοπικό override μόνο για αυτή τη λήψη,
    χωρίς να πειραχτεί η καθολική ρύθμιση.

    Επιστρέφει (path_31, path_53, path_54) — None όπου δεν κατέβηκε.
    """
    import core.downloader as _dl

    ms_user = getattr(config, 'MYSCHOOL_USER', '').strip()
    ms_pass = getattr(config, 'MYSCHOOL_PASS', '').strip()
    if not ms_user or not ms_pass:
        raise RuntimeError('Συμπλήρωσε username και κωδικό MySchool στις Ρυθμίσεις (⚙).')

    orig_31 = next((r for r in _dl.REPORTS if r[0] == '3.1'), None)
    orig_22 = next((r for r in _dl.REPORTS if r[0] == '2.2'), None)
    if orig_31 is None:
        raise RuntimeError('Δεν βρέθηκε ρύθμιση λήψης για το 3.1 στο core/downloader.py.')

    custom_reports = (
        [_with_year(orig_31, SCHOOL_YEAR)] +
        [r for r in _dl.REPORTS if r[0] in ('5.3', '5.4')] +
        ([orig_22] if orig_22 else [])
    )
    rids = ['3.1', '5.3', '5.4'] + (['2.2'] if orig_22 else [])

    today_str = datetime.today().strftime('%Y%m%d')
    dest_dir  = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks',
                             'downloads', f'{today_str}_{SCHOOL_YEAR}')
    os.makedirs(dest_dir, exist_ok=True)

    orig_reports_backup = _dl.REPORTS
    try:
        _dl.REPORTS = custom_reports
        dl = _dl.MySchoolDownloader(
            username=ms_user, password=ms_pass, dest_dir=dest_dir,
            callback=log, reports=rids,
            browser=getattr(config, 'BROWSER', 'chrome'),
        )
        results = dl.run()
    finally:
        _dl.REPORTS = orig_reports_backup

    return results.get('3.1'), results.get('5.3'), results.get('5.4'), results.get('2.2')


# ── CUSTOM RUN ────────────────────────────────────────────────────────────
def run(config):
    import core.framework as _fw
    _fw._current_check_title = CHECK_TITLE

    print('=' * 65)
    print(f'  {CHECK_TITLE}')
    print('=' * 65)

    print(f'\n  Λήψη 3.1 / 5.3 / 5.4 για το σχολικό έτος {SCHOOL_YEAR}...')
    print('  (περιλαμβάνει αυτόματη αλλαγή σχολικού έτους — μπορεί να πάρει 1-2 λεπτά)')
    print('-' * 65)
    try:
        path_31, path_53, path_54, path_22 = _download_inputs(config, log=print)
    except Exception as e:
        import tkinter.messagebox as _mb
        _mb.showerror('Σφάλμα λήψης', str(e))
        return

    missing = [rid for rid, p in (('3.1', path_31), ('5.3', path_53), ('5.4', path_54)) if not p]
    if missing:
        import tkinter.messagebox as _mb
        _mb.showerror(
            'Λείπουν αρχεία',
            f'Δεν κατέβηκαν: {", ".join(missing)}.\n\n'
            f'Δες το log (run_log.txt στον φάκελο λήψης) για λεπτομέρειες.'
        )
        return

    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

    print(f'\n  Ημερομηνία : {today.strftime("%d/%m/%Y")}')
    print('-' * 65)

    if path_22:
        print(f'  ✓ stat2_2: {os.path.basename(path_22)}')
    else:
        print('  ℹ stat2_2 δεν κατέβηκε — Email Σχολείου θα είναι κενό')

    print('\nΕπεξεργασία...')
    try:
        df_ds, df_nip, perif = process(path_31, path_53, path_54, path_22=path_22)
    except Exception as e:
        import tkinter.messagebox as _mb
        _mb.showerror('Σφάλμα επεξεργασίας', str(e))
        return

    dev_ds  = int((df_ds['Διαφορά Τμήματα']  != 0).sum() + (df_ds['Διαφορά Μαθητές']  != 0).sum()) if not df_ds.empty  else 0
    dev_nip = int((df_nip['Διαφορά Τμήματα'] != 0).sum() + (df_nip['Διαφορά Μαθητές'] != 0).sum()) if not df_nip.empty else 0
    print(f'  ✓ Δημοτικά    : {len(df_ds)} σχολεία, {dev_ds} αποκλίσεις')
    print(f'  ✓ Νηπιαγωγεία : {len(df_nip)} σχολεία, {dev_nip} αποκλίσεις')

    if df_ds.empty and df_nip.empty:
        _show_results_popup(
            CHECK_TITLE,
            f'Ημερομηνία ελέγχου: {today.strftime("%d/%m/%Y")}\n\n'
            f'✓  Δεν βρέθηκαν σχολεία στα αρχεία 5.3/5.4.',
            result_type='ok'
        )
        return

    _docs   = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks')
    out_dir = os.path.join(_docs, f'results_{today.strftime("%Y%m%d")}', RESULTS_FOLDER)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'{today.strftime("%Y%m%d")}_{RESULTS_FOLDER}.xlsx')

    try:
        build_workbook(df_ds, df_nip, perif, today, out_path)
        print(f'\n  ✓ Αποθηκεύτηκε: {os.path.basename(out_path)}')
    except PermissionError:
        import tkinter.messagebox as _mb
        _mb.showwarning(
            'Αρχείο ανοιχτό',
            f'Το αρχείο {os.path.basename(out_path)} είναι ανοιχτό σε άλλο πρόγραμμα.\n'
            f'Κλείστε το και τρέξτε ξανά τον έλεγχο.'
        )
        return

    body = (
        f'Σύνοψη ελέγχου τμημάτων γενικής παιδείας — {today.strftime("%d/%m/%Y")}\n'
        f'{"─"*50}\n'
        f'Δημοτικά:     {len(df_ds)} σχολεία  ·  {dev_ds} αποκλίσεις\n'
        f'Νηπιαγωγεία:  {len(df_nip)} σχολεία  ·  {dev_nip} αποκλίσεις\n\n'
        f'{"─"*50}\n'
        f'Αποτελέσματα αποθηκεύτηκαν στο φάκελο:\n{out_dir}'
    )
    _show_results_popup(CHECK_TITLE, body, result_type='warn', excel_path=out_path)
    _show_email_dialog(config, df_ds, df_nip, today)
