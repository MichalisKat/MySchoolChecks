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

GRADES_DS  = ['Α', 'Β', 'Γ', 'Δ', 'Ε', 'ΣΤ']          # Δημοτικά
NIP_GROUP  = 'ΠΡΟΝΗΠΙΑ-ΝΗΠΙΑ'                          # Νηπιαγωγεία — ενιαία στήλη

COLOR_HDR   = '1F4E79'
COLOR_SUB   = 'D6E4F0'
COLOR_ALT   = 'EBF3FB'
COLOR_OK    = '92D050'   # πράσινο — διαφορά 0
COLOR_DEV   = 'FF0000'   # κόκκινο — διαφορά ≠ 0
DEFAULT_PERIFEREIA = 'ΚΕΝΤΡΙΚΗΣ ΜΑΚΕΔΟΝΙΑΣ'

_LEFT_ALIGN_COLS = {'Τύπος', 'Ονομασία', 'Δήμος'}


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
def _build_records(df_src, grade_lookup, multi_grade, dimos_lookup=None):
    """
    df_src        : DataFrame από 5.3 ή 5.4
    grade_lookup  : dict {code: {grade: (tm, ma)}}  (Δημοτικά) ή
                     dict {code: (tm, ma)}           (Νηπιαγωγεία)
    multi_grade   : True → Δημοτικά (πολλαπλές τάξεις Α-ΣΤ)
                     False → Νηπιαγωγεία (μία ενιαία στήλη)
    dimos_lookup  : dict {code: dimos} από το 3.1 (προαιρετικό)
    """
    if dimos_lookup is None:
        dimos_lookup = {}
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
def process(path_31, path_53, path_54):
    """Επιστρέφει (df_ds, df_nip, periфereia_name)."""
    df31 = _load_stat31(path_31)
    df53 = pd.read_excel(path_53)
    df54 = pd.read_excel(path_54)

    dimos = _dimos_lookup(df31)
    df_ds  = _build_records(df54, _grade_lookup_ds(df31),  multi_grade=True,  dimos_lookup=dimos)
    df_nip = _build_records(df53, _grade_lookup_nip(df31), multi_grade=False, dimos_lookup=dimos)

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
    if orig_31 is None:
        raise RuntimeError('Δεν βρέθηκε ρύθμιση λήψης για το 3.1 στο core/downloader.py.')

    custom_reports = (
        [_with_year(orig_31, SCHOOL_YEAR)] +
        [r for r in _dl.REPORTS if r[0] in ('5.3', '5.4')]
    )

    today_str = datetime.today().strftime('%Y%m%d')
    dest_dir  = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks',
                             'downloads', f'{today_str}_{SCHOOL_YEAR}')
    os.makedirs(dest_dir, exist_ok=True)

    orig_reports_backup = _dl.REPORTS
    try:
        _dl.REPORTS = custom_reports
        dl = _dl.MySchoolDownloader(
            username=ms_user, password=ms_pass, dest_dir=dest_dir,
            callback=log, reports=['3.1', '5.3', '5.4'],
            browser=getattr(config, 'BROWSER', 'chrome'),
        )
        results = dl.run()
    finally:
        _dl.REPORTS = orig_reports_backup

    return results.get('3.1'), results.get('5.3'), results.get('5.4')


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
        path_31, path_53, path_54 = _download_inputs(config, log=print)
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

    print('\nΕπεξεργασία...')
    try:
        df_ds, df_nip, perif = process(path_31, path_53, path_54)
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
