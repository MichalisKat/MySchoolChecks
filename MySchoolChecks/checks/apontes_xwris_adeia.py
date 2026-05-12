"""
checks/aportes_xwris_adeia.py
══════════════════════════════
Απουσία στην Τοποθέτηση χωρίς Δήλωση Άδειας (4.16 + 4.21).
Εντοπίζει εκπαιδευτικούς που έχουν καταχωρηθεί ως Μακροχρόνια Απόντες
στο 4.16 (εξαιρουμένων Άνευ Αποδοχών) χωρίς ενεργή άδεια στο 4.21.
"""

from datetime import datetime
import os
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

SCHOOL_COLUMN = 'Κωδικός Σχολείου'
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


# ── Επιλογή Διεύθυνσης ──────────────────────────────────────────────────────
def _ask_periochi(path_416):
    """Φορτώνει μοναδικές τιμές Περιοχής Μετάθεσης από 4.16 και ζητά επιλογή."""
    import tkinter as tk
    from tkinter import ttk

    df = read_csv_fixed(path_416)
    df.columns = [c.strip() for c in df.columns]
    per_col = next((c for c in df.columns if 'Περιοχή Μετάθεσης Εκπαιδευτικού' in c), None)
    if per_col is None:
        print('  ⚠ Δεν βρέθηκε στήλη Περιοχής Μετάθεσης — χωρίς φίλτρο Διεύθυνσης.')
        return None

    perioches = sorted(df[per_col].dropna().astype(str).str.strip()
                       .replace('', None).dropna().unique())
    if not perioches:
        return None

    result = [perioches[0]]
    win = tk.Toplevel()
    win.title('Επιλογή Διεύθυνσης')
    win.configure(bg='#EEF4F0')
    win.resizable(False, False)
    win.grab_set()
    win.attributes('-topmost', True)
    _ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'app.ico')
    if os.path.exists(_ico):
        try: win.iconbitmap(_ico)
        except Exception: pass
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f'380x200+{sw//2-190}+{sh//2-100}')

    # ── Header παραμετροποίησης (ίδιο με τις άλλες φόρμες) ──────────
    import core.framework as _fw
    tk.Label(win, text='ΠΑΡΑΜΕΤΡΟΠΟΙΗΣΗ ΕΛΕΓΧΟΥ',
             bg='#1F4E79', fg='white',
             font=('Arial', 9, 'bold'), pady=5).pack(fill='x')
    if _fw._current_check_title:
        tk.Label(win, text=_fw._current_check_title,
                 bg='#EEF4F0', fg='#1F4E79',
                 font=('Arial', 9, 'bold'), pady=4).pack(fill='x', padx=10)
    tk.Frame(win, bg='#C5D8E8', height=1).pack(fill='x', padx=10)

    tk.Label(win, text='Επιλέξτε Διεύθυνση (Περιοχή Μετάθεσης):',
             bg='#EEF4F0', fg='#1F4E79',
             font=('Arial', 10, 'bold'), pady=10).pack()

    var = tk.StringVar(value=perioches[0])
    combo = ttk.Combobox(win, textvariable=var, values=perioches,
                         width=44, state='readonly', font=('Arial', 9))
    combo.pack(padx=18)

    def confirm():
        result[0] = var.get()
        win.destroy()

    tk.Button(win, text='OK', font=('Arial', 10, 'bold'),
              bg='#1F4E79', fg='white', relief='flat',
              padx=20, pady=6, cursor='hand2',
              command=confirm).pack(pady=12)

    win.wait_window()
    print(f'  ✓ Διεύθυνση: {result[0]}')
    return result[0]


# ── Είσοδος ─────────────────────────────────────────────────────────────────
def ask_inputs():
    path_416 = get_downloaded_file('4.16', 'Αρχείο 4.16 (Αιτιολόγηση Απουσίας) [.csv]:', csv_only=True, silent=True)
    path_421 = get_downloaded_file('4.21', 'Αρχείο 4.21 (Άδειες) [.csv]:', csv_only=True, silent=True)
    if path_416 is None or path_421 is None:
        return {'path_416': path_416, 'path_421': path_421, 'today': None}
    periochi = _ask_periochi(path_416)
    today = ask_date_yyyymmdd()
    return {'path_416': path_416, 'path_421': path_421, 'today': today, 'periochi': periochi}


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

    # Εξαίρεση Ιδιωτικών (βάσει Είδους Σχολείου)
    eid_sxol_col = next((c for c in df16.columns if 'Είδος Σχολείου' in c), None)
    if eid_sxol_col:
        df16 = df16[~df16[eid_sxol_col].astype(str).str.contains('Ιδιωτικ', na=False)].copy()
    print(f'  → {len(df16)} μετά εξαίρεση Ιδιωτικών')

    # Εξαίρεση ΣΔΕΥ
    if eid_col:
        df16 = df16[~df16[eid_col].astype(str).str.strip().isin(EXCLUDED_SPECIALTIES)].copy()
    print(f'  → {len(df16)} μετά εξαίρεση ΣΔΕΥ')

    # Φίλτρο περιοχής μετάθεσης εκπαιδευτικού
    periochi = ctx.get('periochi')
    if per_col and periochi:
        df16 = df16[df16[per_col].astype(str).str.strip() == periochi].copy()
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

    return out.sort_values(['Κωδικός Σχολείου', 'Επώνυμο', 'Όνομα']).reset_index(drop=True)


def test_body(df_out, today, schools):
    sep = '─' * 50
    return (
        f'Σύνοψη ελέγχου απουσίας χωρίς δήλωση άδειας — {today.strftime("%d/%m/%Y")}\n'
        f'{sep}\n'
        f'Βρέθηκαν: {len(df_out)} εκπαιδευτικοί με μακροχρόνια απουσία χωρίς ενεργή άδεια\n'
        f'Σχολεία που εμφανίζονται ({len(schools)}): {", ".join(sorted(str(s) for s in schools))}'
    )


def _find_col(df, *keywords):
    """Βρίσκει στήλη με ακριβές ή case-insensitive contains match."""
    for kw in keywords:
        if kw in df.columns:
            return kw
        for c in df.columns:
            if kw.lower() in c.lower():
                return c
    return None


def custom_full_send(config, today, out_dir, scol, ecol, subject, body_template,
                     cols, ccols, title):
    """Κανονική αποστολή: ζητά επεξεργασμένο Excel, το σπάει ανά σχολείο και στέλνει."""
    import tkinter as tk
    from tkinter import filedialog, messagebox
    from core.framework import send_email, save_workbook

    # Ζητά το επεξεργασμένο αρχείο Excel
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path_xl = filedialog.askopenfilename(
        title='Επιλέξτε το επεξεργασμένο αρχείο Excel για αποστολή',
        filetypes=[('Excel', '*.xlsx *.xls')],
        parent=root
    )
    root.destroy()
    if not path_xl:
        print('  ✗ Δεν επιλέχθηκε αρχείο — αποστολή ακυρώθηκε.')
        return

    # Φόρτωση — το αρχείο έχει merged title rows, οι επικεφαλίδες μπορεί να είναι σε γραμμή 1, 2 ή 3
    print(f'  Φόρτωση: {os.path.basename(path_xl)}')
    try:
        import pandas as pd
        df = None
        # Δοκιμάζουμε header=0,1,2 και κρατάμε αυτό που έχει 'Κωδικός' ή 'Ονομασία' στις στήλες
        for hdr_row in range(4):
            try:
                df_try = pd.read_excel(path_xl, dtype=str, header=hdr_row)
                df_try.columns = [str(c).strip() for c in df_try.columns]
                cols_lower = [c.lower() for c in df_try.columns]
                if any('κωδικός σχολείου' in c or 'ονομασία σχολείου' in c for c in cols_lower):
                    df = df_try
                    print(f'  Βρέθηκαν επικεφαλίδες στη γραμμή {hdr_row + 1}')
                    break
            except Exception:
                continue
        if df is None:
            # Fallback: φόρτωση από γραμμή 1
            df = pd.read_excel(path_xl, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        messagebox.showerror('Σφάλμα', f'Αδυναμία φόρτωσης αρχείου:\n{e}')
        return

    print(f'  Στήλες Excel: {list(df.columns)}')

    # Αναζήτηση στήλης split (Κωδικός ή Ονομασία) — case-insensitive
    split_col = _find_col(df, 'Κωδικός Σχολείου', 'Ονομασία Σχολείου')
    if split_col is None:
        messagebox.showerror('Σφάλμα',
            f'Δεν βρέθηκε στήλη "Κωδικός Σχολείου" ή "Ονομασία Σχολείου" στο αρχείο.\n\n'
            f'Στήλες που βρέθηκαν:\n{", ".join(df.columns)}')
        return

    # Στήλες ονόματος και email — case-insensitive
    name_col    = _find_col(df, 'Ονομασία Σχολείου') or split_col
    ecol_actual = _find_col(df, ecol, 'Email Σχολείου', 'Email') or ecol

    school_codes = sorted(df[split_col].dropna().unique())
    print(f'  Βρέθηκαν {len(school_codes)} σχολεία ({split_col}), {len(df)} εγγραφές.')

    # Split ανά σχολείο + αποστολή
    ok = fail = 0
    sent_school_names = []
    for code in school_codes:
        df_s = df[df[split_col] == code].copy()
        school_name = str(df_s[name_col].iloc[0]).strip() if name_col in df_s.columns else str(code)
        email_s = ''
        if ecol_actual in df_s.columns:
            email_s = str(df_s[ecol_actual].iloc[0]).strip()
        if not email_s or email_s in ('', 'nan', 'None'):
            print(f'  ⚠  [{code}] {school_name[:45]} — ΔΕΝ ΥΠΑΡΧΕΙ EMAIL, παράλειψη.')
            fail += 1
            continue

        safe_name = ''.join(c for c in school_name if c not in r'\/:*?"<>|').strip()[:55]
        path_s = os.path.join(out_dir, f'{today.strftime("%Y%m%d")}_{code}_{safe_name}.xlsx')
        try:
            save_workbook(df_s, title, cols, ccols, today, path_s,
                          subtitle_extra=f'  |  {school_name}')
            body = body_template(school_name) if callable(body_template) else body_template
            send_email(config, [email_s], subject, body, path_s)
            print(f'  ✓ [{code}] {school_name[:45]} → {email_s}')
            ok += 1
            sent_school_names.append(school_name)
        except Exception as e:
            print(f'  ✗ [{code}] {school_name[:45]} → {e}')
            fail += 1

    print(f'\n  Αποστολές: {ok} επιτυχείς, {fail} αποτυχίες')
    if ok > 0:
        from core.framework import _send_notify
        _send_notify(config, 'Απουσία χωρίς Δήλωση Άδειας', today, ok, sent_school_names)
