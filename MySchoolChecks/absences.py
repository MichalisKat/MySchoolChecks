#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
absences.py
============
Αυτόματη καταχώρηση απουσίας (Ολική Διάθεση ή Απόσπαση) σε οργανική
τοποθέτηση στο MySchool.

Πλαίσιο:
  Το αρχείο εξαγωγής έχει ΜΙΑ γραμμή ανά εκπαιδευτικό, με στήλη
  "Σχέση τοποθέτησης" ένα από τα δύο (2η τριάδα — πηγή στοιχείων απουσίας):
    - "Ολική Διάθεση (ανάγκες υπηρεσίας - κύριος φορέας)"
      → λεκτικό απουσίας: ΟΛΙΚΗ ΔΙΑΘΕΣΗ ΣΕ ΑΛΛΗ ΣΧΟΛΙΚΗ ΜΟΝΑΔΑ - Σχολικές
        Μονάδες Πρωτοβάθμιας
    - "Απόσπαση (με αίτηση - κύριος φορέας)"
      → λεκτικό απουσίας: ΑΠΟΣΠΑΣΗ ΣΕ ΣΧΟΛΙΚΗ ΜΟΝΑΔΑ ΕΝΤΟΣ ΤΟΥ ΠΥΣΔΕ / ΠΥΣΠΕ -
        Αλλη Σχολική Μονάδα
  Το «Επί Θητεία» παραμένει εκτός (χειροκίνητη καταχώρηση). Το αρχείο ΔΕΝ
  χρειάζεται να περιέχει καθόλου εγγραφή «Οργανικά» — η οργανική τοποθέτηση
  δεν διαβάζεται από το αρχείο, εντοπίζεται απευθείας στο MySchool (βλ.
  παρακάτω, βήμα 2).

  1η τριάδα (εκεί μπαίνει η απουσία — αναζητείται στο MySchool, όχι στο
  αρχείο): Οργανικά / Οργανικά σε Τμήμα Ένταξης / Οργανικά από Αρση
  Υπεραριθμίας.

Ροή ανά εκπαιδευτικό:
  1. Αναζήτηση με Α.Μ. στο Worker.list.myEmplUnit.aspx.
  2. Από ΟΛΑ τα αποτελέσματα που επιστρέφει η αναζήτηση, εντοπισμός της
     γραμμής με Σχέση τοποθέτησης από την 1η τριάδα (Οργανικά κ.λπ.) → κλικ
     γρανάζι. Αν ΔΕΝ υπάρχει τέτοια γραμμή ανάμεσα στα αποτελέσματα, ο
     εκπαιδευτικός αφήνεται στην άκρη (καμία ενόχληση με popup — απλή
     καταγραφή στη σύνοψη στο τέλος, βλ. "ΔΕΝ ΒΡΕΘΗΚΕ ΟΡΓΑΝΙΚΗ").
  3. Scroll στο πινακάκι Απουσιών (gridAbsences) → κλικ ➕ (Προσθήκη)
  4. Συμπλήρωση:
       Τύπος απουσίας: ανάλογα με τη Σχέση τοποθέτησης του αρχείου (βλ.
         πλαίσιο παραπάνω)
       Ισχύει από: από τη στήλη "Από" της εγγραφής του αρχείου
       Ισχύει έως: από τη στήλη "Έως" της εγγραφής του αρχείου
     Μετά τη συμπλήρωση ξαναδιαβάζονται οι τιμές των πεδίων. Αν ΚΑΠΟΙΟ πεδίο
     δεν έχει τη σωστή τιμή, η νέα γραμμή ακυρώνεται και ΔΕΝ γίνεται
     Αποθήκευση (ΑΠΟΤΥΧΙΑ) — ποτέ μισή/λάθος εγγραφή στο MySchool.
  5. Κλικ ✓ (Αποδοχή) — έλεγχος ότι η γραμμή έκλεισε όντως → Αποθήκευση
  6. Επαλήθευση ΜΕΤΑ την αποθήκευση: η νέα απουσία (τύπος + από + έως)
     πρέπει να φαίνεται στο grid Απουσιών. Αλλιώς ΑΠΟΤΥΧΙΑ (ξαναδοκιμάζεται).
  7. Επόμενος

  Για οτιδήποτε ΑΛΛΟ διαφοροποιείται από το αναμενόμενο (δεν βρέθηκε
  καθόλου ο εκπαιδευτικός, dropdown/πεδίο απέτυχε κ.λπ.) εμφανίζεται μήνυμα
  στον χρήστη — ο χρήστης το κλείνει και συνεχίζει στον επόμενο
  εκπαιδευτικό. Μόνο η περίπτωση «δεν βρέθηκε οργανική» (βήμα 2) ΔΕΝ
  διακόπτει με popup, ώστε να μη χρειάζεται κλικ σε καθέναν από πιθανώς
  πολλούς τέτοιους εκπαιδευτικούς.

Resume & προστασία από διπλοεγγραφές:
  - Δίπλα στο αρχείο δημιουργείται αρχείο προόδου "<όνομα>_status.xlsx" που
    ενημερώνεται μετά από ΚΑΘΕ εκπαιδευτικό. Αν το script σταματήσει (crash,
    κλείσιμο, διακοπή δικτύου), μια επόμενη εκτέλεση με το ΙΔΙΟ αρχείο δεν
    ξαναπερνάει όσους έχουν ήδη καταχωρηθεί (ΕΠΙΤΥΧΙΑ/ΥΠΑΡΧΕΙ ΗΔΗ) — μόνο
    όσους έμειναν σε ΑΠΟΤΥΧΙΑ ή δεν προλάβαιναν να τρέξουν.
  - Επιπλέον, πριν προσθέσει νέα γραμμή απουσίας, ελέγχει το ίδιο το grid
    Απουσιών στο MySchool — αν βρει ήδη γραμμή με ΙΔΙΟ τύπο ΚΑΙ ίδιες
    ημερομηνίες Από/Έως (π.χ. από προηγούμενη εκτέλεση που δεν πρόλαβε να
    γραφτεί στο αρχείο προόδου), δεν την ξαναπροσθέτει, το καταγράφει ως
    "ΥΠΑΡΧΕΙ ΗΔΗ" και προχωράει. Απουσία ίδιου τύπου με ΑΛΛΕΣ ημερομηνίες
    (π.χ. περσινή) ΔΕΝ θεωρείται διπλοεγγραφή.
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Καταχώρηση Απουσίας σε Οργανική'
CHECK_DESCRIPTION = 'Αυτόματη καταχώρηση απουσίας (Ολική Διάθεση / Απόσπαση) στο MySchool'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL   = 'https://app.myschool.sch.gr'
SEARCH_URL = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

AM_FIELD_ID = 'ctl00_ContentData_txtRegistryNo_I'

# 1η τριάδα — εκεί μπαίνει η απουσία
FIRST_TRIAD = {
    'Οργανικά',
    'Οργανικά σε Τμήμα Ένταξης',
    'Οργανικά από Αρση Υπεραριθμίας',
}

# 2η τριάδα — αποδεκτές τιμές (Σχέση τοποθέτησης) → αντίστοιχο λεκτικό
# τύπου απουσίας στο dropdown MySchool. Το «Επί Θητεία» ΔΕΝ περιλαμβάνεται
# (παραμένει χειροκίνητη καταχώρηση).
SECOND_TRIAD_ABSENCE_TYPE = {
    'Ολική Διάθεση (ανάγκες υπηρεσίας - κύριος φορέας)':
        'ΟΛΙΚΗ ΔΙΑΘΕΣΗ ΣΕ ΑΛΛΗ ΣΧΟΛΙΚΗ ΜΟΝΑΔΑ - Σχολικές Μονάδες Πρωτοβάθμιας',
    'Απόσπαση (με αίτηση - κύριος φορέας)':
        'ΑΠΟΣΠΑΣΗ ΣΕ ΣΧΟΛΙΚΗ ΜΟΝΑΔΑ ΕΝΤΟΣ ΤΟΥ ΠΥΣΔΕ / ΠΥΣΠΕ - Αλλη Σχολική Μονάδα',
}

GRID_ID = 'ctl00_ContentData_gridAbsences'


# ── Ημερομηνίες ───────────────────────────────────────────────────────────────

def _to_timestamp(val):
    """Δέχεται Timestamp/datetime ή κείμενο ('1/9/2026', '01-09-2026',
    '2026-09-01') και επιστρέφει pandas Timestamp ή None."""
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, pd.Timestamp):
        return val
    try:
        import datetime as _dt
        if isinstance(val, (_dt.datetime, _dt.date)):
            return pd.Timestamp(val)
    except Exception:
        pass
    txt = str(val).strip()
    if not txt or txt.lower() in ('nan', 'none', 'nat'):
        return None
    import re
    if re.match(r'^\d{4}-\d{1,2}-\d{1,2}', txt):          # ISO: έτος πρώτα
        ts = pd.to_datetime(txt, errors='coerce')
    else:
        ts = pd.to_datetime(txt, dayfirst=True, errors='coerce')
    return None if pd.isna(ts) else ts


def _fmt_date(val):
    """Μετατρέπει ημερομηνία (Timestamp ή κείμενο) σε 'D/M/YYYY' (χωρίς
    μηδενικά πρόθεμα). Κενό string αν δεν είναι έγκυρη ημερομηνία."""
    ts = _to_timestamp(val)
    if ts is None:
        return ''
    return f'{ts.day}/{ts.month}/{ts.year}'


def _date_key(date_str):
    """'1/9/2026' → (1, 9, 2026) — για σύγκριση ανεξάρτητα από μηδενικά/διαχωριστικό."""
    ts = _to_timestamp(date_str)
    return None if ts is None else (ts.day, ts.month, ts.year)


def _dates_in_text(text):
    """Όλες οι ημερομηνίες (d, m, y) που εμφανίζονται σε ένα κείμενο — δέχεται
    '/', '-' ή '.' ως διαχωριστικό και με/χωρίς μηδενικά («01-09-2026», «1/9/2026»)."""
    import re
    out = set()
    for d, m, y in re.findall(r'(?<!\d)(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})(?!\d)', text or ''):
        out.add((int(d), int(m), int(y)))
    return out


def _normalize_text(s):
    """Αφαιρεί τόνους, κάνει κεφαλαία, συμπτύσσει κενά — για ανεκτική σύγκριση."""
    import re
    import unicodedata
    if not s:
        return ''
    s = unicodedata.normalize('NFD', str(s))
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return re.sub(r'\s+', ' ', s.upper()).strip()


# ── Ανάγνωση & φιλτράρισμα αρχείου ────────────────────────────────────────────


def load_data(file_path, log=print):
    """
    Διαβάζει το αρχείο εξαγωγής τοποθετήσεων και επιστρέφει λίστα εγγραφών
    προς επεξεργασία: μόνο εκπαιδευτικοί με Σχέση τοποθέτησης σε
    SECOND_TRIAD_ABSENCE_TYPE (Ολική Διάθεση ή Απόσπαση). Το αρχείο ΔΕΝ
    χρειάζεται να περιέχει γραμμή «Οργανικά» — αυτή εντοπίζεται αργότερα
    απευθείας στο MySchool (βλ. run()).

    Επιστρέφει: [{'am', 'afm', 'name', 'apo', 'eos', 'absence_type'}], skipped: [str, ...]
    """
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(file_path)
        else:
            for enc in ('utf-8', 'iso-8859-7', 'cp1253'):
                try:
                    df = pd.read_csv(file_path, sep=';', encoding=enc)
                    break
                except Exception:
                    pass
            else:
                df = pd.read_csv(file_path)
    except Exception as e:
        log(f'Σφάλμα ανάγνωσης αρχείου: {e}')
        return [], []

    cols = {c.strip(): c for c in df.columns}
    required = ['Α.Μ.', 'Σχέση τοποθέτησης', 'Από', 'Έως']
    missing = [c for c in required if c not in cols]
    if missing:
        log(f'Λείπουν στήλες: {missing}. Διαθέσιμες: {list(cols.keys())}')
        return [], []

    am_col    = cols['Α.Μ.']
    rel_col   = cols['Σχέση τοποθέτησης']
    apo_col   = cols['Από']
    eos_col   = cols['Έως']
    afm_col   = cols.get('Α.Φ.Μ.')
    name_col  = cols.get('Επώνυμο') or cols.get('ΕΠΙΘΕΤΟ') or cols.get('Επίθετο')
    fname_col = cols.get('Όνομα')   or cols.get('ΟΝΟΜΑ')

    records = []
    skipped = []

    # Α.Μ. ως κείμενο χωρίς '.0' (αν το Excel το διάβασε ως αριθμό float)
    def _am_text(v):
        t = str(v).strip()
        return t[:-2] if t.endswith('.0') and t[:-2].isdigit() else t
    df = df[df[am_col].notna()].copy()
    df[am_col] = df[am_col].map(_am_text)
    df = df[df[am_col] != '']

    for am, grp in df.groupby(am_col):
        # Κρατάμε μόνο τη γραμμή με Σχέση τοποθέτησης Ολική Διάθεση/Απόσπαση.
        # ΔΕΝ απαιτούμε πλέον να υπάρχει και γραμμή «Οργανικά» στο ίδιο
        # αρχείο — αυτή αναζητείται απευθείας στο MySchool στο run().
        second_row = grp[grp[rel_col].isin(SECOND_TRIAD_ABSENCE_TYPE.keys())]

        epon = str(grp.iloc[0].get(name_col, '')).strip()  if name_col  else ''
        onom = str(grp.iloc[0].get(fname_col, '')).strip() if fname_col else ''
        full_name = f'{epon} {onom}'.strip()
        am_str = str(am).strip()

        if second_row.empty:
            skipped.append(f'{full_name} ({am_str}) — Σχέση τοποθέτησης δεν είναι '
                            f'Ολική Διάθεση/Απόσπαση')
            continue

        second_rel_value = str(second_row.iloc[0][rel_col]).strip()
        absence_type = SECOND_TRIAD_ABSENCE_TYPE.get(second_rel_value)
        if not absence_type:
            skipped.append(f'{full_name} ({am_str}) — άγνωστη Σχέση τοποθέτησης: {second_rel_value}')
            continue

        apo = _fmt_date(second_row.iloc[0][apo_col])
        if not apo:
            skipped.append(f'{full_name} ({am_str}) — κενή ή μη έγκυρη ημ. Από')
            continue

        eos = _fmt_date(second_row.iloc[0][eos_col])
        if not eos:
            skipped.append(f'{full_name} ({am_str}) — κενή ή μη έγκυρη ημ. Έως')
            continue

        if _to_timestamp(apo) > _to_timestamp(eos):
            skipped.append(f'{full_name} ({am_str}) — ημ. Από ({apo}) μετά την ημ. Έως ({eos})')
            continue

        afm = str(grp.iloc[0].get(afm_col, '')).strip() if afm_col else ''

        records.append({
            'am':           am_str,
            'afm':          afm,
            'name':         full_name,
            'apo':          apo,
            'eos':          eos,
            'absence_type': absence_type,
        })

    log(f'Φορτώθηκαν {len(records)} εγγραφές προς καταχώρηση '
        f'({len(skipped)} παραλείφθηκαν στο φιλτράρισμα)')
    return records, skipped


# ── Αρχείο προόδου (resume — σαν στις Τοποθετήσεις) ───────────────────────────
#
# Το αρχικό αρχείο εξαγωγής (.xls από το MySchool) δεν ξαναγράφεται απευθείας
# (η εγγραφή .xls δεν υποστηρίζεται αξιόπιστα από pandas). Αντ' αυτού κρατάμε
# ένα "αρχείο προόδου" (.xlsx) δίπλα στο αρχικό, με στήλες Α.Μ./ΚΑΤΑΣΤΑΣΗ/
# ΣΧΟΛΙΟ. Ενημερώνεται μετά από ΚΑΘΕ εκπαιδευτικό, ώστε αν σταματήσει το
# script (crash, κλείσιμο, διακοπή σύνδεσης) η επόμενη εκτέλεση με το ΙΔΙΟ
# αρχείο να συνεχίζει από εκεί που έμεινε — δεν ξαναπερνάει όσους έχουν ήδη
# ΕΠΙΤΥΧΙΑ ή ΥΠΑΡΧΕΙ ΗΔΗ.

STATUS_SUFFIX  = '_status.xlsx'
DONE_STATUSES  = {'ΕΠΙΤΥΧΙΑ', 'ΥΠΑΡΧΕΙ ΗΔΗ'}


def _status_file_path(source_path):
    base, _ = os.path.splitext(source_path)
    return base + STATUS_SUFFIX


def _load_status(status_path, log=print):
    """Επιστρέφει dict Α.Μ. → {'name','status','comment'} από προηγούμενη εκτέλεση."""
    if not os.path.exists(status_path):
        return {}
    try:
        sdf = pd.read_excel(status_path, dtype=str)
        out = {}
        for _, r in sdf.iterrows():
            am = str(r.get('Α.Μ.', '')).strip()
            if am and am not in ('nan', 'None'):
                out[am] = {
                    'name':    str(r.get('Όνομα', '')).strip(),
                    'status':  str(r.get('ΚΑΤΑΣΤΑΣΗ', '')).strip(),
                    'comment': str(r.get('ΣΧΟΛΙΟ', '')).strip(),
                }
        return out
    except Exception as e:
        log(f'⚠ Δεν διαβάστηκε το αρχείο προόδου ({status_path}): {e}')
        return {}


def _save_status(status_path, status_dict):
    """Γράφει το dict προόδου στο δίσκο (κλήση μετά από ΚΑΘΕ εκπαιδευτικό)."""
    rows = []
    for am, info in status_dict.items():
        rows.append({
            'Α.Μ.':      am,
            'Όνομα':     info.get('name', ''),
            'ΚΑΤΑΣΤΑΣΗ': info.get('status', ''),
            'ΣΧΟΛΙΟ':    info.get('comment', ''),
            'ΗΜ/ΝΙΑ':    info.get('date', ''),
        })
    sdf = pd.DataFrame(rows, columns=['Α.Μ.', 'Όνομα', 'ΚΑΤΑΣΤΑΣΗ', 'ΣΧΟΛΙΟ', 'ΗΜ/ΝΙΑ'])
    try:
        sdf.to_excel(status_path, index=False)
    except Exception:
        pass  # δεν μπλοκάρουμε την εκτέλεση αν αποτύχει προσωρινά η εγγραφή


# ── Βοηθητικές Selenium ───────────────────────────────────────────────────────

STRIKE_INTERVAL = 0.3


def _send_keys_slow(element, text, delay=STRIKE_INTERVAL):
    """Πληκτρολογεί έναν-έναν χαρακτήρα με καθυστέρηση (απαραίτητο για MySchool)."""
    for char in str(text):
        element.send_keys(char)
        time.sleep(delay)


def _set_dxe_value(driver, element_id, value):
    js = """
        var inp = document.getElementById(arguments[0]);
        if (!inp) return false;
        inp.value = arguments[1];
        inp.dispatchEvent(new Event('change', {bubbles: true}));
        var base = arguments[0].replace(/_I$/, '');
        if (typeof aspxETextChanged  === 'function') aspxETextChanged(base);
        if (typeof aspxEValueChanged === 'function') aspxEValueChanged(base);
        return true;
    """
    return driver.execute_script(js, element_id, value)


def _select_dxe_combo(driver, base_id, text):
    """Επιλέγει τιμή από DevExpress ComboBox με πληκτρολόγηση + κλικ (fallback TAB)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys

    try:
        inp = driver.find_element(By.ID, base_id + '_I')
        driver.execute_script('arguments[0].click();', inp)
        time.sleep(0.5)
        inp.clear()
        time.sleep(0.3)
        _send_keys_slow(inp, text)
        time.sleep(1.5)
        try:
            item = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, f'td[dxtext="{text}"]')))
            driver.execute_script('arguments[0].click();', item)
            time.sleep(0.5)
            return True
        except Exception:
            pass
        inp.send_keys(Keys.TAB)
        time.sleep(0.5)
        return True
    except Exception:
        return False


def _grid_absence_rows(driver):
    """
    Επιστρέφει το κείμενο κάθε γραμμής δεδομένων του grid Απουσιών (όχι τη
    γραμμή επεξεργασίας/νέας εγγραφής, ούτε τη λίστα του dropdown).
    None αν δεν βρέθηκε καθόλου το grid.
    """
    from selenium.webdriver.common.by import By
    try:
        container = driver.find_element(By.ID, GRID_ID)
    except Exception:
        return None
    rows = []
    try:
        els = container.find_elements(
            By.XPATH,
            './/tr[contains(@id,"DXDataRow") or contains(@class,"dxgvDataRow")]')
        for el in els:
            rid = el.get_attribute('id') or ''
            if 'editnew' in rid.lower() or 'dxeditingrow' in (el.get_attribute('class') or '').lower():
                continue
            txt = (el.text or '').strip()
            if txt:
                rows.append(txt)
    except Exception:
        pass
    if not rows:
        # Fallback: αν δεν αναγνωρίστηκαν γραμμές, κάθε γραμμή κειμένου του
        # container αντιμετωπίζεται ως «γραμμή» (DevExpress αποδίδει μία
        # γραμμή κειμένου ανά εγγραφή).
        try:
            rows = [ln for ln in (container.text or '').splitlines() if ln.strip()]
        except Exception:
            rows = []
    return rows


def _find_matching_absence(driver, type_text, apo, eos):
    """
    Ψάχνει στο grid Απουσιών γραμμή με ΙΔΙΟ τύπο απουσίας ΚΑΙ ίδιες
    ημερομηνίες Από/Έως.

    Επιστρέφει (match, same_type_other_dates):
      match                 : True αν υπάρχει ακριβώς η ίδια απουσία
      same_type_other_dates : λίστα κειμένων γραμμών ίδιου τύπου αλλά με
                              άλλες ημερομηνίες (π.χ. περσινή) — ΔΕΝ είναι
                              διπλοεγγραφή, μόνο πληροφοριακά για το log.
    """
    rows = _grid_absence_rows(driver) or []
    t_norm = _normalize_text(type_text)
    k_apo, k_eos = _date_key(apo), _date_key(eos)
    others = []
    for row in rows:
        if t_norm not in _normalize_text(row):
            continue
        dates = _dates_in_text(row)
        if k_apo in dates and k_eos in dates:
            return True, others
        others.append(row)
    return False, others


def _combo_visible(driver, combo_full_id):
    """True όσο η νέα γραμμή είναι ακόμη σε λειτουργία επεξεργασίας (το combo
    τύπου απουσίας της φαίνεται) — δηλαδή η Αποδοχή ΔΕΝ «έπιασε»."""
    from selenium.webdriver.common.by import By
    try:
        els = driver.find_elements(By.ID, combo_full_id)
        return any(el.is_displayed() for el in els)
    except Exception:
        return False


def _cancel_new_row(driver):
    """Ακυρώνει τη νέα (μη αποθηκευμένη) γραμμή του grid Απουσιών. Ακόμη κι αν
    αποτύχει, η επόμενη πλοήγηση (driver.get) απορρίπτει ό,τι δεν
    αποθηκεύτηκε — το σημαντικό είναι ότι ΔΕΝ πατιέται Αποθήκευση."""
    try:
        driver.execute_script(
            "var g = (typeof ASPxClientGridView !== 'undefined') ? "
            "ASPxClientGridView.Cast(arguments[0]) : null; if (g) g.CancelEdit();",
            GRID_ID)
    except Exception:
        pass


def _read_value(driver, element_id):
    from selenium.webdriver.common.by import By
    try:
        return (driver.find_element(By.ID, element_id).get_attribute('value') or '').strip()
    except Exception:
        return ''


def _dismiss_alert(driver):
    """Αν υπάρχει ανοιχτό JS alert (π.χ. μήνυμα λάθους του MySchool), το
    κλείνει και επιστρέφει το κείμενό του — αλλιώς None."""
    try:
        alert = driver.switch_to.alert
        txt = alert.text
        alert.accept()
        return txt
    except Exception:
        return None


# ── Σύνδεση (καλείται από το UI) ─────────────────────────────────────────────

def _start_chrome(log):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService

    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1400,900')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    try:
        log('  Αυτόματη εύρεση/λήψη ChromeDriver...')
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            _wdm_path = ChromeDriverManager().install()
            driver = webdriver.Chrome(service=ChromeService(_wdm_path), options=options)
            log('  ChromeDriver OK')
        except Exception as _e:
            log(f'  webdriver-manager απέτυχε: {_e} — δοκιμάζω χωρίς service...')
            driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        log(f'Αδύνατη εκκίνηση Chrome: {e}')
        return None


def _start_firefox(log):
    """Ίδια λογική με το core/downloader.py: πρώτα Selenium built-in
    (χωρίς internet), μετά webdriver-manager."""
    from selenium import webdriver
    from selenium.webdriver.firefox.service import Service as FirefoxService
    from selenium.webdriver.firefox.options import Options as FirefoxOptions

    ff_options = FirefoxOptions()
    ff_options.add_argument('--width=1400')
    ff_options.add_argument('--height=900')

    log('  Εκκίνηση Firefox...')
    try:
        driver = webdriver.Firefox(options=ff_options)
        log('  GeckoDriver OK (Selenium built-in)')
        return driver
    except Exception as _builtin_err:
        log(f'  Selenium built-in απέτυχε: {_builtin_err} — δοκιμάζω webdriver-manager...')
    try:
        from webdriver_manager.firefox import GeckoDriverManager
        _gecko_path = GeckoDriverManager().install()
        driver = webdriver.Firefox(service=FirefoxService(_gecko_path), options=ff_options)
        log('  GeckoDriver OK (webdriver-manager)')
        return driver
    except Exception as e:
        log(f'Αδύνατη εκκίνηση Firefox: {e} — βεβαιώσου ότι ο Firefox είναι '
            f'εγκατεστημένος, ή άλλαξε σε Chrome από τις Ρυθμίσεις.')
        return None


def connect(log=print):
    """Άνοιγμα browser (Chrome ή Firefox, βάσει Ρυθμίσεις → Σύνδεση) + login."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import config as _cfg

    browser = (getattr(_cfg, 'BROWSER', 'chrome') or 'chrome').lower().strip()
    driver = _start_firefox(log) if browser == 'firefox' else _start_chrome(log)
    if driver is None:
        return None

    try:
        log('Σύνδεση στο MySchool...')
        driver.get(BASE_URL)
        time.sleep(2)

        if 'sso.sch.gr' in driver.current_url or 'login' in driver.current_url.lower():
            user_f = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    '#username, input[name="username"]')))
            user_f.clear()
            user_f.send_keys(_cfg.MYSCHOOL_USER)

            pass_f = driver.find_element(By.CSS_SELECTOR,
                '#password, input[name="password"], input[type="password"]')
            pass_f.clear()
            pass_f.send_keys(_cfg.MYSCHOOL_PASS)

            driver.find_element(By.CSS_SELECTOR,
                'button[type="submit"], input[type="submit"]').click()
            time.sleep(3)

        log('Φόρτωση σελίδας αναζήτησης...')
        driver.get(SEARCH_URL)
        time.sleep(3)
        log('Σύνδεση ΟΚ')
        return driver

    except Exception as e:
        log(f'Σφάλμα σύνδεσης: {e}')
        try:
            driver.quit()
        except Exception:
            pass
        return None


# ── Κύρια εκτέλεση ────────────────────────────────────────────────────────────

def run(ctx, driver, callback=None, ask_user=None):
    """
    Καταχώρηση απουσίας (Ολική Διάθεση / Απόσπαση) στην οργανική τοποθέτηση.

    Παράμετροι:
      ctx      : dict με 'file_path'
      driver   : Selenium WebDriver (ήδη συνδεδεμένος)
      callback : log(msg)
      ask_user : ask_user(title, prompt, options) → str|None
                 Χρησιμοποιείται σαν "ειδοποίηση διαφοροποίησης" — μία επιλογή
                 (Συνέχεια) ώστε ο χρήστης να δει το πρόβλημα και να προχωρήσει
                 στον επόμενο εκπαιδευτικό.
    """
    log  = callback or print
    _ask = ask_user or _fallback_ask

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    file_path = ctx.get('file_path')

    records, skipped_load = load_data(file_path, log=log)
    if skipped_load:
        log('\nΠαραλείφθηκαν κατά τη φόρτωση:')
        for s in skipped_load:
            log(f'   ⏭ {s}')

    if not records:
        log('Καμία εγγραφή προς επεξεργασία.')
        return

    # ── Αρχείο προόδου: φόρτωση + φιλτράρισμα ήδη-ολοκληρωμένων ──────────────
    status_path = _status_file_path(file_path)
    status = _load_status(status_path, log=log)
    log(f'Αρχείο προόδου: {status_path}')

    def _mark(am, name, status_val, comment):
        status[am] = {
            'name':    name,
            'status':  status_val,
            'comment': comment,
            'date':    time.strftime('%d/%m/%Y %H:%M'),
        }
        _save_status(status_path, status)

    already_done = []
    to_process   = []
    for rec in records:
        prev = status.get(rec['am'])
        if prev and prev['status'] in DONE_STATUSES:
            already_done.append(f"{rec['name']} ({rec['am']}) — {prev['status']}")
        else:
            to_process.append(rec)

    if already_done:
        log(f'\n⏭ {len(already_done)} ήδη ολοκληρωμένες από προηγούμενη εκτέλεση '
            f'(δεν ξαναγίνονται):')
        for s in already_done:
            log(f'   {s}')

    if not to_process:
        log('\nΌλες οι εγγραφές του αρχείου έχουν ήδη ολοκληρωθεί.')
        return

    total = len(to_process)
    log(f'\n{total} εγγραφές προς καταχώρηση σε αυτή την εκτέλεση.')

    ok = fail = 0
    ok_list = []
    fail_list = []
    already_live_list = []   # βρέθηκαν ήδη καταχωρημένες στο ίδιο το MySchool
    no_organic_list = []     # καμία γραμμή 1ης τριάδας ανάμεσα στα αποτελέσματα — αφήνονται στην άκρη

    def _notify(msg):
        """Ειδοποίηση διαφοροποίησης — ο χρήστης κλείνει και προχωράμε."""
        _ask('Διαφοροποίηση', msg, ['Συνέχεια στον επόμενο'])

    for idx, rec in enumerate(to_process, 1):
        am           = rec['am']
        name         = rec['name']
        apo          = rec['apo']
        eos          = rec['eos']
        absence_type = rec['absence_type']
        label = f'{name} ({am})' if name else am
        log(f'\n[{idx}/{total}] Α.Μ.: {am}  |  {name}  |  Τύπος: {absence_type}  |  '
            f'Ισχύει από: {apo}  |  Ισχύει έως: {eos}')

        # ── Σελίδα αναζήτησης ─────────────────────────────────────────────
        try:
            driver.get(SEARCH_URL)
            time.sleep(2)
        except Exception as e:
            log(f'  ✗ Πλοήγηση: {e}')
            _notify(f'{label}\nΣφάλμα πλοήγησης: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Πλοήγηση: {e}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Συμπλήρωση Α.Μ. ───────────────────────────────────────────────
        try:
            am_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located((By.ID, AM_FIELD_ID)))
            driver.execute_script("""
                var el  = arguments[0];
                var val = arguments[1];
                el.focus();
                el.value = '';
                for (var i = 0; i < val.length; i++) {
                    var ch = val[i];
                    el.value += ch;
                    el.dispatchEvent(new KeyboardEvent('keydown',  {key: ch, bubbles: true}));
                    el.dispatchEvent(new KeyboardEvent('keypress', {key: ch, bubbles: true}));
                    el.dispatchEvent(new KeyboardEvent('keyup',    {key: ch, bubbles: true}));
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                }
                el.dispatchEvent(new Event('change', {bubbles: true}));
            """, am_field, am)
            time.sleep(1)
        except Exception as e:
            log(f'  ✗ Α.Μ. field: {e}')
            _notify(f'{label}\nΔεν βρέθηκε το πεδίο Α.Μ.: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Πεδίο Α.Μ.: {e}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Κουμπί αναζήτησης ─────────────────────────────────────────────
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}')
            _notify(f'{label}\nΣφάλμα αναζήτησης: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Αναζήτηση: {e}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Εύρεση γραμμής με Σχέση τοποθέτησης από την 1η τριάδα ─────────
        edit_links = driver.find_elements(By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log('  ✗ Δεν βρέθηκε ο εκπαιδευτικός')
            _notify(f'{label}\nΔεν βρέθηκε καμία εγγραφή για το Α.Μ. {am}.')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', 'Δεν βρέθηκε ο εκπαιδευτικός')
            fail += 1
            fail_list.append(label)
            continue

        target = None
        for lnk in edit_links:
            try:
                row = lnk.find_element(By.XPATH, './ancestor::tr[1]')
                row_text = row.text.strip()
                if any(t in row_text for t in FIRST_TRIAD):
                    target = lnk
                    break
            except Exception:
                pass

        if not target:
            # Καμία από τις εγγραφές που επέστρεψε η αναζήτηση Α.Μ. δεν είναι
            # οργανική — αφήνεται στην άκρη ΧΩΡΙΣ popup (μπορεί να είναι πολλοί
            # τέτοιοι σε ένα αρχείο 70+ εγγραφών). Ξεχωριστή κατηγορία στη
            # σύνοψη· ΔΕΝ γράφεται DONE_STATUS ώστε να ξαναδοκιμαστεί σε
            # επόμενη εκτέλεση (μπορεί να προστεθεί οργανική στο μεταξύ).
            log('  ⏭ Δεν βρέθηκε οργανική τοποθέτηση ανάμεσα στα αποτελέσματα — αφήνεται στην άκρη')
            _mark(am, name, 'ΔΕΝ ΒΡΕΘΗΚΕ ΟΡΓΑΝΙΚΗ', 'Καμία γραμμή 1ης τριάδας στα αποτελέσματα αναζήτησης')
            no_organic_list.append(label)
            continue

        # ── Άνοιγμα καρτέλας ──────────────────────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Άνοιγμα καρτέλας: {e}')
            _notify(f'{label}\nΣφάλμα ανοίγματος καρτέλας: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Άνοιγμα καρτέλας: {e}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Έλεγχος διπλοεγγραφής: ίδιος τύπος ΚΑΙ ίδιες ημερομηνίες; ──────
        # Καλύπτει την περίπτωση προηγούμενης διακεκομμένης εκτέλεσης που
        # πρόλαβε να αποθηκεύσει στο MySchool αλλά όχι στο αρχείο προόδου.
        # Απουσία ίδιου τύπου με ΑΛΛΕΣ ημερομηνίες (π.χ. περσινή) ΔΕΝ
        # εμποδίζει την καταχώρηση.
        try:
            exists, others = _find_matching_absence(driver, absence_type, apo, eos)
            if exists:
                log(f'  ⏭ Η απουσία ({apo} – {eos}) υπάρχει ήδη καταχωρημένη — παράλειψη')
                _mark(am, name, 'ΥΠΑΡΧΕΙ ΗΔΗ',
                      f'Βρέθηκε ήδη στο grid ({apo} – {eos})')
                already_live_list.append(label)
                continue
            for o in others:
                log(f'  ℹ Υπάρχει απουσία ίδιου τύπου με άλλες ημερομηνίες (δεν '
                    f'θεωρείται διπλοεγγραφή): {o}')
        except Exception as e:
            log(f'  ⚠ Ο έλεγχος διπλοεγγραφής απέτυχε ({e}) — συνεχίζω με την προσθήκη')

        def _abort(reason, popup_msg=None):
            """Ακύρωση νέας γραμμής ΧΩΡΙΣ Αποθήκευση + καταγραφή ΑΠΟΤΥΧΙΑΣ."""
            nonlocal fail
            _cancel_new_row(driver)
            log(f'  ✗ {reason} — η γραμμή ακυρώθηκε, ΔΕΝ έγινε Αποθήκευση')
            _notify(f'{label}\n{popup_msg or reason}\n\nΔεν αποθηκεύτηκε τίποτα '
                    f'για αυτόν τον εκπαιδευτικό.')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', reason)
            fail += 1
            fail_list.append(label)

        # ── Πινακάκι Απουσιών (gridAbs): scroll + κλικ ➕ (Προσθήκη) ────────
        # Σημείωση: η καρτέλα έχει κι άλλο πινακάκι ("Λεπτομέρειες ωραρίου
        # εργασίας" / gridEmplDet) με δικό του κουμπί Προσθήκη πιο πάνω —
        # στοχεύουμε ρητά στο onclick="...gridAbs.AddNewRow()..." ώστε να
        # μην πατηθεί κατά λάθος το λάθος πινακάκι.
        try:
            add_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//*[contains(@onclick,"gridAbs.AddNewRow") or contains(@href,"gridAbs.AddNewRow")]')))
            driver.execute_script(
                'arguments[0].scrollIntoView({behavior:"smooth",block:"center"});',
                add_btn)
            time.sleep(1)
            driver.execute_script('arguments[0].click();', add_btn)
            time.sleep(2)
            log('  Φόρμα απουσίας ανοιχτή')
        except Exception as e:
            log(f'  ✗ Κουμπί Προσθήκη: {e}')
            _notify(f'{label}\nΔεν βρέθηκε/άνοιξε το κουμπί Προσθήκη απουσίας: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Κουμπί Προσθήκη: {e}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Εντοπισμός πεδίου τύπου απουσίας (δυναμικό index editnew_N) ───
        try:
            combo_input = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[id*="gridAbsences_editnew"][id*="cmbAbsenceType_I"]')))
            combo_full_id = combo_input.get_attribute('id')     # ..._cmbAbsenceType_I
            combo_base_id = combo_full_id[:-2]                  # χωρίς '_I'
        except Exception:
            _abort('Δεν βρέθηκε το πεδίο τύπου απουσίας μετά το Προσθήκη')
            continue

        # Τα πεδία ημερομηνίας (DXEditor5/6) ΔΕΝ έχουν το "editnew_N" prefix
        # του combo — τα IDs τους είναι σταθερά μέσα στο grid. Αν κάποτε
        # αλλάξουν, η επαλήθευση παρακάτω θα το πιάσει (ΑΠΟΤΥΧΙΑ, όχι λάθος
        # εγγραφή).
        from_field_id = GRID_ID + '_DXEditor5_I'
        to_field_id   = GRID_ID + '_DXEditor6_I'

        # ── Συμπλήρωση ────────────────────────────────────────────────────
        _select_dxe_combo(driver, combo_base_id, absence_type)
        try:
            _set_dxe_value(driver, from_field_id, apo)
        except Exception as e:
            log(f'  ⚠ Ισχύει από: {e}')
        try:
            _set_dxe_value(driver, to_field_id, eos)
        except Exception as e:
            log(f'  ⚠ Ισχύει έως: {e}')
        time.sleep(0.7)

        # ── Επαλήθευση ΟΛΩΝ των πεδίων πριν την Αποδοχή ──────────────────
        # Αν έστω και ένα δεν έχει τη σωστή τιμή → ακύρωση, καμία αποθήκευση.
        got_type = _read_value(driver, combo_full_id)
        got_apo  = _read_value(driver, from_field_id)
        got_eos  = _read_value(driver, to_field_id)

        problems = []
        if _normalize_text(got_type) != _normalize_text(absence_type):
            problems.append(f'Τύπος απουσίας: «{got_type or "κενό"}» αντί για «{absence_type}»')
        if _date_key(got_apo) != _date_key(apo):
            problems.append(f'Ισχύει από: «{got_apo or "κενό"}» αντί για «{apo}»')
        if _date_key(got_eos) != _date_key(eos):
            problems.append(f'Ισχύει έως: «{got_eos or "κενό"}» αντί για «{eos}»')

        if problems:
            for pr in problems:
                log(f'  ✗ {pr}')
            _abort('Λάθος/κενό πεδίο: ' + '; '.join(problems),
                   'Τα πεδία δεν συμπληρώθηκαν σωστά:\n• ' + '\n• '.join(problems))
            continue

        log(f'  ✓ Τύπος απουσίας: {got_type}')
        log(f'  ✓ Ισχύει από: {got_apo}')
        log(f'  ✓ Ισχύει έως: {got_eos}')

        # ── Αποδοχή (πράσινο τικ) — σκοπευμένο στο gridAbsences ───────────
        try:
            accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//img[@alt="Αποδοχή" and contains(@onclick,"gridAbsences")]')))
            driver.execute_script('arguments[0].click();', accept_btn)
            time.sleep(2)
        except Exception as e:
            _abort(f'Αποδοχή: {e}', f'Σφάλμα κατά την Αποδοχή: {e}')
            continue

        alert_txt = _dismiss_alert(driver)
        if alert_txt:
            _abort(f'Μήνυμα MySchool στην Αποδοχή: {alert_txt}')
            continue

        # Αν η γραμμή είναι ακόμη σε επεξεργασία, η Αποδοχή δεν «έπιασε»
        # (π.χ. πρόβλημα εγκυρότητας) — 2η προσπάθεια μέσω client API.
        if _combo_visible(driver, combo_full_id):
            log('  ⚠ Η γραμμή παραμένει σε επεξεργασία — δοκιμή μέσω client API...')
            try:
                driver.execute_script(
                    "var g = ASPxClientGridView.Cast(arguments[0]); if (g) g.UpdateEdit();",
                    GRID_ID)
            except Exception as e:
                log(f'  ⚠ UpdateEdit: {e}')
            time.sleep(2)
            alert_txt = _dismiss_alert(driver)
            if alert_txt:
                _abort(f'Μήνυμα MySchool στην Αποδοχή: {alert_txt}')
                continue
        if _combo_visible(driver, combo_full_id):
            _abort('Η γραμμή ΔΕΝ έγινε αποδεκτή (παραμένει σε επεξεργασία — '
                   'πιθανό πρόβλημα εγκυρότητας ή επικάλυψη ημερομηνιών)')
            continue
        log('  ✓ Αποδοχή (η γραμμή έκλεισε κανονικά)')

        # ── Scroll στην κορυφή + Αποθήκευση ─────────────────────────────
        driver.execute_script('window.scrollTo(0, 0);')
        time.sleep(0.5)
        try:
            save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.ID, 'ctl00_ContentData_btnSave')))
            driver.execute_script('arguments[0].click();', save_btn)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}')
            _notify(f'{label}\nΣφάλμα κατά την Αποθήκευση: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Αποθήκευση: {e}')
            fail += 1
            fail_list.append(label)
            continue

        alert_txt = _dismiss_alert(driver)
        if alert_txt:
            log(f'  ✗ Μήνυμα MySchool μετά την Αποθήκευση: {alert_txt}')
            _notify(f'{label}\nΤο MySchool εμφάνισε μήνυμα μετά την Αποθήκευση:\n'
                    f'{alert_txt}\n\nΈλεγξε χειροκίνητα την καρτέλα.')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Μήνυμα μετά την Αποθήκευση: {alert_txt}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Επαλήθευση ΜΕΤΑ την αποθήκευση (ground truth) ────────────────
        # Η νέα απουσία (τύπος + από + έως) πρέπει να φαίνεται στο grid. Αν
        # όχι, ΑΠΟΤΥΧΙΑ → θα ξαναδοκιμαστεί στην επόμενη εκτέλεση· αν τελικά
        # είχε σωθεί, ο έλεγχος διπλοεγγραφής θα τη βρει και δεν θα τη
        # ξαναβάλει.
        try:
            saved, _ = _find_matching_absence(driver, absence_type, apo, eos)
        except Exception:
            saved = False
        if not saved:
            # 2η ματιά: ξαναφόρτωση της καρτέλας με GET (ΟΧΙ refresh(), που
            # μετά από postback μπορεί να ξαναστείλει την Αποθήκευση).
            time.sleep(2)
            try:
                saved, _ = _find_matching_absence(driver, absence_type, apo, eos)
            except Exception:
                saved = False
        if not saved:
            try:
                driver.get(driver.current_url)
                time.sleep(3)
                _dismiss_alert(driver)
                saved, _ = _find_matching_absence(driver, absence_type, apo, eos)
            except Exception:
                saved = False

        if saved:
            log('  ✓ Αποθήκευση — επιβεβαιώθηκε στον πίνακα Απουσιών')
            ok += 1
            ok_list.append(label)
            _mark(am, name, 'ΕΠΙΤΥΧΙΑ', f'Καταχωρήθηκε ({apo} – {eos})')
        else:
            log('  ✗ ΔΕΝ επιβεβαιώθηκε ότι η απουσία εμφανίζεται στον πίνακα μετά '
                'την Αποθήκευση — έλεγξε χειροκίνητα')
            _notify(f'{label}\nΠατήθηκε Αποθήκευση αλλά η νέα απουσία ΔΕΝ φαίνεται '
                    f'στον πίνακα Απουσιών.\nΈλεγξε χειροκίνητα την καρτέλα.')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', 'Μη επιβεβαιωμένη αποθήκευση — έλεγχος χειροκίνητα')
            fail += 1
            fail_list.append(label)

    # ── Σύνοψη ────────────────────────────────────────────────────────────
    log(f'\n{"═" * 55}')
    log(f'  ΑΠΟΤΕΛΕΣΜΑΤΑ ΚΑΤΑΧΩΡΗΣΗΣ ΑΠΟΥΣΙΩΝ')
    log(f'{"─" * 55}')
    log(f'  ✓  Καταχωρήθηκαν τώρα        : {ok}')
    log(f'  ⏭ Υπήρχαν ήδη (live έλεγχος) : {len(already_live_list)}')
    log(f'  ⏭ Ήδη ολοκληρωμένες (προηγ. εκτέλεση) : {len(already_done)}')
    log(f'  ⏭ Δεν βρέθηκε οργανική (αφέθηκαν στην άκρη) : {len(no_organic_list)}')
    log(f'  ✗  Αποτυχίες                 : {fail}')
    log(f'  ⏭ Παραλείφθηκαν (φιλτράρισμα αρχείου) : {len(skipped_load)}')
    log(f'{"═" * 55}')

    if ok_list:
        log(f'\n✓ ΚΑΤΑΧΩΡΗΘΗΚΑΝ ΤΩΡΑ ({len(ok_list)}):')
        for e in ok_list:
            log(f'   {e}')

    if already_live_list:
        log(f'\n⏭ ΥΠΗΡΧΑΝ ΗΔΗ ΣΤΟ MYSCHOOL ({len(already_live_list)}):')
        for e in already_live_list:
            log(f'   {e}')

    if no_organic_list:
        log(f'\n⏭ ΔΕΝ ΒΡΕΘΗΚΕ ΟΡΓΑΝΙΚΗ — ΑΦΕΘΗΚΑΝ ΣΤΗΝ ΑΚΡΗ ({len(no_organic_list)}):')
        for e in no_organic_list:
            log(f'   {e}')

    if fail_list:
        log(f'\n✗ ΑΠΟΤΥΧΙΕΣ ({len(fail_list)}) — θα ξαναδοκιμαστούν στην επόμενη εκτέλεση:')
        for e in fail_list:
            log(f'   {e}')

    log(f'\nΤο αρχείο προόδου ({status_path}) ενημερώνεται μετά από κάθε '
        f'εκπαιδευτικό — αν ξανατρέξεις με το ίδιο αρχείο, δεν θα ξαναγίνουν '
        f'όσοι έχουν ήδη ΕΠΙΤΥΧΙΑ ή ΥΠΑΡΧΕΙ ΗΔΗ.')


# ── Fallback ask (για terminal / testing) ─────────────────────────────────────

def _fallback_ask(title, prompt, options=None):
    print(f'\n{"=" * 50}')
    print(f'[{title}]')
    print(prompt)
    if options:
        for i, opt in enumerate(options, 1):
            print(f'  {i}. {opt}')
        input('Πάτησε Enter για συνέχεια: ')
        return '1'
    else:
        return input('Εισαγωγή: ').strip() or None
