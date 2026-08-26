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
       Ισχύει από: 1/9/2026 (σταθερό για όλους)
       Ισχύει έως: από τη στήλη "Έως" της εγγραφής του αρχείου
  5. Κλικ ✓ (Αποδοχή) → Αποθήκευση
  6. Επόμενος

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
    Απουσιών στο MySchool — αν βρει ήδη καταχωρημένη την ίδια απουσία (π.χ.
    από προηγούμενη εκτέλεση που δεν πρόλαβε να γραφτεί στο αρχείο προόδου),
    δεν την ξαναπροσθέτει, το καταγράφει ως "ΥΠΑΡΧΕΙ ΗΔΗ" και προχωράει.
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

FIXED_FROM_DATE = '1/9/2026'


# ── Ανάγνωση & φιλτράρισμα αρχείου ────────────────────────────────────────────

def _fmt_date(ts):
    """Μετατρέπει pandas Timestamp σε 'D/M/YYYY' (χωρίς μηδενικά πρόθεμα)."""
    if pd.isna(ts):
        return ''
    try:
        return f'{ts.day}/{ts.month}/{ts.year}'
    except Exception:
        return str(ts)


def load_data(file_path, log=print):
    """
    Διαβάζει το αρχείο εξαγωγής τοποθετήσεων και επιστρέφει λίστα εγγραφών
    προς επεξεργασία: μόνο εκπαιδευτικοί με Σχέση τοποθέτησης σε
    SECOND_TRIAD_ABSENCE_TYPE (Ολική Διάθεση ή Απόσπαση). Το αρχείο ΔΕΝ
    χρειάζεται να περιέχει γραμμή «Οργανικά» — αυτή εντοπίζεται αργότερα
    απευθείας στο MySchool (βλ. run()).

    Επιστρέφει: [{'am', 'afm', 'name', 'eos', 'absence_type'}], skipped: [str, ...]
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
    required = ['Α.Μ.', 'Σχέση τοποθέτησης', 'Έως']
    missing = [c for c in required if c not in cols]
    if missing:
        log(f'Λείπουν στήλες: {missing}. Διαθέσιμες: {list(cols.keys())}')
        return [], []

    am_col    = cols['Α.Μ.']
    rel_col   = cols['Σχέση τοποθέτησης']
    eos_col   = cols['Έως']
    afm_col   = cols.get('Α.Φ.Μ.')
    name_col  = cols.get('Επώνυμο') or cols.get('ΕΠΙΘΕΤΟ') or cols.get('Επίθετο')
    fname_col = cols.get('Όνομα')   or cols.get('ΟΝΟΜΑ')

    records = []
    skipped = []

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

        eos = _fmt_date(second_row.iloc[0][eos_col])
        if not eos:
            skipped.append(f'{full_name} ({am_str}) — κενή ημ. Έως')
            continue

        afm = str(grp.iloc[0].get(afm_col, '')).strip() if afm_col else ''

        records.append({
            'am':           am_str,
            'afm':          afm,
            'name':         full_name,
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


def _absence_already_exists(driver, type_text):
    """
    Ελέγχει αν το πινακάκι Απουσιών (gridAbsences) έχει ΗΔΗ γραμμή με το
    δοσμένο λεκτικό τύπου απουσίας — π.χ. από προηγούμενη (διακεκομμένη)
    εκτέλεση που δεν πρόλαβε να καταγράψει την επιτυχία στο αρχείο προόδου.
    Στοχεύουμε ρητά στο container του grid (όχι όλη τη σελίδα) γιατί το ίδιο
    λεκτικό εμφανίζεται και μέσα στη λίστα του dropdown.
    """
    from selenium.webdriver.common.by import By
    try:
        container = driver.find_element(By.ID, 'ctl00_ContentData_gridAbsences')
        if type_text in container.text:
            return True
    except Exception:
        pass
    try:
        containers = driver.find_elements(
            By.XPATH, '//*[contains(@id,"gridAbsences") and not(contains(@id,"editnew"))]')
        for c in containers:
            if type_text in c.text:
                return True
    except Exception:
        pass
    return False


# ── Σύνδεση (καλείται από το UI) ─────────────────────────────────────────────

def connect(log=print):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import config as _cfg

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
    except Exception as e:
        log(f'Αδύνατη εκκίνηση Chrome: {e}')
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
    Καταχώρηση απουσίας Ολικής Διάθεσης στην οργανική τοποθέτηση.

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
        eos          = rec['eos']
        absence_type = rec['absence_type']
        label = f'{name} ({am})' if name else am
        log(f'\n[{idx}/{total}] Α.Μ.: {am}  |  {name}  |  Τύπος: {absence_type}  |  '
            f'Ισχύει έως: {eos}')

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

        # ── Έλεγχος διπλοεγγραφής: υπάρχει ήδη η απουσία στο grid; ────────
        # Καλύπτει την περίπτωση προηγούμενης διακεκομμένης εκτέλεσης που
        # πρόλαβε να αποθηκεύσει στο MySchool αλλά όχι στο αρχείο προόδου.
        try:
            if _absence_already_exists(driver, absence_type):
                log('  ⏭ Η απουσία υπάρχει ήδη καταχωρημένη — παράλειψη')
                _mark(am, name, 'ΥΠΑΡΧΕΙ ΗΔΗ', 'Βρέθηκε ήδη στο grid κατά τον έλεγχο')
                already_live_list.append(label)
                continue
        except Exception:
            pass  # αν αποτύχει ο έλεγχος, συνεχίζουμε κανονικά με την προσθήκη

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
        except Exception as e:
            log(f'  ✗ Δεν βρέθηκε το πεδίο τύπου απουσίας: {e}')
            _notify(f'{label}\nΔεν βρέθηκε το πεδίο τύπου απουσίας μετά το Προσθήκη.')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', 'Δεν βρέθηκε το πεδίο τύπου απουσίας')
            fail += 1
            fail_list.append(label)
            continue

        # Τα πεδία ημερομηνίας (DXEditor5/6) ΔΕΝ έχουν το "editnew_N" prefix
        # του combo — τα IDs τους είναι σταθερά μέσα στο grid.
        from_field_id = 'ctl00_ContentData_gridAbsences_DXEditor5_I'
        to_field_id   = 'ctl00_ContentData_gridAbsences_DXEditor6_I'

        # ── Τύπος απουσίας ─────────────────────────────────────────────────
        ok_c = _select_dxe_combo(driver, combo_base_id, absence_type)
        log(f'  {"✓" if ok_c else "⚠"} Τύπος απουσίας: {absence_type}')
        if not ok_c:
            _notify(f'{label}\nΑποτυχία επιλογής τύπου απουσίας από τη λίστα.')

        # ── Ισχύει από ────────────────────────────────────────────────────
        try:
            set_ok = _set_dxe_value(driver, from_field_id, FIXED_FROM_DATE)
            if set_ok:
                log(f'  ✓ Ισχύει από: {FIXED_FROM_DATE}')
            else:
                log(f'  ⚠ Ισχύει από: πεδίο δεν βρέθηκε ({from_field_id})')
                _notify(f'{label}\nΔεν βρέθηκε το πεδίο "Ισχύει από" ({from_field_id}).')
        except Exception as e:
            log(f'  ⚠ Ισχύει από: {e}')
            _notify(f'{label}\nΑποτυχία συμπλήρωσης "Ισχύει από": {e}')

        # ── Ισχύει έως ────────────────────────────────────────────────────
        try:
            set_ok = _set_dxe_value(driver, to_field_id, eos)
            if set_ok:
                log(f'  ✓ Ισχύει έως: {eos}')
            else:
                log(f'  ⚠ Ισχύει έως: πεδίο δεν βρέθηκε ({to_field_id})')
                _notify(f'{label}\nΔεν βρέθηκε το πεδίο "Ισχύει έως" ({to_field_id}).')
        except Exception as e:
            log(f'  ⚠ Ισχύει έως: {e}')
            _notify(f'{label}\nΑποτυχία συμπλήρωσης "Ισχύει έως": {e}')

        time.sleep(0.5)

        # ── Αποδοχή (πράσινο τικ) — σκοπευμένο στο gridAbsences ───────────
        try:
            accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//img[@alt="Αποδοχή" and contains(@onclick,"gridAbsences")]')))
            driver.execute_script('arguments[0].click();', accept_btn)
            time.sleep(2)
            log('  ✓ Αποδοχή')
        except Exception as e:
            log(f'  ✗ Αποδοχή: {e}')
            _notify(f'{label}\nΣφάλμα κατά την Αποδοχή: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Αποδοχή: {e}')
            fail += 1
            fail_list.append(label)
            continue

        # ── Scroll στην κορυφή + Αποθήκευση ─────────────────────────────
        driver.execute_script('window.scrollTo(0, 0);')
        time.sleep(0.5)
        try:
            save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.ID, 'ctl00_ContentData_btnSave')))
            driver.execute_script('arguments[0].click();', save_btn)
            time.sleep(3)
            log('  ✓ Αποθήκευση')
            ok += 1
            ok_list.append(label)
            _mark(am, name, 'ΕΠΙΤΥΧΙΑ', 'Καταχωρήθηκε')
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}')
            _notify(f'{label}\nΣφάλμα κατά την Αποθήκευση: {e}')
            _mark(am, name, 'ΑΠΟΤΥΧΙΑ', f'Αποθήκευση: {e}')
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
