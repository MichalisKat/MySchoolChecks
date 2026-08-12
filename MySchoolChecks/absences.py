#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
absences.py
============
Αυτόματη καταχώρηση απουσίας (Ολική Διάθεση) σε οργανική τοποθέτηση στο MySchool.

Πλαίσιο:
  Κάθε εκπαιδευτικός έχει 2 εγγραφές τοποθέτησης στο αρχείο εξαγωγής (Στήλη
  "Σχέση τοποθέτησης"):
    1η τριάδα (εκεί μπαίνει η απουσία): Οργανικά / Οργανικά σε Τμήμα Ένταξης /
                                          Οργανικά από Αρση Υπεραριθμίας
    2η τριάδα (πηγή στοιχείων απουσίας): Ολική Διάθεση (ανάγκες υπηρεσίας -
                                          κύριος φορέας) / Απόσπαση (με αίτηση -
                                          κύριος φορέας) / Επί Θητεία

  Επεξεργάζονται ΜΟΝΟ οι εκπαιδευτικοί με 2η εγγραφή = "Ολική Διάθεση
  (ανάγκες υπηρεσίας - κύριος φορέας)". Οι υπόλοιποι (Απόσπαση/Επί Θητεία)
  εξαιρούνται (χειροκίνητη καταχώρηση).

Ροή ανά εκπαιδευτικό:
  1. Αναζήτηση με Α.Μ. στο Worker.list.myEmplUnit.aspx
  2. Εντοπισμός της γραμμής με Σχέση τοποθέτησης από την 1η τριάδα → κλικ γρανάζι
  3. Scroll στο πινακάκι Απουσιών (gridAbsences) → κλικ ➕ (Προσθήκη)
  4. Συμπλήρωση:
       Τύπος απουσίας: ΟΛΙΚΗ ΔΙΑΘΕΣΗ ΣΕ ΑΛΛΗ ΣΧΟΛΙΚΗ ΜΟΝΑΔΑ - Σχολικές
                        Μονάδες Πρωτοβάθμιας
       Ισχύει από: 1/9/2026 (σταθερό για όλους)
       Ισχύει έως: από τη στήλη "Έως" της εγγραφής Ολικής Διάθεσης
  5. Κλικ ✓ (Αποδοχή) → Αποθήκευση
  6. Επόμενος

  Για οτιδήποτε διαφοροποιείται από το αναμενόμενο (δεν βρέθηκε η σωστή
  γραμμή, δεν βρέθηκε καθόλου ο εκπαιδευτικός, dropdown/πεδίο απέτυχε κ.λπ.)
  εμφανίζεται μήνυμα στον χρήστη — ο χρήστης το κλείνει και συνεχίζει στον
  επόμενο εκπαιδευτικό.
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Καταχώρηση Απουσίας σε Οργανική'
CHECK_DESCRIPTION = 'Αυτόματη καταχώρηση απουσίας (Ολική Διάθεση) στο MySchool'
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

# 2η τριάδα — μόνο αυτή η τιμή γίνεται δεκτή (οι άλλες δύο εξαιρούνται)
SECOND_TRIAD_TARGET = 'Ολική Διάθεση (ανάγκες υπηρεσίας - κύριος φορέας)'

# Λεκτικό τύπου απουσίας στο dropdown MySchool
ABSENCE_TYPE_TEXT = ('ΟΛΙΚΗ ΔΙΑΘΕΣΗ ΣΕ ΑΛΛΗ ΣΧΟΛΙΚΗ ΜΟΝΑΔΑ - '
                      'Σχολικές Μονάδες Πρωτοβάθμιας')

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
    προς επεξεργασία: μόνο εκπαιδευτικοί με 2η εγγραφή = SECOND_TRIAD_TARGET.

    Επιστρέφει: [{'am', 'afm', 'name', 'eos'}], skipped: [str, ...]
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
    name_col  = cols.get('Επώνυμο')
    fname_col = cols.get('Όνομα')

    records = []
    skipped = []

    for am, grp in df.groupby(am_col):
        first_row  = grp[grp[rel_col].isin(FIRST_TRIAD)]
        second_row = grp[grp[rel_col] == SECOND_TRIAD_TARGET]

        epon = str(grp.iloc[0].get(name_col, '')).strip()  if name_col  else ''
        onom = str(grp.iloc[0].get(fname_col, '')).strip() if fname_col else ''
        full_name = f'{epon} {onom}'.strip()
        am_str = str(am).strip()

        if first_row.empty or second_row.empty:
            skipped.append(f'{full_name} ({am_str}) — δεν βρέθηκε το κατάλληλο ζεύγος εγγραφών')
            continue

        eos = _fmt_date(second_row.iloc[0][eos_col])
        if not eos:
            skipped.append(f'{full_name} ({am_str}) — κενή ημ. Έως στην Ολική Διάθεση')
            continue

        afm = str(grp.iloc[0].get(afm_col, '')).strip() if afm_col else ''

        records.append({
            'am':   am_str,
            'afm':  afm,
            'name': full_name,
            'eos':  eos,
        })

    log(f'Φορτώθηκαν {len(records)} εγγραφές προς καταχώρηση '
        f'({len(skipped)} παραλείφθηκαν στο φιλτράρισμα)')
    return records, skipped


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

    total = len(records)
    log(f'\n{total} εγγραφές προς καταχώρηση.')

    ok = fail = 0
    ok_list = []
    fail_list = []

    def _notify(msg):
        """Ειδοποίηση διαφοροποίησης — ο χρήστης κλείνει και προχωράμε."""
        _ask('Διαφοροποίηση', msg, ['Συνέχεια στον επόμενο'])

    for idx, rec in enumerate(records, 1):
        am    = rec['am']
        name  = rec['name']
        eos   = rec['eos']
        label = f'{name} ({am})' if name else am
        log(f'\n[{idx}/{total}] Α.Μ.: {am}  |  {name}  |  Ισχύει έως: {eos}')

        # ── Σελίδα αναζήτησης ─────────────────────────────────────────────
        try:
            driver.get(SEARCH_URL)
            time.sleep(2)
        except Exception as e:
            log(f'  ✗ Πλοήγηση: {e}')
            _notify(f'{label}\nΣφάλμα πλοήγησης: {e}')
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
            fail += 1
            fail_list.append(label)
            continue

        # ── Εύρεση γραμμής με Σχέση τοποθέτησης από την 1η τριάδα ─────────
        edit_links = driver.find_elements(By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log('  ✗ Δεν βρέθηκε ο εκπαιδευτικός')
            _notify(f'{label}\nΔεν βρέθηκε καμία εγγραφή για το Α.Μ. {am}.')
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
            log('  ✗ Δεν βρέθηκε γραμμή 1ης τριάδας (Οργανικά κ.λπ.)')
            _notify(f'{label}\nΔεν βρέθηκε γραμμή με Σχέση τοποθέτησης από την '
                     f'1η τριάδα (Οργανικά / Οργανικά σε Τμήμα Ένταξης / '
                     f'Οργανικά από Αρση Υπεραριθμίας).')
            fail += 1
            fail_list.append(label)
            continue

        # ── Άνοιγμα καρτέλας ──────────────────────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Άνοιγμα καρτέλας: {e}')
            _notify(f'{label}\nΣφάλμα ανοίγματος καρτέλας: {e}')
            fail += 1
            fail_list.append(label)
            continue

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
            fail += 1
            fail_list.append(label)
            continue

        # Τα πεδία ημερομηνίας (DXEditor5/6) ΔΕΝ έχουν το "editnew_N" prefix
        # του combo — τα IDs τους είναι σταθερά μέσα στο grid.
        from_field_id = 'ctl00_ContentData_gridAbsences_DXEditor5_I'
        to_field_id   = 'ctl00_ContentData_gridAbsences_DXEditor6_I'

        # ── Τύπος απουσίας ─────────────────────────────────────────────────
        ok_c = _select_dxe_combo(driver, combo_base_id, ABSENCE_TYPE_TEXT)
        log(f'  {"✓" if ok_c else "⚠"} Τύπος απουσίας')
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
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}')
            _notify(f'{label}\nΣφάλμα κατά την Αποθήκευση: {e}')
            fail += 1
            fail_list.append(label)

    # ── Σύνοψη ────────────────────────────────────────────────────────────
    log(f'\n{"═" * 55}')
    log(f'  ΑΠΟΤΕΛΕΣΜΑΤΑ ΚΑΤΑΧΩΡΗΣΗΣ ΑΠΟΥΣΙΩΝ')
    log(f'{"─" * 55}')
    log(f'  ✓  Καταχωρήθηκαν  : {ok}')
    log(f'  ✗  Αποτυχίες      : {fail}')
    log(f'  ⏭ Παραλείφθηκαν (φιλτράρισμα αρχείου) : {len(skipped_load)}')
    log(f'{"═" * 55}')

    if fail_list:
        log('\n✗ ΑΠΟΤΥΧΙΕΣ:')
        for e in fail_list:
            log(f'   {e}')


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
