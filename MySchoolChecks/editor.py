#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
editor.py
=========
Αυτόματη συμπλήρωση Γραμματειακής Υποστήριξης στο MySchool.

Ροή:
1. Επιλογή Excel/CSV (στήλες: Α.Φ.Μ., Ονομασία Σχολείου προαιρετική)
2. Σύνδεση στο MySchool (κουμπί Σύνδεση στο UI)
3. Για κάθε ΑΦΜ: αναζήτηση → ανάγνωση ωρών → άνοιγμα καρτέλας
4. Επιλογή "Γραμματειακή Υποστήριξη" + ώρες + ημερομηνίες (= σήμερα)
5. Αποδοχή → Αποθήκευση
"""

import os
import sys
import time
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Editor — Επεξεργασία Καρτέλας Εκπαιδευτικού'
CHECK_DESCRIPTION = 'Αυτόματη συμπλήρωση Γραμματειακής Υποστήριξης στο MySchool'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL   = 'https://app.myschool.sch.gr'
SEARCH_URL = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

WORK_TYPE_TEXT = 'Γραμματειακή Υποστήριξη'


# ── Ανάγνωση αρχείου ─────────────────────────────────────────────────────────

def load_data(file_path, log=print):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(file_path, dtype=str)
        else:
            for enc in ('utf-8', 'iso-8859-7', 'cp1253'):
                try:
                    df = pd.read_csv(file_path, sep=';', encoding=enc, dtype=str)
                    break
                except Exception:
                    pass
            else:
                df = pd.read_csv(file_path, dtype=str)
    except Exception as e:
        log(f'Σφάλμα ανάγνωσης αρχείου: {e}')
        return []

    cols = {c.strip(): c for c in df.columns}

    afm_col = None
    for candidate in ('Α.Φ.Μ.', 'ΑΦΜ', 'Α.Φ.Μ', 'AFM'):
        if candidate in cols:
            afm_col = cols[candidate]
            break

    school_col = None
    for candidate in ('Ονομασία Σχολείου', 'ΣΧΟΛΕΙΟ', 'Σχολείο', 'ΚΩΔ. ΣΧΟΛΕΙΟΥ'):
        if candidate in cols:
            school_col = cols[candidate]
            break

    if not afm_col:
        log(f'Δεν βρέθηκε στήλη ΑΦΜ. Διαθέσιμες: {list(cols.keys())}')
        return []

    records = []
    for _, row in df.iterrows():
        afm    = str(row.get(afm_col, '')).strip()
        school = str(row.get(school_col, '')).strip() if school_col else ''
        if afm and afm not in ('nan', 'None', ''):
            records.append({'afm': afm.zfill(9), 'school': school})

    log(f'Φορτώθηκαν {len(records)} εγγραφές')
    return records


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
    """Επιλέγει τιμή από DevExpress ComboBox με πληκτρολόγηση + TAB."""
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
        # Πληκτρολόγηση αργά για να φιλτράρει η λίστα
        _send_keys_slow(inp, text)
        time.sleep(1.5)
        # Αναμονή για item με dxtext και κλικ
        try:
            item = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, f'td[dxtext="{text}"]')))
            driver.execute_script('arguments[0].click();', item)
            time.sleep(0.5)
            return True
        except Exception:
            pass
        # Fallback: TAB για επιλογή πρώτου αποτελέσματος
        inp.send_keys(Keys.TAB)
        time.sleep(0.5)
        return True
    except Exception as e:
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


# ── Εκτέλεση (καλείται από το UI) ────────────────────────────────────────────

def run(ctx, driver, callback=None):
    log = callback or print

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    file_path = ctx.get('file_path')
    records   = load_data(file_path, log=log)
    if not records:
        return

    today_str = date.today().strftime('%d/%m/%Y')
    total     = len(records)
    log(f'  {total} εγγραφές  |  Ημερομηνία: {today_str}')

    ok = fail = 0

    for idx, record in enumerate(records, 1):
        afm    = record['afm']
        school = record['school']
        log(f'\n[{idx}/{total}] ΑΦΜ: {afm}  |  {school}')

        # ── Σελίδα αναζήτησης ─────────────────────────────────────────────
        driver.get(SEARCH_URL)
        time.sleep(2)

        # ── Συμπλήρωση ΑΦΜ ────────────────────────────────────────────────
        try:
            afm_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.ID, 'ctl00_ContentData_txtTaxNumber_I')))
            # JS dispatch KeyboardEvent -- μοναδικος τροπος για DevExpress
            driver.execute_script("""
                var el = arguments[0];
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
            """, afm_field, afm)
            time.sleep(1)
        except Exception as e:
            log(f'  ✗ ΑΦΜ field: {e}')
            fail += 1
            continue

        # ── Αναζήτηση ─────────────────────────────────────────────────────
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}')
            fail += 1
            continue

        hours = ''

        # ── Εύρεση εγγραφής ───────────────────────────────────────────────
        edit_links = driver.find_elements(
            By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log('  ✗ Κανένα αποτέλεσμα')
            fail += 1
            continue

        target = None
        if len(edit_links) == 1:
            target = edit_links[0]
        else:
            for link in edit_links:
                try:
                    row = link.find_element(By.XPATH, './ancestor::tr[1]')
                    if school.upper() in row.text.upper():
                        target = link
                        log(f'  Βρέθηκε: {school}')
                        break
                except Exception:
                    pass
            if not target:
                log('  Χρήση πρώτης εγγραφής')
                target = edit_links[0]

        # ── Άνοιγμα καρτέλας ──────────────────────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Καρτέλα: {e}')
            fail += 1
            continue

        # ── Ώρες από καρτέλα (μετά άνοιγμα) ────────────────────────────
        try:
            hours_el = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.ID, 'ctl00_ContentData_txtAvailableHoursForUnit_I')))
            hours = hours_el.get_attribute('value').strip()
            log(f'  Διαθέσιμες ώρες: {hours}')
        except Exception as e:
            log(f'  ⚠ Ώρες: {e}')

        # ── Σταυρός Προσθήκης ─────────────────────────────────────────────
        try:
            add_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.ID, 'ctl00_ContentData_gridEmplDet_header0_new')))
            driver.execute_script(
                'arguments[0].scrollIntoView({behavior:"smooth",block:"center"});',
                add_btn)
            time.sleep(1)
            driver.execute_script('arguments[0].click();', add_btn)
            time.sleep(2)
            log('  Φόρμα ωραρίου ανοιχτή')
        except Exception as e:
            log(f'  ✗ Σταυρός: {e}')
            fail += 1
            continue

        # ── Dropdown "Γραμματειακή Υποστήριξη" ───────────────────────────
        combo_base = ('ctl00_ContentData_gridEmplDet_'
                      'editnew_2_cmbWorkHoursDetailsType')
        try:
            ok_c = _select_dxe_combo(driver, combo_base, WORK_TYPE_TEXT)
            log(f'  {"✓" if ok_c else "⚠"} Τύπος: {WORK_TYPE_TEXT}')
        except Exception as e:
            log(f'  ⚠ Dropdown: {e}')

        # ── Ώρες φόρμα (DXEditor4) — κλικ κλείνει και το dropdown ─────────
        if hours:
            try:
                hours_inp = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.ID, 'ctl00_ContentData_gridEmplDet_DXEditor4_I')))
                driver.execute_script('arguments[0].click(); arguments[0].focus();', hours_inp)
                time.sleep(0.5)
                driver.execute_script("""
                    var el = arguments[0];
                    var val = arguments[1];
                    el.value = '';
                    el.value = val;
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                    aspxEValueChanged('ctl00_ContentData_gridEmplDet_DXEditor4');
                """, hours_inp, hours)
                log(f'  ✓ Ώρες: {hours}')
            except Exception as e:
                log(f'  ⚠ Ώρες φόρμα: {e}')

        # ── Ημ. από (DXEditor5) ───────────────────────────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor5_I', today_str)
            log(f'  ✓ Ημ. από: {today_str}')
        except Exception as e:
            log(f'  ⚠ Ημ. από: {e}')

        # ── Ημ. έως (DXEditor6) ───────────────────────────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor6_I', today_str)
            log(f'  ✓ Ημ. έως: {today_str}')
        except Exception as e:
            log(f'  ⚠ Ημ. έως: {e}')

        time.sleep(0.5)

        # ── Αποδοχή (πράσινο τικ) ────────────────────────────────────────
        try:
            accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//img[@alt="Αποδοχή"]')))
            driver.execute_script('arguments[0].click();', accept_btn)
            time.sleep(2)
            log('  ✓ Αποδοχή')
        except Exception as e:
            log(f'  ✗ Αποδοχή: {e}')
            fail += 1
            continue

        # ── Αποθήκευση ────────────────────────────────────────────────────
        try:
            save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable(
                    (By.ID, 'ctl00_ContentData_btnSave')))
            driver.execute_script('arguments[0].click();', save_btn)
            time.sleep(3)
            log('  ✓ Αποθήκευση')
            ok += 1
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}')
            fail += 1

    log(f'\n{"─"*50}')
    log(f'Ολοκλήρωση: {ok} επιτυχείς, {fail} αποτυχίες')

    # Αποθήκευση path για χρήση από τη Λήξη
    if ok > 0:
        _save_panic_path(file_path)


def _panic_path_file():
    """Επιστρέφει το path του αρχείου που κρατάει το τελευταίο panic Excel."""
    docs = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks')
    os.makedirs(docs, exist_ok=True)
    return os.path.join(docs, '.panic_last_file.txt')


def _save_panic_path(file_path):
    try:
        with open(_panic_path_file(), 'w', encoding='utf-8') as f:
            f.write(file_path)
    except Exception:
        pass


def get_panic_path():
    """Επιστρέφει το path του τελευταίου panic Excel (ή '' αν δεν υπάρχει)."""
    try:
        p = _panic_path_file()
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return f.read().strip()
    except Exception:
        pass
    return ''


# ── Λήξη PANIC — Διαγραφή εγγραφών ──────────────────────────────────────────

def run_delete(ctx, driver, callback=None):
    """Διαγράφει τις εγγραφές Γραμματειακής Υποστήριξης για κάθε ΑΦΜ."""
    log = callback or print

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    file_path = ctx.get('file_path')
    records   = load_data(file_path, log=log)
    if not records:
        return

    total = len(records)
    log(f'  {total} εγγραφές προς διαγραφή')

    ok = fail = 0

    for idx, record in enumerate(records, 1):
        afm    = record['afm']
        school = record['school']
        log(f'\n[{idx}/{total}] ΑΦΜ: {afm}  |  {school}')

        # ── Σελίδα αναζήτησης ─────────────────────────────────────────────
        driver.get(SEARCH_URL)
        time.sleep(2)

        # ── Συμπλήρωση ΑΦΜ ────────────────────────────────────────────────
        try:
            afm_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.ID, 'ctl00_ContentData_txtTaxNumber_I')))
            driver.execute_script("""
                var el = arguments[0];
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
            """, afm_field, afm)
            time.sleep(1)
        except Exception as e:
            log(f'  ✗ ΑΦΜ field: {e}')
            fail += 1
            continue

        # ── Αναζήτηση ─────────────────────────────────────────────────────
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}')
            fail += 1
            continue

        # ── Εύρεση εγγραφής ───────────────────────────────────────────────
        edit_links = driver.find_elements(
            By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log('  ✗ Κανένα αποτέλεσμα')
            fail += 1
            continue

        target = None
        if len(edit_links) == 1:
            target = edit_links[0]
        else:
            for link in edit_links:
                try:
                    row = link.find_element(By.XPATH, './ancestor::tr[1]')
                    if school.upper() in row.text.upper():
                        target = link
                        log(f'  Βρέθηκε: {school}')
                        break
                except Exception:
                    pass
            if not target:
                log('  Χρήση πρώτης εγγραφής')
                target = edit_links[0]

        # ── Άνοιγμα καρτέλας ──────────────────────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Καρτέλα: {e}')
            fail += 1
            continue

        # ── Κουμπί Διαγραφή (πρώτη γραμμή gridEmplDet) ───────────────────
        try:
            del_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//img[@alt="Διαγραφή" and contains(@onclick,"gridEmplDet")]')))
            driver.execute_script('arguments[0].scrollIntoView({block:"center"});', del_btn)
            time.sleep(0.5)
            driver.execute_script('arguments[0].click();', del_btn)
            time.sleep(1)
            log('  Διαγραφή κλικ')
        except Exception as e:
            log(f'  ✗ Διαγραφή: {e}')
            fail += 1
            continue

        # ── Επιβεβαίωση (OK στο confirm dialog) ──────────────────────────
        try:
            from selenium.webdriver.support.ui import WebDriverWait
            WebDriverWait(driver, 5).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
            time.sleep(2)
            log('  ✓ Επιβεβαίωση ΟΚ')
        except Exception:
            # Δεν υπήρξε alert — ίσως ήταν inline confirm
            log('  ⚠ Alert δεν εμφανίστηκε')

        # ── Αποθήκευση ────────────────────────────────────────────────────
        try:
            save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable(
                    (By.ID, 'ctl00_ContentData_btnSave')))
            driver.execute_script('arguments[0].click();', save_btn)
            time.sleep(3)
            log('  ✓ Αποθήκευση')
            ok += 1
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}')
            fail += 1

    log(f'\n{"─"*50}')
    log(f'Λήξη PANIC — Διαγραφές: {ok} επιτυχείς, {fail} αποτυχίες')
