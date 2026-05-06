#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
editor.py
=========
Αυτόματη συμπλήρωση Γραμματειακής Υποστήριξης στο MySchool.

Ροή:
1. Επιλογή Excel/CSV με στήλες: Α.Φ.Μ., Ονομασία Σχολείου (προαιρετική)
2. Σύνδεση στο MySchool
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
    """
    Διαβάζει Excel ή CSV.
    Επιστρέφει list[dict] με 'afm' και 'school'.
    Υποχρεωτική στήλη: Α.Φ.Μ.
    Προαιρετική: Ονομασία Σχολείου (για διαλογή όταν υπάρχουν πολλές τοποθετήσεις).
    """
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

def _set_dxe_value(driver, element_id, value):
    """Ορίζει τιμή σε DevExpress input μέσω JS."""
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
    """Επιλέγει τιμή από DevExpress ComboBox."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    try:
        btn = driver.find_element(By.ID, base_id + '_B-1')
        driver.execute_script('arguments[0].click();', btn)
        time.sleep(1)
    except Exception:
        pass

    try:
        items = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, '.dxeListBoxItem, .dxeLBItem')))
        for item in items:
            if text in item.text:
                driver.execute_script('arguments[0].click();', item)
                time.sleep(0.5)
                return True
    except Exception:
        pass

    # Fallback: πληκτρολόγηση
    try:
        inp = driver.find_element(By.ID, base_id + '_I')
        inp.clear()
        inp.send_keys(text)
        time.sleep(0.3)
        driver.execute_script(f"aspxETextChanged('{base_id}');")
        time.sleep(0.5)
        try:
            first = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '.dxeListBoxItem, .dxeLBItem')))
            driver.execute_script('arguments[0].click();', first)
        except Exception:
            pass
        return True
    except Exception:
        return False


# ── Σύνδεση ───────────────────────────────────────────────────────────────────

def connect(config, log=print):
    """Ανοίγει Chrome και συνδέεται στο MySchool."""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1400,900')
    options.add_argument('--no-sandbox')

    try:
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
            user_f.send_keys(config.MYSCHOOL_USER)

            pass_f = driver.find_element(By.CSS_SELECTOR,
                '#password, input[name="password"], input[type="password"]')
            pass_f.clear()
            pass_f.send_keys(config.MYSCHOOL_PASS)

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


# ── Κύριος βρόχος ─────────────────────────────────────────────────────────────

def run(config):
    """Entry point για CUSTOM_RUN=True."""
    import core.framework as _fw
    _fw._current_check_title = CHECK_TITLE

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # Το framework έχει ήδη ζητήσει το αρχείο — το παίρνουμε από εκεί
    file_path = getattr(_fw, '_current_file_path', None)
    if not file_path or not os.path.exists(file_path):
        from core.framework import ask_file
        file_path = ask_file('Αρχείο εκπαιδευτικών (Excel ή CSV):')
        if not file_path:
            print('Ακύρωση.')
            return

    records = load_data(file_path)
    if not records:
        return

    driver = connect(config)
    if not driver:
        return

    today_str = date.today().strftime('%d/%m/%Y')
    total = len(records)
    print(f'\n  {total} εγγραφές  |  Ημερομηνία: {today_str}')
    print('─' * 62)

    ok = fail = 0

    try:
        for idx, record in enumerate(records, 1):
            afm    = record['afm']
            school = record['school']
            print(f'\n[{idx}/{total}] ΑΦΜ: {afm}  |  {school}')

            driver.get(SEARCH_URL)
            time.sleep(2)

            # ── ΑΦΜ ───────────────────────────────────────────────────────
            try:
                afm_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                    EC.presence_of_element_located(
                        (By.NAME, 'ctl00$ContentData$txtTaxNumber')))
                afm_field.clear()
                afm_field.send_keys(afm)
            except Exception as e:
                print(f'  ✗ ΑΦΜ field: {e}')
                fail += 1
                continue

            # ── Αναζήτηση ─────────────────────────────────────────────────
            try:
                search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
                driver.execute_script('arguments[0].click();', search_link)
                time.sleep(3)
            except Exception as e:
                print(f'  ✗ Αναζήτηση: {e}')
                fail += 1
                continue

            # ── Ώρες από κεντρική ─────────────────────────────────────────
            hours = ''
            try:
                hours_el = driver.find_element(
                    By.ID, 'ctl00_ContentData_txtAvailableHoursForUnit_I')
                hours = hours_el.get_attribute('value').strip()
                print(f'  Διαθέσιμες ώρες: {hours}')
            except Exception as e:
                print(f'  ⚠ Ώρες: {e}')

            # ── Εύρεση εγγραφής ───────────────────────────────────────────
            edit_links = driver.find_elements(
                By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
            print(f'  {len(edit_links)} αποτέλεσμα(-τα)')

            if not edit_links:
                print('  ✗ Κανένα αποτέλεσμα')
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
                            print(f'  Βρέθηκε: {school}')
                            break
                    except Exception:
                        pass
                if not target:
                    print('  Χρήση πρώτης εγγραφής')
                    target = edit_links[0]

            # ── Άνοιγμα καρτέλας ──────────────────────────────────────────
            try:
                driver.execute_script('arguments[0].click();', target)
                time.sleep(3)
                print('  Καρτέλα ανοιχτή')
            except Exception as e:
                print(f'  ✗ Καρτέλα: {e}')
                fail += 1
                continue

            # ── Σταυρός Προσθήκης ─────────────────────────────────────────
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
                print('  Φόρμα ωραρίου ανοιχτή')
            except Exception as e:
                print(f'  ✗ Σταυρός: {e}')
                fail += 1
                continue

            # ── Dropdown ──────────────────────────────────────────────────
            combo_base = ('ctl00_ContentData_gridEmplDet_'
                          'editnew_2_cmbWorkHoursDetailsType')
            try:
                ok_c = _select_dxe_combo(driver, combo_base, WORK_TYPE_TEXT)
                print(f'  {"✓" if ok_c else "⚠"} Τύπος: {WORK_TYPE_TEXT}')
            except Exception as e:
                print(f'  ⚠ Dropdown: {e}')
            time.sleep(0.5)

            # ── Ώρες φόρμα (DXEditor4) ────────────────────────────────────
            if hours:
                try:
                    _set_dxe_value(driver,
                        'ctl00_ContentData_gridEmplDet_DXEditor4_I', hours)
                    print(f'  ✓ Ώρες: {hours}')
                except Exception as e:
                    print(f'  ⚠ Ώρες φόρμα: {e}')

            # ── Ημ. από (DXEditor5) ───────────────────────────────────────
            try:
                _set_dxe_value(driver,
                    'ctl00_ContentData_gridEmplDet_DXEditor5_I', today_str)
                print(f'  ✓ Ημ. από: {today_str}')
            except Exception as e:
                print(f'  ⚠ Ημ. από: {e}')

            # ── Ημ. έως (DXEditor6) ───────────────────────────────────────
            try:
                _set_dxe_value(driver,
                    'ctl00_ContentData_gridEmplDet_DXEditor6_I', today_str)
                print(f'  ✓ Ημ. έως: {today_str}')
            except Exception as e:
                print(f'  ⚠ Ημ. έως: {e}')

            time.sleep(0.5)

            # ── Αποδοχή ───────────────────────────────────────────────────
            try:
                driver.execute_script(
                    "aspxGVScheduleCommand('ctl00_ContentData_gridEmplDet',"
                    "['UpdateEdit'],1)")
                time.sleep(2)
                print('  ✓ Αποδοχή')
            except Exception as e:
                print(f'  ✗ Αποδοχή: {e}')
                fail += 1
                continue

            # ── Αποθήκευση ────────────────────────────────────────────────
            try:
                save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                    EC.element_to_be_clickable(
                        (By.ID, 'ctl00_ContentData_btnSave')))
                driver.execute_script('arguments[0].click();', save_btn)
                time.sleep(3)
                print('  ✓ Αποθήκευση')
                ok += 1
            except Exception as e:
                print(f'  ✗ Αποθήκευση: {e}')
                fail += 1

    finally:
        print(f'\n{"─"*62}')
        print(f'Ολοκλήρωση: {ok} επιτυχείς, {fail} αποτυχίες')
        try:
            driver.quit()
        except Exception:
            pass
