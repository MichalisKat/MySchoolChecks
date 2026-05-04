#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
editor.py
=========
Αυτόματη αναζήτηση και άνοιγμα καρτέλας εκπαιδευτικού στο MySchool.

Ροή:
1. Διάβασμα Excel/CSV (ΑΦΜ, Ονομασία Σχολείου)
2. Αναζήτηση με ΑΦΜ στη σελίδα Worker.list.myEmplUnit.aspx
3. Αν πολλά αποτελέσματα → επιλογή με βάση "Φορέας τοποθέτησης"
4. Κλικ γρανάζι → άνοιγμα καρτέλας → κλικ σταυρός "Λεπτομέρειες ωραρίου"
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Editor - Επεξεργασία Καρτέλας Εκπαιδευτικού'
CHECK_DESCRIPTION = 'Αυτόματη αναζήτηση και άνοιγμα καρτέλας εκπαιδευτικού'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL   = 'https://app.myschool.sch.gr'
SEARCH_URL = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15


# ── Ανάγνωση αρχείου ─────────────────────────────────────────────────────────

def load_data(file_path, log=print):
    """
    Διαβάζει Excel ή CSV και επιστρέφει list[dict] με 'afm' και 'school'.
    Αναγνωρίζει αυτόματα τις στήλες (ΑΦΜ, Α.Φ.Μ., Ονομασία Σχολείου κλπ).
    """
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(file_path, dtype=str)
        else:
            # CSV: δοκιμή με ημιτελείο και UTF-8/ISO
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

    # Αναγνώριση στηλών ΑΦΜ και Σχολείο
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
        afm = str(row.get(afm_col, '')).strip()
        school = str(row.get(school_col, '')).strip() if school_col else ''
        if afm and afm not in ('nan', 'None', ''):
            records.append({'afm': afm.zfill(9), 'school': school})

    log(f'Φορτώθηκαν {len(records)} εγγραφές')
    return records


# ── Σύνδεση ───────────────────────────────────────────────────────────────────

def connect(log=print):
    """
    Ανοίγει Chrome, κάνει SSO login και πλοηγείται στη σελίδα αναζήτησης.
    Επιστρέφει το driver ή None σε αποτυχία.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError:
        log('pip install selenium')
        return None

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
                    '#username, input[name="username"], input[type="text"]'))
            )
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

        log('Έτοιμο για επεξεργασία καρτελών')
        return driver

    except Exception as e:
        log(f'Σφάλμα σύνδεσης: {e}')
        try:
            driver.quit()
        except Exception:
            pass
        return None


# ── Κύριος βρόχος ─────────────────────────────────────────────────────────────

def run(ctx, driver, callback=None):
    """
    Επεξεργάζεται κάθε εγγραφή του αρχείου:
    αναζήτηση ΑΦΜ → επιλογή εγγραφής → άνοιγμα καρτέλας → κλικ Προσθήκη ωραρίου.
    """
    log = callback or print

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    file_path = ctx.get('file_path')
    records   = load_data(file_path, log=log)

    if not records:
        log('Κανένα δεδομένο στο αρχείο')
        return

    total = len(records)
    log(f'{total} εγγραφές προς επεξεργασία')

    for idx, record in enumerate(records, 1):
        afm    = record['afm']
        school = record['school']

        log(f'\n[{idx}/{total}] ΑΦΜ: {afm}  |  {school}')

        # Επιστροφή στη σελίδα αναζήτησης (reload για καθαρή κατάσταση)
        driver.get(SEARCH_URL)
        time.sleep(2)

        # Συμπλήρωση ΑΦΜ
        try:
            afm_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.NAME, 'ctl00$ContentData$txtTaxNumber')))
            afm_field.clear()
            afm_field.send_keys(afm)
        except Exception as e:
            log(f'  ΑΦΜ field: {e}')
            continue

        # Αναζήτηση
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  Αναζήτηση: {e}')
            continue

        # Εύρεση γραναζιών (εικονίδιο Διόρθωσης)
        edit_links = driver.find_elements(
            By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log(f'  Κανένα αποτέλεσμα για ΑΦΜ={afm}')
            continue

        # Επιλογή εγγραφής
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
                log('  Χρήση πρώτης εγγραφής (δεν ταίριαξε σχολείο)')
                target = edit_links[0]

        # Άνοιγμα καρτέλας
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  Άνοιγμα καρτέλας: {e}')
            continue

        # Κλικ σταυρός Προσθήκης "Λεπτομέρειες ωραρίου"
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
            log('  Προσθήκη ωραρίου: ανοιχτή')
        except Exception as e:
            log(f'  Σταυρός Προσθήκης: {e}')

    log('\nΟλοκλήρωση')
