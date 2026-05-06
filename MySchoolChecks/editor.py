#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
editor.py
=========
Αυτόματη συμπλήρωση Γραμματειακής Υποστήριξης στο MySchool.

Ροή:
1. Dialog εισαγωγής ΑΦΜ (ένα ή περισσότερα, χωρισμένα με κόμα)
2. Αναζήτηση κάθε ΑΦΜ στη σελίδα Worker.list.myEmplUnit.aspx
3. Ανάγνωση διαθέσιμων ωρών (txtAvailableHoursForUnit)
4. Άνοιγμα καρτέλας → σταυρός Προσθήκης
5. Επιλογή "Γραμματειακή Υποστήριξη" από dropdown
6. Συμπλήρωση ωρών (DXEditor4) + ημ. από/έως (DXEditor5/6) = σήμερα
7. Αποδοχή → Αποθήκευση
"""

import os
import sys
import time
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Editor - Γραμματειακή Υποστήριξη'
CHECK_DESCRIPTION = 'Αυτόματη συμπλήρωση Γραμματειακής Υποστήριξης στο MySchool'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL   = 'https://app.myschool.sch.gr'
SEARCH_URL = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

WORK_TYPE_TEXT = 'Γραμματειακή Υποστήριξη'


# ── Dialog εισαγωγής ΑΦΜ ─────────────────────────────────────────────────────

def _ask_afm_dialog():
    """
    Εμφανίζει παράθυρο εισαγωγής ΑΦΜ.
    Επιστρέφει list[str] με τα ΑΦΜ ή [] αν ακυρωθεί.
    """
    import tkinter as tk
    from tkinter import ttk

    result = []

    root = tk.Tk()
    root.title('Εισαγωγή ΑΦΜ')
    root.resizable(False, False)
    root.configure(bg='#f5f5f5')

    # Κεντράρισμα
    root.update_idletasks()
    w, h = 420, 200
    x = (root.winfo_screenwidth()  - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f'{w}x{h}+{x}+{y}')

    tk.Label(root, text='Εισαγωγή ΑΦΜ εκπαιδευτικού/ών',
             bg='#f5f5f5', font=('Arial', 11, 'bold')).pack(pady=(18, 4))
    tk.Label(root, text='Ένα ή περισσότερα ΑΦΜ χωρισμένα με κόμα:',
             bg='#f5f5f5', font=('Arial', 9)).pack()

    entry_var = tk.StringVar()
    entry = tk.Entry(root, textvariable=entry_var, font=('Arial', 11), width=34)
    entry.pack(pady=10, ipady=4)
    entry.focus_set()

    def _ok(event=None):
        raw = entry_var.get().strip()
        if not raw:
            return
        afms = [a.strip().zfill(9) for a in raw.replace(';', ',').split(',') if a.strip()]
        if afms:
            result.extend(afms)
            root.destroy()

    def _cancel():
        root.destroy()

    btn_frame = tk.Frame(root, bg='#f5f5f5')
    btn_frame.pack()
    tk.Button(btn_frame, text='ΟΚ', width=10, command=_ok,
              bg='#1a73e8', fg='white', font=('Arial', 9, 'bold'),
              relief='flat').pack(side='left', padx=6)
    tk.Button(btn_frame, text='Ακύρωση', width=10, command=_cancel,
              font=('Arial', 9), relief='flat').pack(side='left', padx=6)

    root.bind('<Return>', _ok)
    root.bind('<Escape>', lambda e: _cancel())
    root.mainloop()

    return result


# ── Βοηθητικές ───────────────────────────────────────────────────────────────

def _set_dxe_value(driver, element_id, value):
    """Ορίζει τιμή σε DevExpress input μέσω JS."""
    js = """
        var inp = document.getElementById(arguments[0]);
        if (!inp) return false;
        inp.value = arguments[1];
        inp.dispatchEvent(new Event('change', {bubbles: true}));
        var base = arguments[0].replace('_I', '');
        if (typeof aspxETextChanged === 'function') aspxETextChanged(base);
        if (typeof aspxEValueChanged === 'function') aspxEValueChanged(base);
        return true;
    """
    return driver.execute_script(js, element_id, value)


def _select_dxe_combo(driver, base_id, text):
    """Επιλέγει τιμή από DevExpress ComboBox."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    button_id = base_id + '_B-1'

    # Άνοιγμα dropdown
    try:
        btn = driver.find_element(By.ID, button_id)
        driver.execute_script('arguments[0].click();', btn)
        time.sleep(1)
    except Exception:
        pass

    # Επιλογή από λίστα
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

def connect(log=print):
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
            import config as _cfg
            user_f = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    '#username, input[name="username"], input[type="text"]')))
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

        log('Έτοιμο')
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
    log = callback or print

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # Dialog εισαγωγής ΑΦΜ
    afm_list = _ask_afm_dialog()
    if not afm_list:
        log('Ακύρωση — δεν δόθηκαν ΑΦΜ.')
        return

    today_str = date.today().strftime('%d/%m/%Y')
    total     = len(afm_list)
    log(f'{total} ΑΦΜ προς επεξεργασία  |  Ημερομηνία: {today_str}')

    ok = fail = 0

    for idx, afm in enumerate(afm_list, 1):
        log(f'\n[{idx}/{total}] ΑΦΜ: {afm}')

        # ── 1. Σελίδα αναζήτησης ─────────────────────────────────────────────
        driver.get(SEARCH_URL)
        time.sleep(2)

        # ── 2. Συμπλήρωση ΑΦΜ ────────────────────────────────────────────────
        try:
            afm_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.NAME, 'ctl00$ContentData$txtTaxNumber')))
            afm_field.clear()
            afm_field.send_keys(afm)
        except Exception as e:
            log(f'  ✗ ΑΦΜ field: {e}')
            fail += 1
            continue

        # ── 3. Αναζήτηση ─────────────────────────────────────────────────────
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}')
            fail += 1
            continue

        # ── 4. Ανάγνωση διαθέσιμων ωρών ─────────────────────────────────────
        hours = ''
        try:
            hours_el = driver.find_element(
                By.ID, 'ctl00_ContentData_txtAvailableHoursForUnit_I')
            hours = hours_el.get_attribute('value').strip()
            log(f'  Διαθέσιμες ώρες: {hours}')
        except Exception as e:
            log(f'  ⚠ Ώρες (κεντρική): {e}')

        # ── 5. Εύρεση γραναζιού ──────────────────────────────────────────────
        edit_links = driver.find_elements(
            By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log(f'  ✗ Κανένα αποτέλεσμα')
            fail += 1
            continue

        target = edit_links[0]

        # ── 6. Άνοιγμα καρτέλας ──────────────────────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Άνοιγμα καρτέλας: {e}')
            fail += 1
            continue

        # ── 7. Σταυρός Προσθήκης ─────────────────────────────────────────────
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
            log(f'  ✗ Σταυρός Προσθήκης: {e}')
            fail += 1
            continue

        # ── 8. Dropdown "Γραμματειακή Υποστήριξη" ────────────────────────────
        combo_base = 'ctl00_ContentData_gridEmplDet_editnew_2_cmbWorkHoursDetailsType'
        try:
            ok_combo = _select_dxe_combo(driver, combo_base, WORK_TYPE_TEXT)
            log(f'  {"✓" if ok_combo else "⚠"} Τύπος: {WORK_TYPE_TEXT}')
        except Exception as e:
            log(f'  ⚠ Dropdown: {e}')
        time.sleep(0.5)

        # ── 9. Ώρες (DXEditor4) ───────────────────────────────────────────────
        if hours:
            try:
                _set_dxe_value(driver,
                    'ctl00_ContentData_gridEmplDet_DXEditor4_I', hours)
                log(f'  ✓ Ώρες: {hours}')
            except Exception as e:
                log(f'  ⚠ Ώρες (φόρμα): {e}')

        # ── 10. Ημερομηνία από (DXEditor5) ───────────────────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor5_I', today_str)
            log(f'  ✓ Ημ. από: {today_str}')
        except Exception as e:
            log(f'  ⚠ Ημ. από: {e}')

        # ── 11. Ημερομηνία έως (DXEditor6) ───────────────────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor6_I', today_str)
            log(f'  ✓ Ημ. έως: {today_str}')
        except Exception as e:
            log(f'  ⚠ Ημ. έως: {e}')

        time.sleep(0.5)

        # ── 12. Αποδοχή ──────────────────────────────────────────────────────
        try:
            driver.execute_script(
                "aspxGVScheduleCommand('ctl00_ContentData_gridEmplDet',"
                "['UpdateEdit'],1)")
            time.sleep(2)
            log('  ✓ Αποδοχή')
        except Exception as e:
            log(f'  ✗ Αποδοχή: {e}')
            fail += 1
            continue

        # ── 13. Αποθήκευση ───────────────────────────────────────────────────
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
