#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
editor.py
=========
Αυτόματη συμπλήρωση Γραμματειακής Υποστήριξης στο MySchool.

Ροή:
1. Dialog εισαγωγής ΑΦΜ (ένα ή περισσότερα, χωρισμένα με κόμα)
2. Σύνδεση στο MySchool με credentials από config
3. Για κάθε ΑΦΜ: αναζήτηση → ανάγνωση ωρών → άνοιγμα καρτέλας
4. Επιλογή "Γραμματειακή Υποστήριξη" + ώρες + ημερομηνίες (= σήμερα)
5. Αποδοχή → Αποθήκευση
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
    """Εμφανίζει παράθυρο εισαγωγής ΑΦΜ. Επιστρέφει list[str] ή []."""
    import tkinter as tk

    result = []
    root = tk.Tk()
    root.title('Εισαγωγή ΑΦΜ')
    root.resizable(False, False)
    root.configure(bg='#f5f5f5')

    w, h = 420, 200
    root.update_idletasks()
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

    # Άνοιγμα dropdown
    try:
        btn = driver.find_element(By.ID, base_id + '_B-1')
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


# ── Κύρια συνάρτηση ───────────────────────────────────────────────────────────

def run(config):
    """Entry point για CUSTOM_RUN=True."""
    import core.framework as _fw
    _fw._current_check_title = CHECK_TITLE

    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # ── Dialog ΑΦΜ ───────────────────────────────────────────────────────────
    afm_list = _ask_afm_dialog()
    if not afm_list:
        print('Ακύρωση.')
        return

    today_str = date.today().strftime('%d/%m/%Y')
    print('=' * 62)
    print(f'  {CHECK_TITLE}')
    print('=' * 62)
    print(f'  ΑΦΜ: {", ".join(afm_list)}')
    print(f'  Ημερομηνία: {today_str}')
    print('-' * 62)

    # ── Εκκίνηση Chrome ──────────────────────────────────────────────────────
    options = webdriver.ChromeOptions()
    options.add_argument('--window-size=1400,900')
    options.add_argument('--no-sandbox')

    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        print(f'Αδύνατη εκκίνηση Chrome: {e}')
        return

    try:
        # ── Login ─────────────────────────────────────────────────────────────
        print('Σύνδεση στο MySchool...')
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

        driver.get(SEARCH_URL)
        time.sleep(3)
        print('Σύνδεση ΟΚ\n')

        ok = fail = 0
        total = len(afm_list)

        for idx, afm in enumerate(afm_list, 1):
            print(f'[{idx}/{total}] ΑΦΜ: {afm}')

            # ── Σελίδα αναζήτησης ─────────────────────────────────────────
            driver.get(SEARCH_URL)
            time.sleep(2)

            # ── Συμπλήρωση ΑΦΜ ────────────────────────────────────────────
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

            # ── Ανάγνωση διαθέσιμων ωρών ──────────────────────────────────
            hours = ''
            try:
                hours_el = driver.find_element(
                    By.ID, 'ctl00_ContentData_txtAvailableHoursForUnit_I')
                hours = hours_el.get_attribute('value').strip()
                print(f'  Διαθέσιμες ώρες: {hours}')
            except Exception as e:
                print(f'  ⚠ Ώρες (κεντρική): {e}')

            # ── Εύρεση γραναζιού ──────────────────────────────────────────
            edit_links = driver.find_elements(
                By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
            print(f'  {len(edit_links)} αποτέλεσμα(-τα)')

            if not edit_links:
                print(f'  ✗ Κανένα αποτέλεσμα')
                fail += 1
                continue

            # ── Άνοιγμα καρτέλας ──────────────────────────────────────────
            try:
                driver.execute_script('arguments[0].click();', edit_links[0])
                time.sleep(3)
                print('  Καρτέλα ανοιχτή')
            except Exception as e:
                print(f'  ✗ Άνοιγμα καρτέλας: {e}')
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
                print(f'  ✗ Σταυρός Προσθήκης: {e}')
                fail += 1
                continue

            # ── Dropdown "Γραμματειακή Υποστήριξη" ───────────────────────
            combo_base = ('ctl00_ContentData_gridEmplDet_'
                          'editnew_2_cmbWorkHoursDetailsType')
            try:
                ok_c = _select_dxe_combo(driver, combo_base, WORK_TYPE_TEXT)
                print(f'  {"✓" if ok_c else "⚠"} Τύπος: {WORK_TYPE_TEXT}')
            except Exception as e:
                print(f'  ⚠ Dropdown: {e}')
            time.sleep(0.5)

            # ── Ώρες (DXEditor4) ──────────────────────────────────────────
            if hours:
                try:
                    _set_dxe_value(driver,
                        'ctl00_ContentData_gridEmplDet_DXEditor4_I', hours)
                    print(f'  ✓ Ώρες: {hours}')
                except Exception as e:
                    print(f'  ⚠ Ώρες (φόρμα): {e}')

            # ── Ημερομηνία από (DXEditor5) ────────────────────────────────
            try:
                _set_dxe_value(driver,
                    'ctl00_ContentData_gridEmplDet_DXEditor5_I', today_str)
                print(f'  ✓ Ημ. από: {today_str}')
            except Exception as e:
                print(f'  ⚠ Ημ. από: {e}')

            # ── Ημερομηνία έως (DXEditor6) ────────────────────────────────
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

        print(f'\n{"─"*62}')
        print(f'Ολοκλήρωση: {ok} επιτυχείς, {fail} αποτυχίες')

    except Exception as e:
        print(f'Κρίσιμο σφάλμα: {e}')
    finally:
        try:
            input('\nΠάτα Enter για κλείσιμο browser...')
            driver.quit()
        except Exception:
            pass
