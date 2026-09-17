#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
diorthosi_imeromhnias_enarxis.py
==================================
One-off εργασία (ΔΕΝ μπαίνει στο μενού της εφαρμογής) — διόρθωση της
Ημερομηνίας Έναρξης Τοποθέτησης από 1/9/2026 σε 7/9/2026, για μια λίστα
τοποθετήσεων από Excel.

Βάλε αυτό το αρχείο μέσα στον φάκελο MySchoolChecks\\ (δίπλα στο main.py,
config.py κλπ.) και τρέξε το με:

    python diorthosi_imeromhnias_enarxis.py "ΔΙΟΡΘΩΣΗ_ΑΠΟ.xlsx"

Excel — αναμενόμενες στήλες (ακριβώς όπως στο αρχείο που δόθηκε):
    ΕΙΔΟΣ ΤΟΠΟΘΕΤΗΣΗΣ | Α.Φ.Μ. | ΕΠΙΘΕΤΟ | ΟΝΟΜΑ | ΚΩΔ. ΣΧΟΛΕΙΟΥ | ΣΧΟΛΕΙΟ

ΣΗΜΑΝΤΙΚΟ: κάθε ΓΡΑΜΜΗ του excel είναι μία ΞΕΧΩΡΙΣΤΗ τοποθέτηση προς
διόρθωση (όχι ένα άτομο) — το ίδιο Α.Φ.Μ. μπορεί να εμφανίζεται πολλές
φορές (π.χ. μία «Από Διάθεση ΠΥΣΠΕ/ΠΥΣΔΕ» + πολλές «Μερική Διάθεση») και
ΟΛΕΣ διορθώνονται ανεξάρτητα.

Ροή ανά γραμμή excel:
  1. Αναζήτηση στο https://app.myschool.sch.gr/Worker.list.myEmplUnit.aspx
     με βάση το Α.Φ.Μ. (ίδιο μοτίβο με το editor.py / Ε5).
  2. Στα αποτελέσματα (συνήθως πάνω από μία γραμμές, μία ανά τοποθέτηση
     του ατόμου) εντοπίζεται η ΜΟΝΑΔΙΚΗ γραμμή που έχει ταυτόχρονα:
       • Είδος Τοποθέτησης = της στήλης «ΕΙΔΟΣ ΤΟΠΟΘΕΤΗΣΗΣ» του excel
         (case-insensitive substring match)
       • Σχολείο = αυτό της στήλης «ΣΧΟΛΕΙΟ» του excel (fuzzy match, ίδια
         λογική κανονικοποίησης με το placements.py).
     Αν βρεθούν 0 ή >1 τέτοιες γραμμές → η εγγραφή ΑΓΝΟΕΙΤΑΙ και καταγράφεται
     (καμία αλλαγή) — δεν μαντεύουμε ποτέ.
  3. Άνοιγμα της καρτέλας (κλικ στο «Διόρθωση»).
  4. Ασφαλής επανέλεγχος: το πεδίο Ημερομηνίας Έναρξης
     (ctl00_ContentData_dtDutyStartDate) στην καρτέλα πρέπει να δείχνει
     ήδη ΑΚΡΙΒΩΣ 1/9/2026 — αν όχι, ΔΕΝ αγγίζουμε τίποτα.
  5. Αλλαγή του πεδίου σε 7/9/2026 (καθαρό text field — χωρίς dropdown).
  6. Επιβεβαίωση ΠΡΙΝ την αποθήκευση ότι το πεδίο δείχνει όντως 7/9/2026.
  7. Κλικ «Αποθήκευση». ΧΩΡΙΣ επιβεβαίωση μετά την αποθήκευση (όπως και
     στο allagi_sxesis_topothetisis.py).

Τρέχει με ΟΡΑΤΟ Chrome, μία τοποθέτηση τη φορά, με πλήρες log στην
κονσόλα — μπορείς να το παρακολουθείς live. Αν κάτι δείχνει λάθος, κλείσε
το παράθυρο της κονσόλας (Ctrl+C) οποιαδήποτε στιγμή· ό,τι έχει ήδη
αποθηκευτεί μέχρι εκείνο το σημείο παραμένει (δεν γίνεται rollback), αλλά
τίποτα παραπέρα δεν θα αγγιχτεί.
"""

import os
import sys
import time
from datetime import datetime

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE_URL     = 'https://app.myschool.sch.gr'
SEARCH_URL   = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

FROM_DATE_TEXT = '1/9/2026'
TO_DATE_TEXT   = '7/9/2026'

DATE_BASE   = 'ctl00_ContentData_dtDutyStartDate'
SAVE_BTN_ID = 'ctl00_ContentData_btnSave'

STRIKE_INTERVAL = 0.3


def _parse_date(s):
    """Επιστρέφει (day,month,year) ή None αν δεν αναγνωρίζεται."""
    s = (s or '').strip()
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d-%m-%y'):
        try:
            d = datetime.strptime(s, fmt)
            return (d.day, d.month, d.year if d.year > 100 else 2000 + d.year)
        except ValueError:
            continue
    return None


def _dates_equal(s1, s2):
    d1, d2 = _parse_date(s1), _parse_date(s2)
    return d1 is not None and d1 == d2


# ── Ανάγνωση Excel ───────────────────────────────────────────────────────────
def load_rows(file_path, log=print):
    """Διαβάζει το excel και επιστρέφει λίστα από dict
    {'afm','eponymo','onoma','eidos','sxol_kod','sxoleio'} — μία ανά γραμμή
    (=τοποθέτηση προς διόρθωση)."""
    df = pd.read_excel(file_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    def _find(candidates):
        for c in candidates:
            if c in df.columns:
                return c
        low = {x.lower(): x for x in df.columns}
        for c in candidates:
            if c.lower() in low:
                return low[c.lower()]
        return None

    afm_col     = _find(('Α.Φ.Μ.', 'ΑΦΜ', 'Α.Φ.Μ'))
    epon_col    = _find(('ΕΠΙΘΕΤΟ', 'Επώνυμο'))
    onoma_col   = _find(('ΟΝΟΜΑ', 'Όνομα'))
    eidos_col   = _find(('ΕΙΔΟΣ ΤΟΠΟΘΕΤΗΣΗΣ',))
    kod_col     = _find(('ΚΩΔ. ΣΧΟΛΕΙΟΥ',))
    sxoleio_col = _find(('ΣΧΟΛΕΙΟ', 'Ονομασία'))

    missing = [n for n, v in (
        ('Α.Φ.Μ.', afm_col), ('ΕΙΔΟΣ ΤΟΠΟΘΕΤΗΣΗΣ', eidos_col), ('ΣΧΟΛΕΙΟ', sxoleio_col),
    ) if v is None]
    if missing:
        raise ValueError(f'Λείπουν στήλες: {", ".join(missing)}. Διαθέσιμες: {list(df.columns)}')

    rows = []
    for _, row in df.iterrows():
        afm = str(row.get(afm_col, '')).strip()
        if not afm or afm.lower() in ('nan', 'none', ''):
            continue
        rows.append({
            'afm':      afm,
            'eponymo':  str(row.get(epon_col, '')).strip() if epon_col else '',
            'onoma':    str(row.get(onoma_col, '')).strip() if onoma_col else '',
            'eidos':    str(row.get(eidos_col, '')).strip(),
            'sxol_kod': str(row.get(kod_col, '')).strip() if kod_col else '',
            'sxoleio':  str(row.get(sxoleio_col, '')).strip(),
        })
    log(f'  ✓ Διαβάστηκαν {len(rows)} γραμμές (τοποθετήσεις) από το excel.')
    return rows


# ── Selenium βοηθητικά ───────────────────────────────────────────────────────
def _send_keys_slow(element, text, delay=STRIKE_INTERVAL):
    for char in str(text):
        element.send_keys(char)
        time.sleep(delay)


def _set_date_field(driver, base_id, date_str, log):
    """Καθαρίζει και συμπληρώνει ένα απλό text/date πεδίο (χωρίς dropdown,
    ίδιο μοτίβο με το _set_date του placements.py — αποδεδειγμένα δουλεύει
    εκεί). Επιστρέφει True/False (True = το πεδίο δείχνει την σωστή τιμή)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.common.keys import Keys

    inp = driver.find_element(By.ID, base_id + '_I')
    driver.execute_script('arguments[0].click();', inp)
    time.sleep(0.3)

    inp.click()
    inp.send_keys(Keys.CONTROL, 'a')
    inp.send_keys(Keys.DELETE)
    time.sleep(0.2)
    if (inp.get_attribute('value') or '').strip():
        inp.clear()
        time.sleep(0.2)
    if (inp.get_attribute('value') or '').strip():
        cur_len = len(inp.get_attribute('value') or '')
        for _ in range(cur_len + 5):
            inp.send_keys(Keys.BACKSPACE)
        time.sleep(0.2)
    if (inp.get_attribute('value') or '').strip():
        driver.execute_script("""
            var el = arguments[0];
            el.value = '';
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        """, inp)
        time.sleep(0.2)

    log(f'    (μετά το καθάρισμα: τιμή στο πεδίο = "{inp.get_attribute("value")}")')

    _send_keys_slow(inp, date_str)
    time.sleep(0.5)

    # Κλικ αλλού ώστε να "κλειδώσει" η τιμή (onblur/onchange) πριν διαβάσουμε
    driver.execute_script('arguments[0].blur();', inp)
    time.sleep(0.5)

    final_val = (inp.get_attribute('value') or '').strip()
    log(f'    (μετά την πληκτρολόγηση «{date_str}»: τιμή στο πεδίο = "{final_val}")')
    return _dates_equal(final_val, date_str)


def connect(log=print):
    """Άνοιγμα ορατού Chrome + login στο MySchool — ίδιο μοτίβο με
    editor.py / dioikitiko_ergo_entry.py."""
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


# ── Κανονικοποίηση ονόματος σχολείου (ίδια λογική με placements.py) ────────
def _normalize_school_name(s):
    import placements as _pl
    return _pl._normalize_school_name(s)


def _search_by_afm(driver, afm, log):
    """Πάει στη σελίδα αναζήτησης, συμπληρώνει Α.Φ.Μ., πατάει αναζήτηση.
    Επιστρέφει λίστα <a> edit links (Διόρθωση)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(SEARCH_URL)
    time.sleep(2)

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

    search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
    driver.execute_script('arguments[0].click();', search_link)
    time.sleep(3)

    return driver.find_elements(By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')


def _pick_matching_row(driver, edit_links, eidos, target_school_norm, log):
    """
    Ανάμεσα στα edit_links, βρίσκει τη ΜΟΝΑΔΙΚΗ γραμμή που έχει:
      • Είδος Τοποθέτησης = eidos (case-insensitive substring)
      • σχολείο που ταιριάζει (κανονικοποιημένα) με το target_school_norm
    Επιστρέφει (link, reason) — reason ∈ {'ok','notfound','ambiguous'}.
    """
    from selenium.webdriver.common.by import By

    eidos_up = (eidos or '').strip().upper()
    candidates = []
    for link in edit_links:
        try:
            row = link.find_element(By.XPATH, './ancestor::tr[1]')
            row_text = row.text
        except Exception:
            continue

        if eidos_up and eidos_up not in row_text.upper():
            continue

        row_norm = _normalize_school_name(row_text)
        if target_school_norm and target_school_norm in row_norm:
            candidates.append(link)

    if len(candidates) == 1:
        return candidates[0], 'ok'
    if len(candidates) == 0:
        return None, 'notfound'
    return None, 'ambiguous'


def process_row(driver, item, log):
    """Επεξεργάζεται μία γραμμή (=τοποθέτηση). Επιστρέφει status string:
    'ok' | 'notfound' | 'ambiguous' | 'value_mismatch' | 'error'."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    afm     = item['afm']
    eidos   = item['eidos']
    sxoleio = item['sxoleio']
    target_school_norm = _normalize_school_name(sxoleio)

    try:
        edit_links = _search_by_afm(driver, afm, log)
    except Exception as e:
        log(f'  ✗ Αναζήτηση απέτυχε: {e}')
        return 'error'

    if not edit_links:
        log('  ✗ Κανένα αποτέλεσμα αναζήτησης')
        return 'notfound'

    link, reason = _pick_matching_row(driver, edit_links, eidos, target_school_norm, log)
    if reason != 'ok':
        if reason == 'notfound':
            log(f'  ⚠ Δεν βρέθηκε γραμμή με «{eidos}» + σχολείο «{sxoleio}» '
                f'({len(edit_links)} αποτέλεσμα(-τα) συνολικά) — παράλειψη')
        else:
            log(f'  ⚠ Πάνω από μία γραμμές ταιριάζουν — παράλειψη (χειροκίνητος έλεγχος)')
        return reason

    # ── Άνοιγμα καρτέλας ─────────────────────────────────────────────────
    try:
        driver.execute_script('arguments[0].click();', link)
        time.sleep(2.5)
    except Exception as e:
        log(f'  ✗ Άνοιγμα καρτέλας: {e}')
        return 'error'

    # ── Ασφαλής επανέλεγχος τιμής πριν αλλάξουμε οτιδήποτε ─────────────────
    try:
        date_inp = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.ID, DATE_BASE + '_I')))
        current_val = (date_inp.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ✗ Δεν βρέθηκε το πεδίο Ημερομηνίας Έναρξης στην καρτέλα: {e}')
        return 'error'

    if not _dates_equal(current_val, FROM_DATE_TEXT):
        log(f'  ⚠ Η τιμή στην καρτέλα είναι «{current_val}» (όχι «{FROM_DATE_TEXT}») — παράλειψη, καμία αλλαγή')
        return 'value_mismatch'

    # ── Αλλαγή τιμής ─────────────────────────────────────────────────────
    ok_c = _set_date_field(driver, DATE_BASE, TO_DATE_TEXT, log)
    if not ok_c:
        log('  ✗ Αποτυχία συμπλήρωσης νέας ημερομηνίας — καμία αλλαγή')
        return 'error'
    log(f'  ✓ Συμπληρώθηκε «{TO_DATE_TEXT}» (επιβεβαιώθηκε στη φόρμα)')

    # ── Αποθήκευση ────────────────────────────────────────────────────────
    try:
        save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.element_to_be_clickable((By.ID, SAVE_BTN_ID)))
        driver.execute_script('arguments[0].click();', save_btn)
        time.sleep(3)
        log('  ✓ Αποθήκευση')
        return 'ok'
    except Exception as e:
        log(f'  ✗ Αποθήκευση: {e}')
        return 'error'


def _ask_for_file():
    """Ανοίγει παράθυρο επιλογής αρχείου (Windows file dialog) και επιστρέφει
    τη διαδρομή του excel — ή None αν ο χρήστης πατήσει Άκυρο."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askopenfilename(
        title='Επίλεξε το αρχείο Excel (ΕΙΔΟΣ ΤΟΠΟΘΕΤΗΣΗΣ, Α.Φ.Μ., ..., ΣΧΟΛΕΙΟ)',
        filetypes=[('Excel αρχεία', '*.xlsx *.xls'), ('Όλα τα αρχεία', '*.*')],
    )
    root.destroy()
    return path or None


# ── Main ──────────────────────────────────────────────────────────────────
def main():
    limit = None

    if len(sys.argv) >= 2:
        file_path = sys.argv[1]
        limit = int(sys.argv[2]) if len(sys.argv) > 2 else None
    else:
        file_path = _ask_for_file()
        if not file_path:
            print('Δεν επιλέχθηκε αρχείο — τέλος.')
            sys.exit(1)

    print('=' * 65)
    print(f'  Διόρθωση Ημερομηνίας Έναρξης Τοποθέτησης — {FROM_DATE_TEXT} → {TO_DATE_TEXT}')
    print('=' * 65)

    rows = load_rows(file_path)
    if limit:
        rows = rows[:limit]
        print(f'  (δοκιμαστικό τρέξιμο — μόνο οι πρώτες {limit} γραμμές)')

    driver = connect()
    if driver is None:
        print('✗ Αποτυχία σύνδεσης — τέλος.')
        sys.exit(1)

    results = {'ok': [], 'notfound': [], 'ambiguous': [], 'value_mismatch': [], 'error': []}

    total = len(rows)
    try:
        for idx, item in enumerate(rows, 1):
            print(f'\n[{idx}/{total}] Α.Φ.Μ. {item["afm"]}  {item["eponymo"]} {item["onoma"]}  '
                  f'—  {item["eidos"]}  —  {item["sxoleio"]}')
            status = process_row(driver, item, print)
            results[status].append(f'{item["afm"]} ({item["eponymo"]} {item["onoma"]} - {item["sxoleio"]})')
            time.sleep(0.5)
    except KeyboardInterrupt:
        print('\n\n⚠ Διακόπηκε από τον χρήστη (Ctrl+C). Ό,τι έχει ήδη αποθηκευτεί παραμένει.')
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    print('\n' + '─' * 65)
    print(f'✓ Αλλάχθηκαν επιτυχώς : {len(results["ok"])}')
    print(f'⚠ Δεν βρέθηκε ταίριασμα (είδος + σχολείο) : {len(results["notfound"])}')
    print(f'⚠ Διφορούμενα (πάνω από μία γραμμές ταίριαξαν) : {len(results["ambiguous"])}')
    print(f'⚠ Η τιμή στην καρτέλα δεν ήταν «{FROM_DATE_TEXT}» : {len(results["value_mismatch"])}')
    print(f'✗ Σφάλματα εκτέλεσης : {len(results["error"])}')
    for key, label in (
        ('notfound', 'Δεν βρέθηκε ταίριασμα'),
        ('ambiguous', 'Διφορούμενα'),
        ('value_mismatch', f'Τιμή ≠ {FROM_DATE_TEXT}'),
        ('error', 'Σφάλματα'),
    ):
        if results[key]:
            print(f'  {label}:')
            for x in results[key]:
                print(f'    - {x}')
    print('─' * 65)


if __name__ == '__main__':
    main()
