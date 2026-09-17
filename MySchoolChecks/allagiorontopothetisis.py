#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
allagiorontopothetisis.py
==========================
One-off εργασία (ΔΕΝ μπαίνει στο μενού της εφαρμογής) — ενημέρωση του
πεδίου «Διαθέσιμες ώρες μονάδας» (txtAvailableHoursForUnit) στην καρτέλα
Οργανικής/Προσωρινής Τοποθέτησης του MySchool, για μια λίστα ατόμων από
Excel (το φύλλο «ΣΧΟΛΕΙΟ ΤΟΠΟΘΕΤΗΣΗΣ» που παράγει το specialties pipeline).

Βάλε αυτό το αρχείο μέσα στον φάκελο MySchoolChecks\ (δίπλα στο main.py,
config.py, allagi_sxesis_topothetisis.py κλπ.) και τρέξε το με:

    python allagiorontopothetisis.py "ΣΧΟΛΕΙΟ ΤΟΠΟΘΕΤΗΣΗΣ.xlsx"

(Χωρίς όρισμα, ανοίγει παράθυρο επιλογής αρχείου. Προαιρετικό 2ο όρισμα =
αριθμός ατόμων για δοκιμαστικό τρέξιμο, π.χ. `... 5`.)

Excel — αναμενόμενες στήλες (όπως παράγονται από το φύλλο «ΣΧΟΛΕΙΟ
ΤΟΠΟΘΕΤΗΣΗΣ (ΥΠΟΛΟΙΠΟ)»):
    Α.Μ. | ΕΠΩΝΥΜΟ | ΟΝΟΜΑ | Α.Φ.Μ. | ΚΛΑΔΟΣ | Κωδικός Υπουργείου | ΣΧΟΛΕΙΟ | ΩΡΕΣ

Ροή ανά άτομο:
  1. Αναζήτηση στο https://app.myschool.sch.gr/Worker.list.myEmplUnit.aspx
     με βάση το Α.Μ. (ίδια σελίδα/μηχανισμός με το allagi_sxesis_topothetisis.py).
  2. Στα αποτελέσματα (μπορεί να βγουν πάνω από μία γραμμές — π.χ. οργανική +
     προσωρινή τοποθέτηση) εντοπίζεται η ΜΟΝΑΔΙΚΗ γραμμή που αναφέρει το
     σχολείο της στήλης «ΣΧΟΛΕΙΟ» του excel (fuzzy match, ίδια λογική
     κανονικοποίησης με το placements.py — ανεκτικό σε τόνους/συντομογραφίες).
     Αν βρεθούν 0 ή >1 τέτοιες γραμμές → η εγγραφή ΑΓΝΟΕΙΤΑΙ και καταγράφεται
     (καμία αλλαγή) — δεν μαντεύουμε ποτέ.
  3. Άνοιγμα της καρτέλας (κλικ στο «Διόρθωση» — το γρανάζι/μολύβι).
  4. Ανάγνωση της τρέχουσας τιμής στο txtAvailableHoursForUnit. Αν είναι ήδη
     ίση με τη νέα τιμή → καμία αλλαγή/αποθήκευση (no-op, καταγράφεται ως
     «ήδη σωστό»).
  5. Αλλιώς: καθαρισμός πεδίου, πληκτρολόγηση της νέας τιμής (στήλη ΩΡΕΣ).
  6. Επανέλεγχος ότι το πεδίο δείχνει πράγματι τη νέα τιμή, και μόνο τότε
     κλικ «Αποθήκευση».

Τρέχει με ΟΡΑΤΟ Chrome, ένα άτομο τη φορά, με πλήρες log στην κονσόλα —
μπορείς να το παρακολουθείς live. Αν κάτι δείχνει λάθος, κλείσε το
παράθυρο της κονσόλας (Ctrl+C) οποιαδήποτε στιγμή· ό,τι έχει ήδη
αποθηκευτεί μέχρι εκείνο το σημείο παραμένει (δεν γίνεται rollback), αλλά
τίποτα παραπέρα δεν θα αγγιχτεί.
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE_URL     = 'https://app.myschool.sch.gr'
SEARCH_URL   = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

HOURS_FIELD_ID = 'ctl00_ContentData_txtAvailableHoursForUnit_I'
SAVE_BTN_ID    = 'ctl00_ContentData_btnSave'


# ── Ανάγνωση Excel ───────────────────────────────────────────────────────────
def load_people(file_path, log=print):
    """Διαβάζει το excel και επιστρέφει λίστα από dict
    {'am','eponymo','onoma','sxoleio','ores'} — ένα ανά γραμμή."""
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

    am_col      = _find(('Α.Μ.', 'ΑΜ', 'Α.Μ'))
    epon_col    = _find(('ΕΠΩΝΥΜΟ', 'Επώνυμο'))
    onoma_col   = _find(('ΟΝΟΜΑ', 'Όνομα'))
    sxoleio_col = _find(('ΣΧΟΛΕΙΟ', 'Ονομασία'))
    ores_col    = _find(('ΩΡΕΣ', 'Ώρες'))

    missing = [n for n, v in (
        ('Α.Μ.', am_col), ('ΣΧΟΛΕΙΟ', sxoleio_col), ('ΩΡΕΣ', ores_col),
    ) if v is None]
    if missing:
        raise ValueError(f'Λείπουν στήλες: {", ".join(missing)}. Διαθέσιμες: {list(df.columns)}')

    people = []
    for _, row in df.iterrows():
        am = str(row.get(am_col, '')).strip()
        ores_raw = str(row.get(ores_col, '')).strip()
        if not am or am.lower() in ('nan', 'none', ''):
            continue
        if not ores_raw or ores_raw.lower() in ('nan', 'none', ''):
            continue
        try:
            ores = int(float(ores_raw))
        except ValueError:
            log(f'  ⚠ Μη αριθμητική τιμή ωρών «{ores_raw}» για Α.Μ. {am} — παράλειψη γραμμής')
            continue
        people.append({
            'am':      am,
            'eponymo': str(row.get(epon_col, '')).strip() if epon_col else '',
            'onoma':   str(row.get(onoma_col, '')).strip() if onoma_col else '',
            'sxoleio': str(row.get(sxoleio_col, '')).strip() if sxoleio_col else '',
            'ores':    ores,
        })
    log(f'  ✓ Διαβάστηκαν {len(people)} εγγραφές από το excel.')
    return people


def connect(log=print):
    """Άνοιγμα ορατού Chrome + login στο MySchool — ίδιο μοτίβο με
    allagi_sxesis_topothetisis.connect()."""
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


def _search_by_am(driver, am, log):
    """Πάει στη σελίδα αναζήτησης, συμπληρώνει Α.Μ., πατάει αναζήτηση.
    Επιστρέφει λίστα <a> edit links (Διόρθωση)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(SEARCH_URL)
    time.sleep(2)

    am_field = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.presence_of_element_located(
            (By.ID, 'ctl00_ContentData_txtRegistryNo_I')))
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
    """, am_field, am)
    time.sleep(1)

    search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
    driver.execute_script('arguments[0].click();', search_link)
    time.sleep(3)

    return driver.find_elements(By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')


def _pick_matching_row(driver, edit_links, target_school_norm, log):
    """
    Ανάμεσα στα edit_links, βρίσκει τη ΜΟΝΑΔΙΚΗ γραμμή που αναφέρει το
    σχολείο-στόχο (κανονικοποιημένα). Επιστρέφει (link, reason) —
    reason ∈ {'ok','notfound','ambiguous'}.
    """
    from selenium.webdriver.common.by import By

    candidates = []
    for link in edit_links:
        try:
            row = link.find_element(By.XPATH, './ancestor::tr[1]')
            row_text = row.text
        except Exception:
            continue

        row_norm = _normalize_school_name(row_text)
        if target_school_norm and target_school_norm in row_norm:
            candidates.append(link)

    if len(candidates) == 1:
        return candidates[0], 'ok'
    if len(candidates) == 0:
        return None, 'notfound'
    return None, 'ambiguous'


def _set_hours_field(driver, new_value, log):
    """Καθαρίζει και συμπληρώνει το txtAvailableHoursForUnit με τη νέα τιμή.
    Επιστρέφει True/False ανάλογα με το αν η τιμή στο DOM επιβεβαιώθηκε."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    inp = driver.find_element(By.ID, HOURS_FIELD_ID)
    driver.execute_script('arguments[0].click();', inp)
    time.sleep(0.3)

    inp.send_keys(Keys.CONTROL, 'a')
    inp.send_keys(Keys.DELETE)
    time.sleep(0.2)
    if (inp.get_attribute('value') or '').strip():
        driver.execute_script("""
            var el = arguments[0];
            el.value = '';
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        """, inp)
        time.sleep(0.2)

    inp.send_keys(str(new_value))
    time.sleep(0.3)
    inp.send_keys(Keys.TAB)   # πυροδοτεί onchange/onblur (aspxELostFocus)
    time.sleep(0.5)

    final_val = (inp.get_attribute('value') or '').strip()
    return final_val == str(new_value)


def process_person(driver, person, log):
    """Επεξεργάζεται μία εγγραφή. Επιστρέφει status string:
    'ok' | 'already_ok' | 'notfound' | 'ambiguous' | 'error'."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    am      = person['am']
    sxoleio = person['sxoleio']
    ores    = person['ores']
    target_school_norm = _normalize_school_name(sxoleio)

    try:
        edit_links = _search_by_am(driver, am, log)
    except Exception as e:
        log(f'  ✗ Αναζήτηση απέτυχε: {e}')
        return 'error'

    if not edit_links:
        log('  ✗ Κανένα αποτέλεσμα αναζήτησης')
        return 'notfound'

    link, reason = _pick_matching_row(driver, edit_links, target_school_norm, log)
    if reason != 'ok':
        if reason == 'notfound':
            log(f'  ⚠ Δεν βρέθηκε γραμμή με σχολείο «{sxoleio}» '
                f'({len(edit_links)} αποτέλεσμα(-τα) συνολικά) — παράλειψη')
        else:
            log('  ⚠ Πάνω από μία γραμμές ταιριάζουν — παράλειψη (χειροκίνητος έλεγχος)')
        return reason

    # ── Άνοιγμα καρτέλας ─────────────────────────────────────────────────
    try:
        driver.execute_script('arguments[0].click();', link)
        time.sleep(2.5)
    except Exception as e:
        log(f'  ✗ Άνοιγμα καρτέλας: {e}')
        return 'error'

    # ── Ανάγνωση τρέχουσας τιμής ────────────────────────────────────────
    try:
        hours_inp = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.ID, HOURS_FIELD_ID)))
        current_val = (hours_inp.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ✗ Δεν βρέθηκε το πεδίο Διαθέσιμες ώρες μονάδας στην καρτέλα: {e}')
        return 'error'

    log(f'  (τρέχουσα τιμή: «{current_val}», νέα τιμή: «{ores}»)')

    if current_val == str(ores):
        log('  ✓ Η τιμή είναι ήδη σωστή — καμία αλλαγή')
        return 'already_ok'

    # ── Αλλαγή τιμής ─────────────────────────────────────────────────────
    ok_set = _set_hours_field(driver, ores, log)
    if not ok_set:
        log('  ✗ Αποτυχία συμπλήρωσης νέας τιμής — καμία αποθήκευση')
        return 'error'
    log(f'  ✓ Συμπληρώθηκε «{ores}» (επιβεβαιώθηκε στη φόρμα)')

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
        title='Επίλεξε το αρχείο Excel (Α.Μ., ΕΠΩΝΥΜΟ, ΟΝΟΜΑ, ..., ΣΧΟΛΕΙΟ, ΩΡΕΣ)',
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
    print('  Αλλαγή Ωρών Τοποθέτησης (txtAvailableHoursForUnit)')
    print('=' * 65)

    people = load_people(file_path)
    if limit:
        people = people[:limit]
        print(f'  (δοκιμαστικό τρέξιμο — μόνο οι πρώτες {limit} εγγραφές)')

    driver = connect()
    if driver is None:
        print('✗ Αποτυχία σύνδεσης — τέλος.')
        sys.exit(1)

    results = {'ok': [], 'already_ok': [], 'notfound': [], 'ambiguous': [], 'error': []}

    total = len(people)
    try:
        for idx, person in enumerate(people, 1):
            print(f'\n[{idx}/{total}] Α.Μ. {person["am"]}  {person["eponymo"]} {person["onoma"]}  '
                  f'—  {person["sxoleio"]}  ({person["ores"]} ώρες)')
            status = process_person(driver, person, print)
            results[status].append(person['am'])
            time.sleep(0.5)
    except KeyboardInterrupt:
        print('\n\n⚠ Διακόπηκε από τον χρήστη (Ctrl+C). Ό,τι έχει ήδη αποθηκευτεί παραμένει.')
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    print('\n' + '─' * 65)
    print(f'✓ Ενημερώθηκαν επιτυχώς : {len(results["ok"])}')
    print(f'= Ήταν ήδη σωστές (καμία αλλαγή) : {len(results["already_ok"])}')
    print(f'⚠ Δεν βρέθηκε ταίριασμα σχολείου : {len(results["notfound"])}')
    print(f'⚠ Διφορούμενα (πάνω από μία γραμμές ταίριαξαν) : {len(results["ambiguous"])}')
    print(f'✗ Σφάλματα εκτέλεσης : {len(results["error"])}')
    for key, label in (
        ('notfound', 'Δεν βρέθηκε ταίριασμα'),
        ('ambiguous', 'Διφορούμενα'),
        ('error', 'Σφάλματα'),
    ):
        if results[key]:
            print(f'  {label} — Α.Μ.: ' + ' | '.join(results[key]))
    print('─' * 65)


if __name__ == '__main__':
    main()
