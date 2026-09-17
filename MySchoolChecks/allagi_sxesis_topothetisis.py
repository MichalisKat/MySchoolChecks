#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
allagi_sxesis_topothetisis.py
==============================
One-off εργασία (ΔΕΝ μπαίνει στο μενού της εφαρμογής) — αλλαγή Σχέσης
Τοποθέτησης από «Οργανικά» σε «Οργανικά σε Τμήμα Ένταξης», για μια λίστα
ατόμων από Excel.

Βάλε αυτό το αρχείο μέσα στον φάκελο MySchoolChecks\ (δίπλα στο main.py,
config.py κλπ.) και τρέξε το με:

    python allagi_sxesis_topothetisis.py "20260911  Οργανικά σε ΤΕ.xlsx"

Excel — αναμενόμενες στήλες (ακριβώς όπως στο αρχείο που δόθηκε):
    Α.Μ. | Επώνυμο | Όνομα | Κωδ. Ειδικότητας |
    Μονάδα Οργανικής/Προσωρινής Τοποθέτησης | Ονομασία

Ροή ανά άτομο:
  1. Αναζήτηση στο https://app.myschool.sch.gr/Worker.list.myEmplUnit.aspx
     με βάση το Α.Μ.
  2. Στα αποτελέσματα (μπορεί να βγουν πάνω από μία γραμμές — π.χ. αν το
     άτομο έχει και δεύτερη τοποθέτηση) εντοπίζεται η ΜΟΝΑΔΙΚΗ γραμμή που
     έχει ταυτόχρονα:
       • Σχέση Τοποθέτησης = «Οργανικά» (ακριβώς — όχι ήδη «Οργανικά σε
         Τμήμα Ένταξης», ώστε να μην ξαναπειράξουμε κάτι που έχει ήδη
         αλλάξει)
       • Σχολείο = αυτό της στήλης «Ονομασία» του excel (fuzzy match, ίδια
         λογική κανονικοποίησης με το placements.py — ανεκτικό σε τόνους/
         συντομογραφίες).
     Αν βρεθούν 0 ή >1 τέτοιες γραμμές → η εγγραφή ΑΓΝΟΕΙΤΑΙ και καταγράφεται
     (καμία αλλαγή) — δεν μαντεύουμε ποτέ.
  3. Άνοιγμα της καρτέλας (κλικ στο «Διόρθωση» — το γρανάζι/μολύβι).
  4. Ασφαλής επανέλεγχος: το πεδίο cmbEmploymentType στην ίδια την καρτέλα
     πρέπει να δείχνει ήδη «Οργανικά» — αν όχι, ΔΕΝ αγγίζουμε τίποτα.
  5. Αλλαγή του πεδίου σε «Οργανικά σε Τμήμα Ένταξης».
  6. Κλικ «Αποθήκευση».

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

FROM_TEXT   = 'Οργανικά'
TO_TEXT     = 'Οργανικά σε Τμήμα Ένταξης'
FILTER_TEXT = 'Τμήμα Ένταξης'   # μοναδικό κομμάτι του TO_TEXT — βλ. σχόλιο σε _select_employment_type

COMBO_BASE = 'ctl00_ContentData_cmbEmploymentType'
SAVE_BTN_ID = 'ctl00_ContentData_btnSave'

STRIKE_INTERVAL = 0.3


# ── Ανάγνωση Excel ───────────────────────────────────────────────────────────
def load_people(file_path, log=print):
    """Διαβάζει το excel και επιστρέφει λίστα από dict
    {'am','eponymo','onoma','sxoleio'} — ένα ανά άτομο."""
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
    epon_col    = _find(('Επώνυμο',))
    onoma_col   = _find(('Όνομα',))
    sxoleio_col = _find(('Ονομασία',))

    missing = [n for n, v in (('Α.Μ.', am_col), ('Ονομασία', sxoleio_col)) if v is None]
    if missing:
        raise ValueError(f'Λείπουν στήλες: {", ".join(missing)}. Διαθέσιμες: {list(df.columns)}')

    people = []
    for _, row in df.iterrows():
        am = str(row.get(am_col, '')).strip()
        if not am or am.lower() in ('nan', 'none', ''):
            continue
        people.append({
            'am':      am,
            'eponymo': str(row.get(epon_col, '')).strip() if epon_col else '',
            'onoma':   str(row.get(onoma_col, '')).strip() if onoma_col else '',
            'sxoleio': str(row.get(sxoleio_col, '')).strip() if sxoleio_col else '',
        })
    log(f'  ✓ Διαβάστηκαν {len(people)} άτομα από το excel.')
    return people


# ── Selenium βοηθητικά (ίδιο μοτίβο με dioikitiko_ergo_entry.py) ────────────
def _send_keys_slow(element, text, delay=STRIKE_INTERVAL):
    for char in str(text):
        element.send_keys(char)
        time.sleep(delay)


def _select_employment_type(driver, base_id, exact_text, filter_text, log):
    """
    Επιλέγει την ακριβή τιμή σε ASPxComboBox με πληκτρολόγηση — ίδιο μοτίβο
    με το _select_dxe_combo του editor.py (Ε5) / dioikitiko_ergo_entry.py (Ε8),
    που δουλεύει αξιόπιστα εκεί.

    ΣΗΜΑΝΤΙΚΗ ΔΙΑΦΟΡΑ / FIX: ΔΕΝ πληκτρολογούμε το πλήρες exact_text — το
    «Οργανικά» είναι ΠΡΟΘΕΜΑ του «Οργανικά σε Τμήμα Ένταξης», οπότε καθώς
    πληκτρολογούσαμε γράμμα-γράμμα, μόλις γραφόταν "ΟΡΓΑΝΙΚΑ" το autocomplete
    έβρισκε ήδη ακριβές match με την ΥΠΑΡΧΟΥΣΑ τιμή «Οργανικά» και κολλούσε
    εκεί — το κενό/συνέχεια μετά δεν φιλτράριζε παραπέρα σωστά. Γι' αυτό
    πληκτρολογούμε αντ' αυτού ένα ΜΟΝΑΔΙΚΟ κομμάτι (filter_text, π.χ. «Τμήμα
    Ένταξης») που ΔΕΝ εμφανίζεται καθόλου στο «Οργανικά» — έτσι το φιλτράρισμα
    πάει κατευθείαν στη σωστή (και μοναδική) επιλογή, χωρίς ποτέ να περάσει
    από ψευδές exact-match στο μικρότερο.

    Επιστρέφει True/False.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys

    inp = driver.find_element(By.ID, base_id + '_I')
    driver.execute_script('arguments[0].click();', inp)
    time.sleep(0.5)

    # ── Καθάρισμα πεδίου — δοκιμάζουμε πολλαπλούς τρόπους μέχρι να αδειάσει
    #    όντως (το απλό inp.clear() δεν αρκούσε σε αυτό το combo) ───────────
    inp.click()
    inp.send_keys(Keys.CONTROL, 'a')
    inp.send_keys(Keys.DELETE)
    time.sleep(0.3)
    if (inp.get_attribute('value') or '').strip():
        inp.clear()
        time.sleep(0.2)
    if (inp.get_attribute('value') or '').strip():
        # Backspace επαναλαμβανόμενα (μήκος τρέχουσας τιμής + περιθώριο)
        cur_len = len(inp.get_attribute('value') or '')
        for _ in range(cur_len + 5):
            inp.send_keys(Keys.BACKSPACE)
        time.sleep(0.3)
    if (inp.get_attribute('value') or '').strip():
        driver.execute_script("""
            var el = arguments[0];
            el.value = '';
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        """, inp)
        time.sleep(0.3)

    log(f'    (μετά το καθάρισμα: τιμή στο πεδίο = "{inp.get_attribute("value")}")')

    _send_keys_slow(inp, filter_text)
    time.sleep(1.5)
    log(f'    (μετά την πληκτρολόγηση «{filter_text}»: τιμή στο πεδίο = "{inp.get_attribute("value")}")')

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'td[dxtext]')))
    except Exception:
        log('    ⚠ Δεν εμφανίστηκε λίστα προτάσεων')

    def _norm(s):
        return ' '.join(str(s or '').split()).strip().upper()

    items = driver.find_elements(By.CSS_SELECTOR, 'td[dxtext]')
    seen = [(it.get_attribute('dxtext') or it.text) for it in items]
    log(f'    Επιλογές στη λίστα ({len(seen)}): {seen}')

    target_norm = _norm(exact_text)
    match = None
    for it, txt in zip(items, seen):
        if _norm(txt) == target_norm:
            match = it
            break

    if match is None:
        log(f'    ✗ Καμία επιλογή δεν ταιριάζει ακριβώς με «{exact_text}» — καμία αλλαγή')
        return False

    driver.execute_script('arguments[0].scrollIntoView({block:"center"});', match)
    time.sleep(0.2)

    # ── Κλικ στο item ─────────────────────────────────────────────────────
    # ΣΗΜΑΝΤΙΚΟ: το DevExpress dropdown "ακούει" mousedown (όχι μόνο click)
    # για να προλάβει το blur του πεδίου κειμένου. Ένα JS-based
    # arguments[0].click() πυροδοτεί ΜΟΝΟ synthetic 'click' event — όχι
    # mousedown/mouseup — οπότε ο εσωτερικός χειριστής επιλογής της
    # DevExpress ποτέ δεν ενεργοποιείται: το dropdown απλώς κλείνει και
    # μένει ό,τι είχε πληκτρολογηθεί. Γι' αυτό χρησιμοποιούμε πραγματικό
    # WebDriver click (πυροδοτεί πραγματικά mousedown/mouseup/click) με
    # ActionChains ως fallback αν παρεμποδιστεί.
    from selenium.webdriver.common.action_chains import ActionChains
    clicked_ok = False
    for attempt in range(3):
        try:
            match.click()
            clicked_ok = True
        except Exception as e:
            try:
                ActionChains(driver).move_to_element(match).click().perform()
                clicked_ok = True
            except Exception as e2:
                log(f'    ⚠ Απέτυχε το κλικ (προσπάθεια {attempt+1}): {e2}')
        time.sleep(0.8)
        final_val = (inp.get_attribute('value') or '').strip()
        if final_val == exact_text:
            break
        # ξαναβρίσκουμε το item σε περίπτωση που το DOM ανανεώθηκε
        try:
            items2 = driver.find_elements(By.CSS_SELECTOR, 'td[dxtext]')
            for it2 in items2:
                if _norm(it2.get_attribute('dxtext') or it2.text) == target_norm:
                    match = it2
                    break
        except Exception:
            pass

    final_val = (inp.get_attribute('value') or '').strip()
    log(f'    (μετά το κλικ επιλογής: τιμή στο πεδίο = "{final_val}")')
    if not clicked_ok or final_val != exact_text:
        log(f'    ✗ Μετά το κλικ η τιμή είναι «{final_val}» αντί «{exact_text}» — καμία αλλαγή')
        return False
    return True


def connect(log=print):
    """Άνοιγμα ορατού Chrome + login στο MySchool — ίδιο μοτίβο με
    dioikitiko_ergo_entry.connect()."""
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
    Ανάμεσα στα edit_links, βρίσκει τη ΜΟΝΑΔΙΚΗ γραμμή που έχει:
      • Σχέση Τοποθέτησης ακριβώς «Οργανικά» (όχι «Οργανικά σε Τμήμα
        Ένταξης» — άτομο που έχει ήδη αλλάξει)
      • σχολείο που ταιριάζει (κανονικοποιημένα) με το target_school_norm
    Επιστρέφει (link, reason) — reason ∈ {'ok','notfound','ambiguous'}.
    """
    from selenium.webdriver.common.by import By

    candidates = []
    for link in edit_links:
        try:
            row = link.find_element(By.XPATH, './ancestor::tr[1]')
            row_text = row.text
        except Exception:
            continue

        # Πρέπει να αναφέρει "Οργανικά" αλλά ΟΧΙ ήδη "Τμήμα Ένταξης"
        if FROM_TEXT not in row_text or 'ΤΜΗΜΑ ΕΝΤΑΞΗΣ' in row_text.upper():
            continue

        row_norm = _normalize_school_name(row_text)
        if target_school_norm and target_school_norm in row_norm:
            candidates.append(link)

    if len(candidates) == 1:
        return candidates[0], 'ok'
    if len(candidates) == 0:
        return None, 'notfound'
    return None, 'ambiguous'


def process_person(driver, person, log):
    """Επεξεργάζεται ένα άτομο. Επιστρέφει status string:
    'ok' | 'notfound' | 'ambiguous' | 'value_mismatch' | 'error'."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    am      = person['am']
    sxoleio = person['sxoleio']
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
            log(f'  ⚠ Δεν βρέθηκε γραμμή με «{FROM_TEXT}» + σχολείο «{sxoleio}» '
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
        combo_inp = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.ID, COMBO_BASE + '_I')))
        current_val = (combo_inp.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ✗ Δεν βρέθηκε το πεδίο Σχέση Τοποθέτησης στην καρτέλα: {e}')
        return 'error'

    if current_val != FROM_TEXT:
        log(f'  ⚠ Η τιμή στην καρτέλα είναι «{current_val}» (όχι «{FROM_TEXT}») — παράλειψη, καμία αλλαγή')
        return 'value_mismatch'

    # ── Αλλαγή τιμής (πληκτρολόγηση, με σωστό καθάρισμα + αναμονή dropdown) ─
    ok_c = _select_employment_type(driver, COMBO_BASE, TO_TEXT, FILTER_TEXT, log)
    if not ok_c:
        log('  ✗ Αποτυχία επιλογής νέας τιμής στο dropdown — καμία αλλαγή')
        return 'error'

    # ── Επιβεβαίωση ΠΡΙΝ την αποθήκευση (το DOM πρέπει ήδη να δείχνει τη νέα τιμή) ──
    try:
        combo_inp2 = driver.find_element(By.ID, COMBO_BASE + '_I')
        after_select_val = (combo_inp2.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ✗ Δεν διαβάστηκε η τιμή μετά την επιλογή: {e}')
        return 'error'

    if after_select_val != TO_TEXT:
        log(f'  ✗ Μετά την επιλογή το πεδίο δείχνει «{after_select_val}» αντί «{TO_TEXT}» — ΔΕΝ αποθηκεύεται')
        return 'error'
    log(f'  ✓ Επιλέχθηκε «{TO_TEXT}» (επιβεβαιώθηκε στη φόρμα)')

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
        title='Επίλεξε το αρχείο Excel (Α.Μ., Επώνυμο, Όνομα, ..., Ονομασία)',
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
    print('  Αλλαγή Σχέσης Τοποθέτησης — Οργανικά → Οργανικά σε Τμήμα Ένταξης')
    print('=' * 65)

    people = load_people(file_path)
    if limit:
        people = people[:limit]
        print(f'  (δοκιμαστικό τρέξιμο — μόνο τα πρώτα {limit} άτομα)')

    driver = connect()
    if driver is None:
        print('✗ Αποτυχία σύνδεσης — τέλος.')
        sys.exit(1)

    results = {'ok': [], 'notfound': [], 'ambiguous': [], 'value_mismatch': [], 'error': []}

    total = len(people)
    try:
        for idx, person in enumerate(people, 1):
            print(f'\n[{idx}/{total}] Α.Μ. {person["am"]}  {person["eponymo"]} {person["onoma"]}  '
                  f'—  {person["sxoleio"]}')
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
    print(f'✓ Αλλάχθηκαν επιτυχώς : {len(results["ok"])}')
    print(f'⚠ Δεν βρέθηκε ταίριασμα (Οργανικά + σχολείο) : {len(results["notfound"])}')
    print(f'⚠ Διφορούμενα (πάνω από μία γραμμές ταίριαξαν) : {len(results["ambiguous"])}')
    print(f'⚠ Η τιμή στην καρτέλα δεν ήταν «Οργανικά» : {len(results["value_mismatch"])}')
    print(f'✗ Σφάλματα εκτέλεσης : {len(results["error"])}')
    for key, label in (
        ('notfound', 'Δεν βρέθηκε ταίριασμα'),
        ('ambiguous', 'Διφορούμενα'),
        ('value_mismatch', 'Τιμή ≠ Οργανικά'),
        ('error', 'Σφάλματα'),
    ):
        if results[key]:
            print(f'  {label} — Α.Μ.: ' + ' | '.join(results[key]))
    print('─' * 65)


if __name__ == '__main__':
    main()
