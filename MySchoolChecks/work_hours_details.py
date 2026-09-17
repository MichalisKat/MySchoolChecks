#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
work_hours_details.py
======================
Μόνιμη ενέργεια MySchool — «Λεπτομέρειες ωραρίου».

Καταχωρεί μια νέα εγγραφή στις «Λεπτομέρειες ωραρίου εργασίας» της
καρτέλας Τοποθέτησης, για μια λίστα ατόμων από Excel. Ανοίγεται από το
μενού της εφαρμογής (καρτέλα «Ενέργειες MySchool» → «Λεπτομέρειες
ωραρίου»), το οποίο ζητάει πριν την εκτέλεση:

  1. Περιγραφή (μία από τις δύο, ραδιοπλήκτρα στο παράθυρο):
       α) «Γραμματειακή Υποστήριξη Συμπλήρωση»
       β) «ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π. Συμπλήρωση»
     (Το ίδιο dropdown περιγραφής στο MySchool — cmbWorkHoursDetailsType —
     έχει το κάθε στοιχείο σαν ΕΝΙΑΙΟ κείμενο μαζί με την κατηγορία, οπότε
     αρκεί η ίδια λογική επιλογής με το editor.py, χωρίς δεύτερο combo.)
  2. «Ισχύει από»: πεδίο ημερομηνίας. Αν μείνει ΚΕΝΟ (Enter/τίποτα), το
     script παίρνει για ΤΟΝ ΚΑΘΕΝΑ την τιμή που ΗΔΗ έχει στο πεδίο
     dtDutyStartDate της καρτέλας τοποθέτησής του. Αν δοθεί ημερομηνία,
     αυτή μπαίνει ΣΕ ΟΛΟΥΣ.
  3. «Έως»: ίδια λογική με το dtDutyStopDate.

Οι ΩΡΕΣ της νέας εγγραφής παίρνουν ΠΑΝΤΑ, για τον καθένα, την τιμή που
ήδη έχει στο πεδίο «Διαθέσιμες ώρες μονάδας» (txtAvailableHoursForUnit).

Excel — αναμενόμενες στήλες (ευέλικτη αναγνώριση, ανεξαρτήτως σειράς):
    Α.Μ. και/ή Α.Φ.Μ.  (τουλάχιστον ένα από τα δύο ανά γραμμή)
    ΕΠΩΝΥΜΟ, ΟΝΟΜΑ      (προαιρετικά — μόνο για log)
    Κωδικός Σχολείου και/ή Ονομασία Σχολείου (τουλάχιστον ένα από τα δύο)

Ροή ανά άτομο:
  1. Αναζήτηση στο Worker.list.myEmplUnit.aspx — με ΑΦΜ (txtTaxNumber) αν
     υπάρχει, αλλιώς με Α.Μ. (txtRegistryNo).
  2. Μοναδικό ταίριασμα γραμμής με βάση Ονομασία/Κωδικό Σχολείου (fuzzy
     ονόματος + κυριολεκτικό ταίριασμα κωδικού). 0 ή >1 ταιριάσματα →
     παράλειψη, καταγράφεται (δεν μαντεύουμε ποτέ).
  3. Άνοιγμα καρτέλας, ανάγνωση τρεχουσών τιμών (ώρες μονάδας, ημ. έναρξης/
     λήξης τοποθέτησης).
  4. Κλικ στον σταυρό προσθήκης «Λεπτομέρειες ωραρίου εργασίας».
  5. Επιλογή περιγραφής, συμπλήρωση ωρών, ημερομηνιών.
  6. Αποδοχή → Αποθήκευση.

Τρέχει με ΟΡΑΤΟ Chrome, ένα άτομο τη φορά, με πλήρες log. Ctrl+C /
κλείσιμο παραθύρου οποιαδήποτε στιγμή — ό,τι έχει ήδη αποθηκευτεί
παραμένει, τίποτα παραπέρα δεν αγγίζεται.
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Λεπτομέρειες ωραρίου'
CHECK_DESCRIPTION = 'Καταχώρηση νέας εγγραφής στις Λεπτομέρειες ωραρίου εργασίας τοποθέτησης.'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL     = 'https://app.myschool.sch.gr'
SEARCH_URL   = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

# Οι δύο επιλογές περιγραφής — ΑΚΡΙΒΕΣ κείμενο όπως εμφανίζεται στο
# dropdown του MySchool (cmbWorkHoursDetailsType), κατηγορία
# «συμπλήρωση» ήδη ενσωματωμένη στο κείμενο.
DESCRIPTION_OPTIONS = [
    'Γραμματειακή Υποστήριξη Συμπλήρωση',
    'ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π. Συμπλήρωση',
]

TYPE_COMBO_BASE_ID = 'ctl00_ContentData_gridEmplDet_editnew_2_cmbWorkHoursDetailsType'
ADD_BTN_ID          = 'ctl00_ContentData_gridEmplDet_header0_new'
NEW_ROW_HOURS_ID     = 'ctl00_ContentData_gridEmplDet_DXEditor4_I'
NEW_ROW_DATE_FROM_ID = 'ctl00_ContentData_gridEmplDet_DXEditor5_I'
NEW_ROW_DATE_TO_ID   = 'ctl00_ContentData_gridEmplDet_DXEditor6_I'

CARD_HOURS_ID     = 'ctl00_ContentData_txtAvailableHoursForUnit_I'
CARD_DATE_FROM_ID = 'ctl00_ContentData_dtDutyStartDate_I'
CARD_DATE_TO_ID   = 'ctl00_ContentData_dtDutyStopDate_I'


# ── Ανάγνωση Excel ───────────────────────────────────────────────────────────
def load_people(file_path, log=print):
    """Διαβάζει το excel και επιστρέφει λίστα από dict:
    {'afm','am','eponymo','onoma','school_name','school_code'}."""
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

    afm_col      = _find(('Α.Φ.Μ.', 'ΑΦΜ', 'Α.Φ.Μ'))
    am_col       = _find(('Α.Μ.', 'ΑΜ', 'Α.Μ'))
    epon_col     = _find(('ΕΠΩΝΥΜΟ', 'Επώνυμο'))
    onoma_col    = _find(('ΟΝΟΜΑ', 'Όνομα'))
    sxname_col   = _find(('Ονομασία Σχολείου', 'ΣΧΟΛΕΙΟ', 'Σχολείο'))
    sxcode_col   = _find(('Κωδικός Σχολείου', 'Κωδικός Υπουργείου', 'ΚΩΔ. ΣΧΟΛΕΙΟΥ'))

    if not afm_col and not am_col:
        raise ValueError(f'Λείπει στήλη Α.Φ.Μ. ή Α.Μ. Διαθέσιμες: {list(df.columns)}')
    if not sxname_col and not sxcode_col:
        raise ValueError(f'Λείπει στήλη Ονομασία ή Κωδικός Σχολείου. Διαθέσιμες: {list(df.columns)}')

    people = []
    skipped_no_id = 0
    for _, row in df.iterrows():
        afm_raw = str(row.get(afm_col, '')).strip() if afm_col else ''
        am_raw  = str(row.get(am_col, '')).strip() if am_col else ''
        afm = afm_raw.zfill(9) if afm_raw and afm_raw.lower() not in ('nan', 'none') else ''
        am  = am_raw if am_raw and am_raw.lower() not in ('nan', 'none') else ''
        if not afm and not am:
            skipped_no_id += 1
            continue

        school_name = str(row.get(sxname_col, '')).strip() if sxname_col else ''
        school_code = str(row.get(sxcode_col, '')).strip() if sxcode_col else ''
        if school_name.lower() in ('nan', 'none'):
            school_name = ''
        if school_code.lower() in ('nan', 'none'):
            school_code = ''

        people.append({
            'afm':         afm,
            'am':          am,
            'eponymo':     str(row.get(epon_col, '')).strip() if epon_col else '',
            'onoma':       str(row.get(onoma_col, '')).strip() if onoma_col else '',
            'school_name': school_name,
            'school_code': school_code,
        })

    log(f'  ✓ Διαβάστηκαν {len(people)} εγγραφές από το excel.')
    if skipped_no_id:
        log(f'  ⚠ Παραλείφθηκαν {skipped_no_id} γραμμές χωρίς Α.Φ.Μ. ούτε Α.Μ.')
    return people


def connect(log=print):
    """Άνοιγμα ορατού Chrome + login στο MySchool — ίδιο μοτίβο με τα
    υπόλοιπα εργαλεία του project (editor.py, termination.py κλπ.)."""
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


def _dispatch_keyboard_value(driver, field, value):
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
    """, field, value)


def _search_person(driver, person, log):
    """Ψάχνει με ΑΦΜ (αν υπάρχει) αλλιώς με Α.Μ. Επιστρέφει λίστα edit
    links (Διόρθωση)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(SEARCH_URL)
    time.sleep(2)

    if person['afm']:
        field_id, value = 'ctl00_ContentData_txtTaxNumber_I', person['afm']
        log(f'  Αναζήτηση με ΑΦΜ: {value}')
    else:
        field_id, value = 'ctl00_ContentData_txtRegistryNo_I', person['am']
        log(f'  Αναζήτηση με Α.Μ.: {value}')

    field = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.presence_of_element_located((By.ID, field_id)))
    _dispatch_keyboard_value(driver, field, value)
    time.sleep(1)

    search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
    driver.execute_script('arguments[0].click();', search_link)
    time.sleep(3)

    return driver.find_elements(By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')


def _pick_matching_row(driver, edit_links, school_name_norm, school_code, log):
    """
    Μοναδικό ταίριασμα με βάση Ονομασία Σχολείου (κανονικοποιημένα, fuzzy)
    ή/και Κωδικό Σχολείου (κυριολεκτικό substring στο κείμενο γραμμής).
    Επιστρέφει (link, reason) — reason ∈ {'ok','notfound','ambiguous'}.
    """
    from selenium.webdriver.common.by import By

    rows = []
    for link in edit_links:
        try:
            row = link.find_element(By.XPATH, './ancestor::tr[1]')
            rows.append((link, row.text))
        except Exception:
            continue

    candidates = rows
    if school_name_norm:
        candidates = [(l, t) for (l, t) in candidates
                      if school_name_norm in _normalize_school_name(t)]

    if school_code:
        by_code = [(l, t) for (l, t) in candidates if school_code in t]
        if by_code:
            candidates = by_code
        elif not school_name_norm:
            # Μόνο κωδικός διαθέσιμος και δεν βρέθηκε καθόλου -> ψάξε σε
            # όλα τα rows (όχι μόνο στα ήδη φιλτραρισμένα).
            candidates = [(l, t) for (l, t) in rows if school_code in t]

    if len(candidates) == 1:
        return candidates[0][0], 'ok'
    if len(candidates) == 0:
        return None, 'notfound'
    return None, 'ambiguous'


def _select_dxe_combo(driver, base_id, text, log):
    """Επιλέγει τιμή από DevExpress ComboBox με πληκτρολόγηση + κλικ στο
    ταιριαστό στοιχείο (ίδιος μηχανισμός με το editor.py)."""
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
        for char in text:
            inp.send_keys(char)
            time.sleep(0.05)
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
    except Exception as e:
        log(f'  ⚠ Combo {base_id}: {e}')
        return False


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


def process_person(driver, person, description_text, fixed_date_from, fixed_date_to, log):
    """Επεξεργάζεται μία εγγραφή. Επιστρέφει status string:
    'ok' | 'notfound' | 'ambiguous' | 'error'."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    school_name_norm = _normalize_school_name(person['school_name']) if person['school_name'] else ''
    school_code      = person['school_code']

    try:
        edit_links = _search_person(driver, person, log)
    except Exception as e:
        log(f'  ✗ Αναζήτηση απέτυχε: {e}')
        return 'error'

    if not edit_links:
        log('  ✗ Κανένα αποτέλεσμα αναζήτησης')
        return 'notfound'

    link, reason = _pick_matching_row(driver, edit_links, school_name_norm, school_code, log)
    if reason != 'ok':
        if reason == 'notfound':
            log(f'  ⚠ Δεν βρέθηκε γραμμή με σχολείο «{person["school_name"] or person["school_code"]}» '
                f'({len(edit_links)} αποτέλεσμα(-τα) συνολικά) — παράλειψη')
        else:
            log('  ⚠ Πάνω από μία γραμμές ταιριάζουν — παράλειψη (χειροκίνητος έλεγχος)')
        return reason

    # ── Άνοιγμα καρτέλας ─────────────────────────────────────────────────
    try:
        driver.execute_script('arguments[0].click();', link)
        time.sleep(3)
        log('  Καρτέλα ανοιχτή')
    except Exception as e:
        log(f'  ✗ Άνοιγμα καρτέλας: {e}')
        return 'error'

    # ── Ανάγνωση τρεχουσών τιμών από την καρτέλα ──────────────────────────
    try:
        hours_el = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.ID, CARD_HOURS_ID)))
        person_hours = (hours_el.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ⚠ Διαθέσιμες ώρες μονάδας: {e}')
        person_hours = ''

    if fixed_date_from:
        date_from = fixed_date_from
    else:
        try:
            el = driver.find_element(By.ID, CARD_DATE_FROM_ID)
            date_from = (el.get_attribute('value') or '').strip()
        except Exception as e:
            log(f'  ⚠ Ημ. έναρξης τοποθέτησης (dtDutyStartDate): {e}')
            date_from = ''

    if fixed_date_to:
        date_to = fixed_date_to
    else:
        try:
            el = driver.find_element(By.ID, CARD_DATE_TO_ID)
            date_to = (el.get_attribute('value') or '').strip()
        except Exception as e:
            log(f'  ⚠ Ημ. λήξης τοποθέτησης (dtDutyStopDate): {e}')
            date_to = ''

    log(f'  (ώρες: «{person_hours}», από: «{date_from}», έως: «{date_to}»)')

    if not date_from or not date_to:
        log('  ✗ Λείπει ημερομηνία («από» ή/και «έως») και δεν δόθηκε σταθερή '
            'τιμή στο παράθυρο — παράλειψη (δεν μαντεύουμε ημερομηνία)')
        return 'error'

    # ── Σταυρός προσθήκης ────────────────────────────────────────────────
    try:
        add_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.ID, ADD_BTN_ID)))
        driver.execute_script(
            'arguments[0].scrollIntoView({behavior:"smooth",block:"center"});', add_btn)
        time.sleep(1)
        driver.execute_script('arguments[0].click();', add_btn)
        time.sleep(2)
        log('  Φόρμα ωραρίου ανοιχτή')
    except Exception as e:
        log(f'  ✗ Σταυρός: {e}')
        return 'error'

    # ── Περιγραφή ─────────────────────────────────────────────────────────
    ok_c = _select_dxe_combo(driver, TYPE_COMBO_BASE_ID, description_text, log)
    log(f'  {"✓" if ok_c else "⚠"} Περιγραφή: {description_text}')

    # ── Ώρες νέας εγγραφής ───────────────────────────────────────────────
    if person_hours:
        try:
            hours_inp = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, NEW_ROW_HOURS_ID)))
            driver.execute_script('arguments[0].click(); arguments[0].focus();', hours_inp)
            time.sleep(0.5)
            driver.execute_script("""
                var el = arguments[0];
                var val = arguments[1];
                el.value = '';
                el.value = val;
                el.dispatchEvent(new Event('change', {bubbles: true}));
                aspxEValueChanged('ctl00_ContentData_gridEmplDet_DXEditor4');
            """, hours_inp, person_hours)
            log(f'  ✓ Ώρες νέας εγγραφής: {person_hours}')
        except Exception as e:
            log(f'  ⚠ Ώρες φόρμα: {e}')
    else:
        log('  ⚠ Καμία τιμή ωρών να συμπληρωθεί στη νέα εγγραφή')

    # ── Ημερομηνίες ───────────────────────────────────────────────────────
    try:
        _set_dxe_value(driver, NEW_ROW_DATE_FROM_ID, date_from)
        log(f'  ✓ Ημ. από: {date_from}')
    except Exception as e:
        log(f'  ⚠ Ημ. από: {e}')
    try:
        _set_dxe_value(driver, NEW_ROW_DATE_TO_ID, date_to)
        log(f'  ✓ Ημ. έως: {date_to}')
    except Exception as e:
        log(f'  ⚠ Ημ. έως: {e}')

    time.sleep(0.5)

    # ── Αποδοχή ───────────────────────────────────────────────────────────
    try:
        accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.XPATH, '//img[@alt="Αποδοχή"]')))
        driver.execute_script('arguments[0].click();', accept_btn)
        time.sleep(2)
        log('  ✓ Αποδοχή')
    except Exception as e:
        log(f'  ✗ Αποδοχή: {e}')
        return 'error'

    # ── Αποθήκευση ────────────────────────────────────────────────────────
    try:
        save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_ContentData_btnSave')))
        driver.execute_script('arguments[0].click();', save_btn)
        time.sleep(3)
        log('  ✓ Αποθήκευση')
        return 'ok'
    except Exception as e:
        log(f'  ✗ Αποθήκευση: {e}')
        return 'error'


def run(ctx, driver, callback=None):
    """
    ctx αναμενόμενα κλειδιά:
      'file_path'   — απαραίτητο
      'description' — μία από τις DESCRIPTION_OPTIONS
      'date_from'   — κενό ή ημερομηνία (ΗΗ/Μ/ΕΕΕΕ) — αν κενό, ανά άτομο
      'date_to'     — κενό ή ημερομηνία (ΗΗ/Μ/ΕΕΕΕ) — αν κενό, ανά άτομο
    """
    log = callback or print

    file_path        = ctx.get('file_path')
    description_text = (ctx.get('description') or '').strip()
    fixed_date_from   = (ctx.get('date_from') or '').strip()
    fixed_date_to     = (ctx.get('date_to') or '').strip()

    if not description_text:
        log('✗ Δεν επιλέχθηκε περιγραφή — τέλος.')
        return

    people = load_people(file_path, log=log)
    if not people:
        return

    total = len(people)
    log(f'\n  {total} εγγραφές  |  Περιγραφή: {description_text}')
    log(f'  Από: {fixed_date_from or "(ανά άτομο, από την καρτέλα)"}  |  '
        f'Έως: {fixed_date_to or "(ανά άτομο, από την καρτέλα)"}')

    results = {'ok': [], 'notfound': [], 'ambiguous': [], 'error': []}

    try:
        for idx, person in enumerate(people, 1):
            ident = person['afm'] or person['am']
            log(f'\n[{idx}/{total}] {ident}  {person["eponymo"]} {person["onoma"]}  '
                f'—  {person["school_name"] or person["school_code"]}')
            status = process_person(driver, person, description_text,
                                     fixed_date_from, fixed_date_to, log)
            results[status].append(ident)
            time.sleep(0.5)
    except KeyboardInterrupt:
        log('\n\n⚠ Διακόπηκε από τον χρήστη. Ό,τι έχει ήδη αποθηκευτεί παραμένει.')

    log('\n' + '─' * 65)
    log(f'✓ Καταχωρήθηκαν επιτυχώς : {len(results["ok"])}')
    log(f'⚠ Δεν βρέθηκε ταίριασμα σχολείου : {len(results["notfound"])}')
    log(f'⚠ Διφορούμενα (πάνω από μία γραμμές ταίριαξαν) : {len(results["ambiguous"])}')
    log(f'✗ Σφάλματα εκτέλεσης : {len(results["error"])}')
    for key, label in (
        ('notfound', 'Δεν βρέθηκε ταίριασμα'),
        ('ambiguous', 'Διφορούμενα'),
        ('error', 'Σφάλματα'),
    ):
        if results[key]:
            log(f'  {label}: ' + ' | '.join(results[key]))
    log('─' * 65)
