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
       α) «Γραμματειακή Υποστήριξη»
       β) «ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π.»
     Η επιλογή από το dropdown περιγραφής (cmbWorkHoursDetailsType) γίνεται
     με την ΙΔΙΑ απλή μέθοδο του editor.py (Ε5) — πληκτρολόγηση αργά +
     κλικ στο πρώτο td[dxtext=κείμενο] + fallback TAB — δοκιμασμένη και
     αξιόπιστη εκεί. Για την ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ το dropdown δείχνει 2 γραμμές
     με το ίδιο κείμενο (Συμπλήρωση/Υπερωρία, με τη Συμπλήρωση πάντα πρώτη
     στην πράξη) — συνειδητή επιλογή απλότητας αντί για πολύπλοκη ανά-γραμμή
     αντιστοίχιση κατηγορίας.
  2. «Ισχύει από»: πεδίο ημερομηνίας, προσυμπληρωμένο με τη ΣΗΜΕΡΙΝΗ
     ημερομηνία (όπως το Ε5). Αν το αδειάσεις (Enter/τίποτα), το script
     παίρνει για ΤΟΝ ΚΑΘΕΝΑ την τιμή που ΗΔΗ έχει στο πεδίο dtDutyStartDate
     της καρτέλας τοποθέτησής του.
  3. «Έως»: ίδια λογική με το dtDutyStopDate (κενό = ανά άτομο).

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

# Οι δύο επιλογές περιγραφής — ΑΚΡΙΒΕΣ κείμενο της στήλης «Περιγραφή» στο
# dropdown του combo περιγραφής (cmbWorkHoursDetailsType) — ίδιο μοτίβο με
# το WORK_TYPE_TEXT = 'Γραμματειακή Υποστήριξη' του editor.py (Ε5). Η
# επιλογή γίνεται με την ΙΔΙΑ απλή μέθοδο του Ε5 (_select_work_type_combo):
# πρώτο ταίριασμα, χωρίς διάκριση κατηγορίας.
DESCRIPTION_OPTIONS = [
    'Γραμματειακή Υποστήριξη',
    'ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π.',
]

# Η κατηγορία που εμφανίζεται συνήθως ΠΡΩΤΗ στο dropdown (και άρα
# επιλέγεται από την απλή μέθοδο) — χρησιμοποιείται μόνο για τον
# προαιρετικό ground-truth έλεγχο μετά την αποθήκευση.
CATEGORY_TEXT = 'Συμπλήρωση'

TYPE_COMBO_BASE_ID = 'ctl00_ContentData_gridEmplDet_editnew_2_cmbWorkHoursDetailsType'
ADD_BTN_ID          = 'ctl00_ContentData_gridEmplDet_header0_new'
NEW_ROW_HOURS_ID     = 'ctl00_ContentData_gridEmplDet_DXEditor4_I'
NEW_ROW_DATE_FROM_ID = 'ctl00_ContentData_gridEmplDet_DXEditor5_I'
NEW_ROW_DATE_TO_ID   = 'ctl00_ContentData_gridEmplDet_DXEditor6_I'

CARD_HOURS_ID     = 'ctl00_ContentData_txtAvailableHoursForUnit_I'
CARD_DATE_FROM_ID = 'ctl00_ContentData_dtDutyStartDate_I'
CARD_DATE_TO_ID   = 'ctl00_ContentData_dtDutyStopDate_I'

GRID_ID = 'ctl00_ContentData_gridEmplDet'


# ── Ανάγνωση Excel ───────────────────────────────────────────────────────────
def _clean_cell(v):
    v = str(v if v is not None else '').strip()
    return '' if v.lower() in ('nan', 'none') else v


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
    # Προαιρετική στήλη — χρησιμοποιείται ΜΟΝΟ από το κουμπί ΑΠΟΦΑΣΗ
    # (run_decision). Αν υπάρχει και έχει τιμή, υπερισχύει του κειμένου
    # που δόθηκε στο παράθυρο για τη συγκεκριμένη γραμμή.
    apof_col     = _find(('Απόφαση', 'ΑΠΟΦΑΣΗ', 'Αρ. Απόφασης', 'Αριθμός Απόφασης'))

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
            'apofasi':     _clean_cell(row.get(apof_col, '')) if apof_col else '',
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


def _fill_search_field(driver, field_id, value, log):
    """Συμπληρώνει το πεδίο αναζήτησης (field_id) με value, ΕΠΙΒΕΒΑΙΩΝΟΝΤΑΣ
    ότι η τιμή όντως «κόλλησε» στο DOM — αν όχι, δοκιμάζει fallback με
    πραγματικό πληκτρολόγημα (Selenium send_keys) πριν παραιτηθεί.
    Επιστρέφει το WebElement του πεδίου (για να ξαναδιαβαστεί αν χρειαστεί)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    field = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.presence_of_element_located((By.ID, field_id)))

    _dispatch_keyboard_value(driver, field, value)
    time.sleep(0.5)
    current = (field.get_attribute('value') or '').strip()

    if current != value:
        log(f'  ⚠ Το πεδίο έμεινε «{current}» αντί για «{value}» μετά το JS — '
            f'δοκιμή με πραγματικό πληκτρολόγημα...')
        try:
            driver.execute_script('arguments[0].click();', field)
            time.sleep(0.2)
            field.clear()
            time.sleep(0.2)
            field.send_keys(value)
            time.sleep(0.3)
            field.send_keys('\t')  # blur -> trigger change/lostfocus handlers
            time.sleep(0.5)
        except Exception as e:
            log(f'  ⚠ Fallback πληκτρολόγηση απέτυχε: {e}')
        current = (field.get_attribute('value') or '').strip()

    if current != value:
        log(f'  ✗ Το πεδίο αναζήτησης ΔΕΝ συμπληρώθηκε σωστά (τελική τιμή: «{current}») — '
            'η αναζήτηση πιθανόν θα επιστρέψει λάθος/όλα τα αποτελέσματα')
    else:
        log(f'  ✓ Πεδίο αναζήτησης επιβεβαιώθηκε: «{current}»')

    return field


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

    _fill_search_field(driver, field_id, value, log)

    search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
    driver.execute_script('arguments[0].click();', search_link)
    time.sleep(3)

    return driver.find_elements(By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')


def _pick_matching_row(driver, edit_links, school_name_norm, school_code, log):
    """
    Μοναδικό ταίριασμα με βάση Ονομασία Σχολείου (κανονικοποιημένα, fuzzy)
    ή/και Κωδικό Σχολείου (κυριολεκτικό substring στο κείμενο γραμμής).

    Αν η αναζήτηση (ΑΦΜ/Α.Μ.) επέστρεψε ΜΟΝΟ ΕΝΑ αποτέλεσμα συνολικά, αυτό
    το αποτέλεσμα είναι ήδη μοναδικό (το άτομο δεν έχει άλλη τοποθέτηση) —
    το εμπιστευόμαστε ΑΝΕΥ ΟΡΩΝ, ΧΩΡΙΣ να εφαρμόζουμε επιπλέον το φίλτρο
    ονόματος/κωδικού σχολείου (ίδια λογική με το editor.py / Ε5:
    `if len(edit_links) == 1: target = edit_links[0]`). Το φίλτρο σχολείου
    χρησιμεύει μόνο για να διαλέξουμε ανάμεσα σε ΠΟΛΛΑΠΛΑ αποτελέσματα.

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

    if len(rows) == 1:
        return rows[0][0], 'ok'

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


def _normalize_combo_text(s):
    """Κανονικοποίηση κειμένου επιλογής combo για ανεκτική σύγκριση —
    αφαιρεί τόνους, πεζά/κεφαλαία, περιττά κενά."""
    import re
    import unicodedata
    if not s:
        return ''
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    s = s.upper()
    s = re.sub(r'\s+', ' ', s).strip()
    return s


STRIKE_INTERVAL = 0.3


def _send_keys_slow(element, text, delay=STRIKE_INTERVAL):
    """Πληκτρολογεί έναν-έναν χαρακτήρα με καθυστέρηση — απαραίτητο για να
    φιλτράρει σωστά η λίστα του DevExpress combo (ίδια τεχνική με το Ε5)."""
    for char in str(text):
        element.send_keys(char)
        time.sleep(delay)


def _select_work_type_combo(driver, base_id, text, log):
    """
    Επιλέγει τιμή από το combo περιγραφής ωραρίου — ΙΔΙΑ μέθοδος με το
    editor.py (Ε5) ._select_dxe_combo: πληκτρολόγηση αργά για να φιλτράρει
    η λίστα, κλικ στο ΠΡΩΤΟ td[dxtext=κείμενο] που εμφανιστεί, με fallback
    TAB αν δεν βρεθεί. Δοκιμασμένη, αξιόπιστη μέθοδος στο Ε5.

    ΣΗΜΕΙΩΣΗ: για την «ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π.» το
    dropdown δείχνει 2 γραμμές με το ΙΔΙΟ κείμενο (κατηγορία Συμπλήρωση/
    Υπερωρία) — εδώ επιλέγεται η ΠΡΩΤΗ που βρεθεί, χωρίς να ξεχωρίζεται η
    κατηγορία (συνειδητή επιλογή απλότητας, όπως το Ε5 — στην πράξη το
    MySchool δείχνει πάντα πρώτα τη «Συμπλήρωση»).
    """
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
    except Exception as e:
        log(f'  ⚠ Άνοιγμα combo περιγραφής ({text!r}): {e}')
        return False

    try:
        item = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, f'td[dxtext="{text}"]')))
        driver.execute_script('arguments[0].click();', item)
        time.sleep(0.5)
        log(f'  ✓ Επιλέχθηκε: {text}')
        return True
    except Exception:
        pass

    try:
        inp.send_keys(Keys.TAB)
        time.sleep(0.5)
        log(f'  ⚠ Επιλογή μέσω TAB (fallback): {text}')
        return True
    except Exception as e:
        log(f'  ✗ Dropdown περιγραφής: {e}')
        return False

def _date_in_text(date_str, normalized_text):
    """Ελέγχει αν η ημερομηνία date_str (πχ '7/9/2026') εμφανίζεται μέσα στο
    ήδη κανονικοποιημένο (μέσω _normalize_combo_text) κείμενο, ΑΝΕΞΑΡΤΗΤΩΣ
    leading zeros ΚΑΙ ανεξαρτήτως διαχωριστικού — ο πίνακας «Λεπτομέρειες
    ωραρίου» εμφανίζει τις ημερομηνίες με ΠΑΥΛΑ και δύο ψηφία («07-09-2026»,
    επιβεβαιωμένο με screenshot), ενώ το πεδίο της καρτέλας/φόρμας δίνει
    «7/9/2026» με «/» και χωρίς μηδενικά — οπότε ελέγχουμε όλους τους
    συνδυασμούς μηδενικού/χωρίς-μηδενικό ΚΑΙ «/»/«-»."""
    import re
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', (date_str or '').strip())
    if not m:
        return date_str in normalized_text
    d, mo, y = (int(x) for x in m.groups())
    candidates = set()
    for sep in ('/', '-'):
        for dd in (f'{d}', f'{d:02d}'):
            for mm in (f'{mo}', f'{mo:02d}'):
                candidates.add(f'{dd}{sep}{mm}{sep}{y}')
    return any(c in normalized_text for c in candidates)


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
    ok_c = _select_work_type_combo(driver, TYPE_COMBO_BASE_ID, description_text, log)
    if not ok_c:
        log(f'  ✗ Δεν επιλέχθηκε περιγραφή «{description_text}» — παράλειψη εγγραφής '
            '(δεν αποθηκεύουμε ημιτελή εγγραφή)')
        return 'error'

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
            time.sleep(0.3)
            confirmed_hours = (hours_inp.get_attribute('value') or '').strip()
            if confirmed_hours == str(person_hours):
                log(f'  ✓ Ώρες νέας εγγραφής επιβεβαιώθηκαν: «{confirmed_hours}»')
            else:
                log(f'  ⚠ Ώρες νέας εγγραφής: πεδίο δείχνει «{confirmed_hours}» αντί για '
                    f'«{person_hours}»')
        except Exception as e:
            log(f'  ⚠ Ώρες φόρμα: {e}')
    else:
        log('  ⚠ Καμία τιμή ωρών να συμπληρωθεί στη νέα εγγραφή')

    # ── Ημερομηνίες ───────────────────────────────────────────────────────
    for field_id, val, label in (
        (NEW_ROW_DATE_FROM_ID, date_from, 'Ημ. από'),
        (NEW_ROW_DATE_TO_ID, date_to, 'Ημ. έως'),
    ):
        try:
            _set_dxe_value(driver, field_id, val)
            time.sleep(0.3)
            confirmed = (driver.find_element(By.ID, field_id).get_attribute('value') or '').strip()
            # Ανεκτική σύγκριση (leading zeros) — το πεδίο συχνά κανονικοποιεί
            # την εμφάνιση («18/09/2026» -> «18/9/2026») χωρίς να αλλάζει την
            # ουσιαστική τιμή, οπότε δεν είναι πραγματικό πρόβλημα.
            if confirmed == val or _date_in_text(val, _normalize_combo_text(confirmed)):
                log(f'  ✓ {label} επιβεβαιώθηκε: «{confirmed}»')
            else:
                log(f'  ⚠ {label}: πεδίο δείχνει «{confirmed}» αντί για «{val}»')
        except Exception as e:
            log(f'  ⚠ {label}: {e}')

    time.sleep(0.5)

    # ── Αποδοχή ───────────────────────────────────────────────────────────
    # ΣΗΜΑΝΤΙΚΟ: επιβεβαιώθηκε (log χρήστη + χειροκίνητος έλεγχος στο ίδιο
    # το MySchool) ότι όταν το combo περιγραφής παραμένει ΟΡΑΤΟ μετά το
    # «Αποδοχή», η γραμμή ΔΕΝ έχει όντως καταχωρηθεί στο grid — το
    # «Αποθήκευση» που ακολουθεί απλώς αποθηκεύει την καρτέλα ΧΩΡΙΣ τη νέα
    # εγγραφή, και το log δείχνει ψευδώς «✓» παντού. Άρα ΕΔΩ μπλοκάρουμε αν
    # η γραμμή δεν έχει όντως κλείσει — δεν το αγνοούμε πια σαν
    # «πληροφοριακό».
    def _row_still_editing():
        try:
            els = driver.find_elements(By.ID, TYPE_COMBO_BASE_ID + '_I')
            return any(el.is_displayed() for el in els)
        except Exception:
            return False

    try:
        accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.XPATH, '//img[@alt="Αποδοχή"]')))
        driver.execute_script('arguments[0].click();', accept_btn)
        time.sleep(2)
    except Exception as e:
        log(f'  ✗ Αποδοχή: {e}')
        return 'error'

    if _row_still_editing():
        # 2η προσπάθεια: απευθείας κλήση του DevExpress client API αντί για
        # κλικ στο <img> — πιο αξιόπιστο, ίδια ενέργεια με το εικονίδιο
        # «Αποδοχή» (UpdateEdit κλείνει τη γραμμή επεξεργασίας του grid).
        log('  ⚠ Η γραμμή παραμένει σε λειτουργία επεξεργασίας μετά το κλικ «Αποδοχή» — '
            'δοκιμή μέσω client API του grid...')
        try:
            driver.execute_script(f"""
                var g = ASPxClientGridView.Cast('{GRID_ID}');
                if (g) g.UpdateEdit();
            """)
        except Exception as e:
            log(f'  ⚠ Client API UpdateEdit: {e}')
        time.sleep(2)

    if _row_still_editing():
        log('  ✗ Η γραμμή ΔΕΝ έγινε αποδεκτή — παραμένει σε λειτουργία επεξεργασίας και μετά '
            'τις δύο προσπάθειες (κλικ + client API) — ΔΕΝ πατιέται Αποθήκευση ώστε να μη '
            'χαθεί/μπερδευτεί η φόρμα (πιθανό πρόβλημα εγκυρότητας σε κάποιο πεδίο)')
        return 'error'

    log('  ✓ Αποδοχή (η γραμμή έκλεισε κανονικά)')

    # ── Αποθήκευση ────────────────────────────────────────────────────────
    try:
        save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_ContentData_btnSave')))
        driver.execute_script('arguments[0].click();', save_btn)
        time.sleep(3)
        log('  ✓ Αποθήκευση')
    except Exception as e:
        log(f'  ✗ Αποθήκευση: {e}')
        return 'error'

    # ── Επιβεβαίωση ΜΕΤΑ την αποθήκευση: εμφανίζεται όντως η νέα εγγραφή
    # στον πίνακα «Λεπτομέρειες ωραρίου εργασίας»; (ground-truth έλεγχος,
    # αντί να υποθέτουμε ότι το Αποδοχή+Αποθήκευση «έπιασε»). ─────────────
    try:
        grid_el = driver.find_element(By.ID, GRID_ID)
        grid_text = _normalize_combo_text(grid_el.text)
    except Exception:
        grid_text = _normalize_combo_text(driver.page_source)

    desc_ok  = _normalize_combo_text(description_text) in grid_text
    categ_ok = _normalize_combo_text(CATEGORY_TEXT) in grid_text
    from_ok  = _date_in_text(date_from, grid_text)

    if desc_ok and categ_ok and from_ok:
        log('  ✓ Επιβεβαιώθηκε: η νέα εγγραφή εμφανίζεται στον πίνακα μετά την αποθήκευση')
        return 'ok'
    else:
        log('  ⚠ ΔΕΝ επιβεβαιώθηκε ότι η νέα εγγραφή εμφανίζεται στον πίνακα μετά την '
            f'αποθήκευση (περιγραφή βρέθηκε: {desc_ok}, κατηγορία: {categ_ok}, '
            f'ημ. από: {from_ok}) — έλεγξε χειροκίνητα στο MySchool')
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

    # Αποθήκευση path για χρήση από το ΛΗΞΗ (ίδιος μηχανισμός με το Ε5) —
    # μόνο αν έγινε τουλάχιστον μία επιτυχής καταχώρηση.
    if results['ok']:
        _save_panic_path(file_path)


# ── ΛΗΞΗ — Διαγραφή εγγραφών (ίδια λογική με το editor.py / Ε5) ────────────
def _panic_path_file():
    """Path του αρχείου που κρατάει το τελευταίο excel που έτρεξε επιτυχώς
    (για το κουμπί ΛΗΞΗ). Ξεχωριστό από το αντίστοιχο του Ε5, ώστε τα δύο
    εργαλεία να μην μπερδεύουν το ένα το «τελευταίο αρχείο» του άλλου."""
    docs = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks')
    os.makedirs(docs, exist_ok=True)
    return os.path.join(docs, '.panic_last_file_whd.txt')


def _save_panic_path(file_path):
    try:
        with open(_panic_path_file(), 'w', encoding='utf-8') as f:
            f.write(file_path)
    except Exception:
        pass


def get_panic_path():
    """Επιστρέφει το path του τελευταίου excel (ή '' αν δεν υπάρχει) — για
    να προσυμπληρώνεται στο παράθυρο ΛΗΞΗ."""
    try:
        p = _panic_path_file()
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return f.read().strip()
    except Exception:
        pass
    return ''


def run_delete(ctx, driver, callback=None):
    """
    ΛΗΞΗ — διαγράφει την εγγραφή «Γραμματειακή Υποστήριξη» στον πίνακα
    «Λεπτομέρειες ωραρίου εργασίας» για κάθε άτομο του excel. Σκοπός είναι
    το γρήγορο undo αμέσως μετά από μια εσφαλμένη μαζική καταχώρηση, όχι
    στοχευμένος καθαρισμός — ΧΩΡΙΣ έλεγχο ημερομηνιών/ωρών. Επειδή όμως οι
    γραμμές του πίνακα δεν στοιχίζονται πάντα με προβλέψιμη σειρά (μπορεί
    να υπάρχει ΚΑΙ άλλη, παλαιότερη εγγραφή), ΔΕΝ διαγράφουμε τυφλά την
    πρώτη γραμμή· επιβεβαιώνουμε πρώτα ότι η γραμμή που θα διαγραφεί
    περιέχει όντως «Γραμματειακή» — 0 ή >1 τέτοιες γραμμές → παράλειψη
    (δεν μαντεύουμε). Ο εντοπισμός ατόμου/σχολείου γίνεται με την πιο
    αξιόπιστη λογική του Ε7 (ΑΦΜ/Α.Μ. + Κωδικός/Ονομασία Σχολείου).

    ctx αναμενόμενα κλειδιά:
      'file_path' — η ίδια λίστα (ή ίδιας μορφής) με αυτή που χρησιμοποιήθηκε
                    στην καταχώρηση που θέλουμε να αναιρέσουμε.
    """
    log = callback or print
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    file_path = ctx.get('file_path')
    if not file_path:
        log('✗ Δεν δόθηκε αρχείο — τέλος.')
        return

    people = load_people(file_path, log=log)
    if not people:
        return

    total = len(people)
    log(f'\n  {total} εγγραφές προς ΔΙΑΓΡΑΦΗ (Λήξη — πρώτη/πιο πρόσφατη εγγραφή ανά άτομο)')

    ok = fail = 0

    try:
        for idx, person in enumerate(people, 1):
            ident = person['afm'] or person['am']
            log(f'\n[{idx}/{total}] {ident}  {person["eponymo"]} {person["onoma"]}  '
                f'—  {person["school_name"] or person["school_code"]}')

            try:
                edit_links = _search_person(driver, person, log)
            except Exception as e:
                log(f'  ✗ Αναζήτηση απέτυχε: {e}')
                fail += 1
                continue

            if not edit_links:
                log('  ✗ Κανένα αποτέλεσμα αναζήτησης')
                fail += 1
                continue

            school_name_norm = _normalize_school_name(person['school_name']) if person['school_name'] else ''
            link, reason = _pick_matching_row(driver, edit_links, school_name_norm,
                                               person['school_code'], log)
            if reason != 'ok':
                if reason == 'notfound':
                    log(f'  ⚠ Δεν βρέθηκε γραμμή με σχολείο «{person["school_name"] or person["school_code"]}» '
                        f'({len(edit_links)} αποτέλεσμα(-τα) συνολικά) — παράλειψη')
                else:
                    log('  ⚠ Πάνω από μία γραμμές ταιριάζουν — παράλειψη (χειροκίνητος έλεγχος)')
                fail += 1
                continue

            try:
                driver.execute_script('arguments[0].click();', link)
                time.sleep(3)
                log('  Καρτέλα ανοιχτή')
            except Exception as e:
                log(f'  ✗ Καρτέλα: {e}')
                fail += 1
                continue

            # ── Κουμπί Διαγραφή — ΜΟΝΟ αν η γραμμή είναι όντως
            # «Γραμματειακή Υποστήριξη» ────────────────────────────────────
            # Οι γραμμές του πίνακα «Λεπτομέρειες ωραρίου εργασίας» δεν
            # στοιχίζονται πάντα με προβλέψιμη σειρά (μπορεί να υπάρχει
            # ΚΑΙ άλλη, παλαιότερη εγγραφή) — άρα ΔΕΝ παίρνουμε πλέον
            # τυφλά το πρώτο εικονίδιο «Διαγραφή» στη σειρά DOM. Μαζεύουμε
            # ΟΛΑ τα εικονίδια «Διαγραφή» της gridEmplDet, διαβάζουμε το
            # κείμενο της γραμμής (ancestor <tr>) του καθενός, και
            # προχωράμε ΜΟΝΟ αν ΑΚΡΙΒΩΣ ΜΙΑ γραμμή περιέχει «Γραμματειακή» —
            # 0 ή >1 ταιριάσματα → παράλειψη με καθαρό log (δεν μαντεύουμε).
            try:
                del_links = driver.find_elements(
                    By.XPATH, '//img[@alt="Διαγραφή" and contains(@onclick,"gridEmplDet")]')
            except Exception as e:
                log(f'  ✗ Διαγραφή: {e}')
                fail += 1
                continue

            if not del_links:
                log('  ✗ Καμία εγγραφή στις «Λεπτομέρειες ωραρίου» προς διαγραφή')
                fail += 1
                continue

            gram_candidates = []
            for dl in del_links:
                try:
                    row = dl.find_element(By.XPATH, './ancestor::tr[1]')
                    if 'ΓΡΑΜΜΑΤΕΙΑΚΗ' in _normalize_combo_text(row.text):
                        gram_candidates.append(dl)
                except Exception:
                    continue

            if len(gram_candidates) == 1:
                del_btn = gram_candidates[0]
            elif len(gram_candidates) == 0:
                log(f'  ✗ Δεν βρέθηκε γραμμή «Γραμματειακή Υποστήριξη» προς διαγραφή '
                    f'({len(del_links)} εγγραφή(-ές) συνολικά στον πίνακα) — παράλειψη '
                    '(δεν διαγράφουμε άλλη/λάθος εγγραφή)')
                fail += 1
                continue
            else:
                log(f'  ⚠ Βρέθηκαν {len(gram_candidates)} γραμμές «Γραμματειακή Υποστήριξη» — '
                    'παράλειψη (χειροκίνητος έλεγχος, δεν μαντεύουμε ποια)')
                fail += 1
                continue

            try:
                driver.execute_script('arguments[0].scrollIntoView({block:"center"});', del_btn)
                time.sleep(0.5)
                driver.execute_script('arguments[0].click();', del_btn)
                time.sleep(1)
                log('  Διαγραφή κλικ (επιβεβαιωμένη γραμμή «Γραμματειακή»)')
            except Exception as e:
                log(f'  ✗ Διαγραφή: {e}')
                fail += 1
                continue

            # ── Επιβεβαίωση (OK στο confirm dialog) ─────────────────────────
            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                driver.switch_to.alert.accept()
                time.sleep(2)
                log('  ✓ Επιβεβαίωση ΟΚ')
            except Exception:
                log('  ⚠ Alert δεν εμφανίστηκε')

            # ── Αποθήκευση ───────────────────────────────────────────────────
            try:
                save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                    EC.element_to_be_clickable((By.ID, 'ctl00_ContentData_btnSave')))
                driver.execute_script('arguments[0].click();', save_btn)
                time.sleep(3)
                log('  ✓ Αποθήκευση')
                ok += 1
            except Exception as e:
                log(f'  ✗ Αποθήκευση: {e}')
                fail += 1
    except KeyboardInterrupt:
        log('\n\n⚠ Διακόπηκε από τον χρήστη.')

    log('\n' + '─' * 65)
    log(f'ΛΗΞΗ — Διαγραφές: {ok} επιτυχείς, {fail} αποτυχίες')
    log('─' * 65)


# ── ΑΠΟΦΑΣΗ — Συμπλήρωση Παρατηρήσεων σε ΥΠΑΡΧΟΥΣΑ «Γραμματειακή» ─────────────
# Δεν προσθέτει νέα εγγραφή. Για κάθε άτομο του excel:
#   1. Αναζήτηση (ΑΦΜ/Α.Μ.) + μοναδικό ταίριασμα σχολείου (ίδια λογική με Ε5).
#   2. Στον πίνακα «Λεπτομέρειες ωραρίου εργασίας» εντοπίζει τη γραμμή που
#      περιέχει «Γραμματειακή» — ΑΚΡΙΒΩΣ ΜΙΑ, αλλιώς παράλειψη.
#   3. Την ανοίγει για επεξεργασία (εικονίδιο ή client API StartEditRow).
#   4. Σβήνει ό,τι υπάρχει στις Παρατηρήσεις (DXEditor7 — ίδιο πεδίο με το
#      «ΠΔΕ-» του Ε8) και γράφει το κείμενο της απόφασης.
#   5. Αποδοχή → Αποθήκευση → έλεγχος ότι το νέο κείμενο φαίνεται στον πίνακα.
# Η παλιά τιμή των Παρατηρήσεων γράφεται στο log ΚΑΙ σε CSV backup
# (Documents/MySchoolChecks/apofasi_backup_*.csv) για χειροκίνητη επαναφορά.

NOTES_EDITOR_ID = 'ctl00_ContentData_gridEmplDet_DXEditor7_I'   # Παρατηρήσεις
DATA_ROW_PREFIX = GRID_ID + '_DXDataRow'
EDIT_ROW_PREFIX = GRID_ID + '_DXEditingRow'


def _find_gram_row(driver, log):
    """Επιστρέφει (tr_element, visible_index) της ΜΟΝΑΔΙΚΗΣ γραμμής
    «Γραμματειακή» ή (None, reason) με reason ∈ {'notfound','ambiguous'}."""
    from selenium.webdriver.common.by import By
    rows = driver.find_elements(By.CSS_SELECTOR, f'tr[id^="{DATA_ROW_PREFIX}"]')
    cands = []
    for tr in rows:
        try:
            if 'ΓΡΑΜΜΑΤΕΙΑΚΗ' in _normalize_combo_text(tr.text):
                idx = int(tr.get_attribute('id')[len(DATA_ROW_PREFIX):])
                cands.append((tr, idx))
        except Exception:
            continue
    if len(cands) == 1:
        return cands[0]
    if not cands:
        log(f'  ✗ Δεν βρέθηκε γραμμή «Γραμματειακή Υποστήριξη» '
            f'({len(rows)} εγγραφή(-ές) στον πίνακα) — παράλειψη')
        return None, 'notfound'
    log(f'  ⚠ Βρέθηκαν {len(cands)} γραμμές «Γραμματειακή» — παράλειψη '
        '(χειροκίνητος έλεγχος, δεν μαντεύουμε ποια)')
    return None, 'ambiguous'


def _editing_open(driver):
    from selenium.webdriver.common.by import By
    try:
        els = driver.find_elements(By.ID, NOTES_EDITOR_ID)
        return any(el.is_displayed() for el in els)
    except Exception:
        return False


def _open_row_for_edit(driver, tr, vis_idx, log):
    """Ανοίγει τη γραμμή για επεξεργασία. 1η προσπάθεια: το εικονίδιο
    επεξεργασίας ΜΕΣΑ στη γραμμή· 2η: ASPxClientGridView.StartEditRow."""
    from selenium.webdriver.common.by import By

    try:
        driver.execute_script('arguments[0].scrollIntoView({block:"center"});', tr)
        time.sleep(0.5)
    except Exception:
        pass

    icons = []
    try:
        icons = tr.find_elements(
            By.XPATH, './/img[@alt="Επεξεργασία" or @alt="Διόρθωση" or @alt="Edit"]')
    except Exception:
        pass
    if icons:
        try:
            driver.execute_script('arguments[0].click();', icons[0])
            time.sleep(2)
        except Exception as e:
            log(f'  ⚠ Κλικ εικονιδίου επεξεργασίας: {e}')

    if not _editing_open(driver):
        log('  (δοκιμή ανοίγματος μέσω client API του grid...)')
        try:
            driver.execute_script(f"""
                var g = ASPxClientGridView.Cast('{GRID_ID}');
                if (g) g.StartEditRow(arguments[0]);
            """, vis_idx)
        except Exception as e:
            log(f'  ⚠ StartEditRow: {e}')
        time.sleep(2.5)

    return _editing_open(driver)


def _set_notes(driver, text, log):
    """Σβήνει τις Παρατηρήσεις και γράφει text. Επιστρέφει (ok, old_value)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    el = driver.find_element(By.ID, NOTES_EDITOR_ID)
    old = (el.get_attribute('value') or '').strip()
    log(f'  Παλιές Παρατηρήσεις: «{old}»')

    _set_dxe_value(driver, NOTES_EDITOR_ID, '')
    time.sleep(0.3)
    _set_dxe_value(driver, NOTES_EDITOR_ID, text)
    time.sleep(0.4)
    cur = (driver.find_element(By.ID, NOTES_EDITOR_ID).get_attribute('value') or '').strip()

    if cur != text.strip():
        log('  ⚠ Η τιμή δεν «κόλλησε» μέσω JS — δοκιμή με πληκτρολόγηση...')
        try:
            el = driver.find_element(By.ID, NOTES_EDITOR_ID)
            driver.execute_script('arguments[0].click(); arguments[0].focus();', el)
            el.send_keys(Keys.CONTROL, 'a')
            el.send_keys(Keys.DELETE)
            time.sleep(0.2)
            el.send_keys(text)
            time.sleep(0.3)
        except Exception as e:
            log(f'  ⚠ Πληκτρολόγηση: {e}')
        cur = (driver.find_element(By.ID, NOTES_EDITOR_ID).get_attribute('value') or '').strip()

    if cur == text.strip():
        log(f'  ✓ Νέες Παρατηρήσεις: «{cur}»')
        return True, old
    log(f'  ✗ Το πεδίο δείχνει «{cur}» αντί για «{text}»')
    return False, old


def _backup_file():
    from datetime import datetime
    docs = os.path.join(os.path.expanduser('~'), 'Documents', 'MySchoolChecks')
    os.makedirs(docs, exist_ok=True)
    return os.path.join(docs, f'apofasi_backup_{datetime.now():%Y%m%d_%H%M}.csv')


def _append_backup(path, person, old, new):
    import csv
    first = not os.path.exists(path)
    try:
        with open(path, 'a', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f, delimiter=';')
            if first:
                w.writerow(['ΑΦΜ', 'Α.Μ.', 'Επώνυμο', 'Όνομα', 'Σχολείο',
                            'Παλιές Παρατηρήσεις', 'Νέες Παρατηρήσεις'])
            w.writerow([person['afm'], person['am'], person['eponymo'], person['onoma'],
                        person['school_name'] or person['school_code'], old, new])
    except Exception:
        pass


def process_person_decision(driver, person, text, backup_path, log):
    """'ok' | 'notfound' | 'ambiguous' | 'error'"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    school_name_norm = _normalize_school_name(person['school_name']) if person['school_name'] else ''
    try:
        edit_links = _search_person(driver, person, log)
    except Exception as e:
        log(f'  ✗ Αναζήτηση απέτυχε: {e}')
        return 'error'
    if not edit_links:
        log('  ✗ Κανένα αποτέλεσμα αναζήτησης')
        return 'notfound'

    link, reason = _pick_matching_row(driver, edit_links, school_name_norm,
                                       person['school_code'], log)
    if reason != 'ok':
        log('  ⚠ ' + ('Δεν βρέθηκε γραμμή με το σχολείο' if reason == 'notfound'
                      else 'Πάνω από μία γραμμές ταιριάζουν') + ' — παράλειψη')
        return reason

    try:
        driver.execute_script('arguments[0].click();', link)
        time.sleep(3)
        WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.ID, GRID_ID)))
        log('  Καρτέλα ανοιχτή')
    except Exception as e:
        log(f'  ✗ Άνοιγμα καρτέλας: {e}')
        return 'error'

    tr, vis_idx = _find_gram_row(driver, log)
    if tr is None:
        return vis_idx   # 'notfound' / 'ambiguous'

    if not _open_row_for_edit(driver, tr, vis_idx, log):
        log('  ✗ Η γραμμή «Γραμματειακή» δεν άνοιξε για επεξεργασία — παράλειψη')
        return 'error'
    log('  ✓ Γραμμή «Γραμματειακή» σε επεξεργασία')

    ok_n, old = _set_notes(driver, text, log)
    if not ok_n:
        try:
            driver.execute_script(f"ASPxClientGridView.Cast('{GRID_ID}').CancelEdit();")
        except Exception:
            pass
        log('  ✗ Ακύρωση επεξεργασίας — ΔΕΝ αποθηκεύτηκε τίποτα')
        return 'error'

    # ── Αποδοχή ───────────────────────────────────────────────────────────
    try:
        accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located((By.XPATH, '//img[@alt="Αποδοχή"]')))
        driver.execute_script('arguments[0].click();', accept_btn)
        time.sleep(2)
    except Exception as e:
        log(f'  ⚠ Κλικ Αποδοχή: {e}')
    if _editing_open(driver):
        log('  ⚠ Η γραμμή παραμένει σε επεξεργασία — δοκιμή UpdateEdit μέσω client API...')
        try:
            driver.execute_script(f"ASPxClientGridView.Cast('{GRID_ID}').UpdateEdit();")
        except Exception as e:
            log(f'  ⚠ UpdateEdit: {e}')
        time.sleep(2)
    if _editing_open(driver):
        log('  ✗ Η γραμμή ΔΕΝ έγινε αποδεκτή — ΔΕΝ πατιέται Αποθήκευση')
        return 'error'
    log('  ✓ Αποδοχή')

    # ── Αποθήκευση ────────────────────────────────────────────────────────
    try:
        save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_ContentData_btnSave')))
        driver.execute_script('arguments[0].click();', save_btn)
        time.sleep(3)
        log('  ✓ Αποθήκευση')
    except Exception as e:
        log(f'  ✗ Αποθήκευση: {e}')
        return 'error'

    _append_backup(backup_path, person, old, text)

    # ── Επιβεβαίωση στον πίνακα (πρώτοι 15 χαρακτήρες — ο πίνακας μπορεί
    #    να κόβει μεγάλα κείμενα) ──────────────────────────────────────────
    try:
        grid_text = _normalize_combo_text(driver.find_element(By.ID, GRID_ID).text)
    except Exception:
        grid_text = _normalize_combo_text(driver.page_source)
    if _normalize_combo_text(text)[:15] in grid_text:
        log('  ✓ Επιβεβαιώθηκε: το κείμενο της απόφασης εμφανίζεται στον πίνακα')
        return 'ok'
    log('  ⚠ ΔΕΝ επιβεβαιώθηκε το νέο κείμενο στον πίνακα — έλεγξε χειροκίνητα')
    return 'error'


def run_decision(ctx, driver, callback=None):
    """
    ΑΠΟΦΑΣΗ — ctx:
      'file_path' — ίδιας μορφής excel με το Ε5 (+ προαιρετική στήλη «Απόφαση»)
      'text'      — κείμενο απόφασης για όλους (η στήλη «Απόφαση» υπερισχύει
                    ανά γραμμή όταν έχει τιμή)
    """
    log = callback or print
    file_path    = ctx.get('file_path')
    default_text = (ctx.get('text') or '').strip()

    people = load_people(file_path, log=log)
    if not people:
        return

    backup_path = _backup_file()
    total = len(people)
    log(f'\n  {total} εγγραφές — ΑΠΟΦΑΣΗ (Παρατηρήσεις σε υπάρχουσα «Γραμματειακή»)')
    log(f'  Backup παλιών τιμών: {backup_path}')

    results = {'ok': [], 'notfound': [], 'ambiguous': [], 'error': []}
    try:
        for idx, person in enumerate(people, 1):
            ident = person['afm'] or person['am']
            log(f'\n[{idx}/{total}] {ident}  {person["eponymo"]} {person["onoma"]}  '
                f'—  {person["school_name"] or person["school_code"]}')
            text = person.get('apofasi') or default_text
            if not text:
                log('  ✗ Κενό κείμενο απόφασης (ούτε στο παράθυρο ούτε στη στήλη «Απόφαση») — παράλειψη')
                results['error'].append(ident)
                continue
            status = process_person_decision(driver, person, text, backup_path, log)
            results[status].append(ident)
            time.sleep(0.5)
    except KeyboardInterrupt:
        log('\n\n⚠ Διακόπηκε από τον χρήστη. Ό,τι έχει ήδη αποθηκευτεί παραμένει.')

    log('\n' + '─' * 65)
    log(f'✓ Ενημερώθηκαν : {len(results["ok"])}')
    log(f'⚠ Δεν βρέθηκαν (σχολείο ή γραμμή Γραμματειακής) : {len(results["notfound"])}')
    log(f'⚠ Διφορούμενα : {len(results["ambiguous"])}')
    log(f'✗ Σφάλματα : {len(results["error"])}')
    for key, label in (('notfound', 'Δεν βρέθηκαν'), ('ambiguous', 'Διφορούμενα'),
                       ('error', 'Σφάλματα')):
        if results[key]:
            log(f'  {label}: ' + ' | '.join(results[key]))
    log('─' * 65)
