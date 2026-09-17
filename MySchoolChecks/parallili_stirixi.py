#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
parallili_stirixi.py
=====================
One-off εργασία (ΔΕΝ μπαίνει στο μενού της εφαρμογής) — καταχώρηση νέας
εγγραφής «Λεπτομέρειες ωραρίου εργασίας» στην καρτέλα Τοποθέτησης του
MySchool, για μια λίστα ατόμων από Excel (π.χ. ΠΑΡΑΛΛΗΛΕΣ_ΕΒΠ_ΕΕΠ.xlsx).

Βάλε αυτό το αρχείο μέσα στον φάκελο MySchoolChecks\\ (δίπλα στο main.py,
config.py, placements.py, editor.py κλπ.) και τρέξε το με:

    python parallili_stirixi.py "ΠΑΡΑΛΛΗΛΕΣ_ΕΒΠ_ΕΕΠ.xlsx"

(Χωρίς όρισμα, ανοίγει παράθυρο επιλογής αρχείου. Προαιρετικό 2ο όρισμα =
αριθμός ατόμων για δοκιμαστικό τρέξιμο, π.χ. `... 5`.)

Excel — αναμενόμενες στήλες (όπως στο ΠΑΡΑΛΛΗΛΕΣ_ΕΒΠ_ΕΕΠ.xlsx):
    Κωδικός Σχολείου | Ονομασία Σχολείου | Α.Φ.Μ. | Επώνυμο | Όνομα |
    Κωδικός Κύριας Ειδικότητας | Ώρες Υποχ. Διδακτικού Ωραρίου Υπηρέτησης στο Φορέα

Ροή ανά άτομο:
  1. Αναζήτηση στο https://app.myschool.sch.gr/Worker.list.myEmplUnit.aspx
     με βάση το Α.Φ.Μ. (ίδιο πεδίο/μηχανισμός με το editor.py — txtTaxNumber,
     ΑΦΜ συμπληρωμένο σε 9 ψηφία).
  2. Στα αποτελέσματα (μπορεί να βγουν πάνω από μία γραμμές — π.χ. πάνω από
     μία τοποθέτηση) εντοπίζεται η ΜΟΝΑΔΙΚΗ γραμμή που αναφέρει το σχολείο
     της στήλης «Ονομασία Σχολείου» του excel (fuzzy match, ίδια λογική
     κανονικοποίησης με το placements.py — ανεκτικό σε τόνους/συντομογραφίες,
     όπως στο allagiorontopothetisis.py). Αν βρεθούν 0 ή >1 τέτοιες γραμμές
     → η εγγραφή ΑΓΝΟΕΙΤΑΙ και καταγράφεται (καμία αλλαγή) — δεν μαντεύουμε.
  3. Άνοιγμα της καρτέλας (κλικ στο «Διόρθωση»).
  4. Κλικ στον σταυρό/γρανάζι προσθήκης της ενότητας «Λεπτομέρειες ωραρίου
     εργασίας» (ίδιο κουμπί με το editor.py: gridEmplDet_header0_new).
  5. Στο dropdown περιγραφής επιλέγεται το κείμενο:
         «ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π.»
     (ίδιος μηχανισμός επιλογής με το editor.py — cmbWorkHoursDetailsType).
  6. ΣΗΜΑΝΤΙΚΗ ΠΑΡΑΔΟΧΗ (δεν έχει επιβεβαιωθεί ακόμα πάνω στην πραγματική
     σελίδα, γιατί δεν υπήρχε δυνατότητα ζωντανής επιθεώρησης της φόρμας σε
     αυτό το περιβάλλον): μετά την επιλογή περιγραφής, το script ψάχνει
     ΑΥΤΟΜΑΤΑ μέσα στην ίδια γραμμή προσθήκης για ένα ΔΕΥΤΕΡΟ combo/dropdown
     (πέρα από αυτό της περιγραφής) — αν βρεθεί, το ανοίγει και επιλέγει το
     ΠΡΩΤΟ διαθέσιμο στοιχείο του (= «κατηγορία: συμπλήρωση», η 1η από τις 2
     επιλογές που περιέγραψες). Αν ΔΕΝ βρεθεί δεύτερο combo, το script δεν
     αγγίζει τίποτα άλλο (πιθανό να είναι ενσωματωμένο στην περιγραφή) και
     το καταγράφει καθαρά στο log.
     ΓΙ' ΑΥΤΟ: στην ΠΡΩΤΗ εγγραφή που θα φτάσει μέχρι εδώ, το script
     ΣΤΑΜΑΤΑΕΙ πριν το «Αποδοχή» και σου ζητάει να ελέγξεις με το μάτι στο
     ίδιο το παράθυρο του Chrome ότι η περιγραφή/κατηγορία/ημερομηνίες είναι
     σωστές, πατώντας Enter για να συνεχίσει (ή Ctrl+C για διακοπή). Μετά
     την πρώτη επιβεβαίωση, οι επόμενες εγγραφές τρέχουν χωρίς διακοπή.
  7. Ημερομηνίες περιόδου:
       «Από»  = η ημ. έναρξης τοποθέτησης ΠΟΥ ΗΔΗ ΕΧΕΙ Ο ΚΑΘΕΝΑΣ στην
                καρτέλα του (πεδίο dtDutyStartDate) — ΔΕΝ είναι ίδια για
                όλους, διαβάζεται ξεχωριστά ανά άτομο. Αν δεν βρεθεί τιμή
                εκεί, η εγγραφή παραλείπεται (δεν μαντεύουμε ημερομηνία).
       «Έως»  = σταθερή τιμή 21/6/2027 για όλους, όπως ζητήθηκε.
  8. Ώρες νέας εγγραφής (DXEditor4) = η τιμή του πεδίου «Διαθέσιμες ώρες
     μονάδας» (txtAvailableHoursForUnit) που ήδη έχει ο καθένας στην
     καρτέλα του (ίδιο πεδίο/λογική με το editor.py) — διαβάζεται πριν
     ανοίξει η φόρμα προσθήκης και μπαίνει αυτούσια στη νέα εγγραφή.
  9. Αποδοχή (πράσινο τικ) → Αποθήκευση.

Τρέχει με ΟΡΑΤΟ Chrome, ένα άτομο τη φορά, με πλήρες log στην κονσόλα —
μπορείς να το παρακολουθείς live. Αν κάτι δείχνει λάθος, Ctrl+C οποιαδήποτε
στιγμή· ό,τι έχει ήδη αποθηκευτεί μέχρι εκείνο το σημείο παραμένει (δεν
γίνεται rollback), αλλά τίποτα παραπέρα δεν θα αγγιχτεί.
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE_URL     = 'https://app.myschool.sch.gr'
SEARCH_URL   = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

DESCRIPTION_TEXT = 'ΠΑΡΑΛΛΗΛΗ ΣΤΗΡΙΞΗ / ΣΤΗΡΙΞΗ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π.'
# Η «Από» ημερομηνία ΔΕΝ είναι σταθερή — διαβάζεται ανά άτομο από το πεδίο
# dtDutyStartDate της υπάρχουσας καρτέλας τοποθέτησης (βλ. process_person).
DUTY_TO = '21/6/2027'  # σταθερή για όλους, όπως ζητήθηκε

TYPE_COMBO_BASE_ID = 'ctl00_ContentData_gridEmplDet_editnew_2_cmbWorkHoursDetailsType'
ADD_BTN_ID          = 'ctl00_ContentData_gridEmplDet_header0_new'
DATE_FROM_ID        = 'ctl00_ContentData_gridEmplDet_DXEditor5_I'
DATE_TO_ID           = 'ctl00_ContentData_gridEmplDet_DXEditor6_I'
HOURS_FIELD_ID_SUFFIX = 'DXEditor4'


# ── Ανάγνωση Excel ───────────────────────────────────────────────────────────
def load_people(file_path, log=print):
    """Διαβάζει το excel και επιστρέφει λίστα από dict
    {'afm','eponymo','onoma','sxoleio','kodikos_sxoleiou','kladeos'}."""
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
    school_col  = _find(('Ονομασία Σχολείου', 'ΣΧΟΛΕΙΟ', 'Σχολείο'))
    epon_col    = _find(('Επώνυμο', 'ΕΠΩΝΥΜΟ'))
    onoma_col   = _find(('Όνομα', 'ΟΝΟΜΑ'))
    kodsx_col   = _find(('Κωδικός Σχολείου',))
    kladeos_col = _find(('Κωδικός Κύριας Ειδικότητας', 'ΚΛΑΔΟΣ'))

    missing = [n for n, v in (
        ('Α.Φ.Μ.', afm_col), ('Ονομασία Σχολείου', school_col),
    ) if v is None]
    if missing:
        raise ValueError(f'Λείπουν στήλες: {", ".join(missing)}. Διαθέσιμες: {list(df.columns)}')

    people = []
    for _, row in df.iterrows():
        afm_raw = str(row.get(afm_col, '')).strip()
        if not afm_raw or afm_raw.lower() in ('nan', 'none', ''):
            continue
        afm = afm_raw.zfill(9)
        people.append({
            'afm':              afm,
            'sxoleio':          str(row.get(school_col, '')).strip(),
            'eponymo':          str(row.get(epon_col, '')).strip() if epon_col else '',
            'onoma':            str(row.get(onoma_col, '')).strip() if onoma_col else '',
            'kodikos_sxoleiou': str(row.get(kodsx_col, '')).strip() if kodsx_col else '',
            'kladeos':          str(row.get(kladeos_col, '')).strip() if kladeos_col else '',
        })
    log(f'  ✓ Διαβάστηκαν {len(people)} εγγραφές από το excel.')
    return people


def connect(log=print):
    """Άνοιγμα ορατού Chrome + login στο MySchool — ίδιο μοτίβο με τα
    υπόλοιπα one-off scripts του project."""
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
    """Πάει στη σελίδα αναζήτησης, συμπληρώνει ΑΦΜ, πατάει αναζήτηση.
    Επιστρέφει λίστα <a> edit links (Διόρθωση) — ίδιο πεδίο με το editor.py
    (txtTaxNumber)."""
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


def _send_keys_slow(element, text, delay=0.3):
    for char in str(text):
        element.send_keys(char)
        time.sleep(delay)


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
        _send_keys_slow(inp, text)
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


def _find_extra_combo_base_ids(driver, exclude_base_id, log):
    """Ψάχνει μέσα στην ΙΔΙΑ γραμμή προσθήκης (tr) για ΑΛΛΑ DevExpress combo
    inputs, πέρα από το γνωστό combo περιγραφής (exclude_base_id) και τα
    γνωστά πεδία ωρών/ημερομηνιών (DXEditor4/5/6). Επιστρέφει λίστα από
    base_id (χωρίς το τελικό «_I»)."""
    js = """
        var known = document.getElementById(arguments[0] + '_I');
        if (!known) return [];
        var row = known.closest('tr');
        if (!row) return [];
        var inputs = row.querySelectorAll('input[id$="_I"]');
        var res = [];
        inputs.forEach(function(inp){ res.push(inp.id); });
        return res;
    """
    try:
        ids = driver.execute_script(js, exclude_base_id) or []
    except Exception as e:
        log(f'  ⚠ Αναζήτηση επιπλέον combo: {e}')
        return []

    combo_bases = []
    for full_id in ids:
        base = full_id[:-2] if full_id.endswith('_I') else full_id
        if base == exclude_base_id:
            continue
        if base.endswith(('DXEditor4', 'DXEditor5', 'DXEditor6')):
            continue
        combo_bases.append(base)
    return combo_bases


def _open_combo_and_pick_first(driver, base_id, log):
    """Ανοίγει combo (base_id) και επιλέγει το ΠΡΩΤΟ διαθέσιμο ορατό στοιχείο
    (= «η 1η από τις δύο επιλογές» όπως περιγράφηκε για την κατηγορία).
    Επιστρέφει το κείμενο που επιλέχθηκε, ή None αν απέτυχε."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    try:
        inp = driver.find_element(By.ID, base_id + '_I')
        driver.execute_script('arguments[0].click();', inp)
        time.sleep(0.8)
        items = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'td[dxtext]')))
        visible_items = [it for it in items if it.is_displayed()]
        if not visible_items:
            log(f'  ⚠ Combo {base_id}: άνοιξε αλλά δεν βρέθηκαν επιλογές')
            return None
        first = visible_items[0]
        text = first.get_attribute('dxtext')
        driver.execute_script('arguments[0].click();', first)
        time.sleep(0.5)
        return text
    except Exception as e:
        log(f'  ⚠ Άνοιγμα combo {base_id}: {e}')
        return None


_first_confirmed = False  # global — χειροκίνητος έλεγχος μόνο στην 1η εγγραφή


def process_person(driver, person, log):
    """Επεξεργάζεται μία εγγραφή. Επιστρέφει status string:
    'ok' | 'notfound' | 'ambiguous' | 'error' | 'aborted'."""
    global _first_confirmed
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    afm     = person['afm']
    sxoleio = person['sxoleio']
    target_school_norm = _normalize_school_name(sxoleio)

    try:
        edit_links = _search_by_afm(driver, afm, log)
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
        time.sleep(3)
        log('  Καρτέλα ανοιχτή')
    except Exception as e:
        log(f'  ✗ Άνοιγμα καρτέλας: {e}')
        return 'error'

    # ── Ανάγνωση τρεχουσών τιμών από την καρτέλα τοποθέτησης (ανά άτομο) ──
    # Ώρες: ίδιο πεδίο με το editor.py (txtAvailableHoursForUnit).
    try:
        hours_el = WebDriverWait(driver, TIME_TO_WAIT).until(
            EC.presence_of_element_located(
                (By.ID, 'ctl00_ContentData_txtAvailableHoursForUnit_I')))
        person_hours = (hours_el.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ⚠ Διαθέσιμες ώρες μονάδας: {e}')
        person_hours = ''

    # Ημ. έναρξης τοποθέτησης: πεδίο dtDutyStartDate — χρησιμοποιείται σαν
    # «Ημ. από» της νέας εγγραφής, ΞΕΧΩΡΙΣΤΑ για τον καθένα (όχι σταθερή τιμή).
    try:
        duty_start_el = driver.find_element(By.ID, 'ctl00_ContentData_dtDutyStartDate_I')
        person_duty_from = (duty_start_el.get_attribute('value') or '').strip()
    except Exception as e:
        log(f'  ⚠ Ημ. έναρξης τοποθέτησης (dtDutyStartDate): {e}')
        person_duty_from = ''

    log(f'  (ώρες μονάδας: «{person_hours}», ημ. έναρξης τοποθέτησης: «{person_duty_from}»)')

    if not person_duty_from:
        log('  ✗ Δεν βρέθηκε ημερομηνία έναρξης τοποθέτησης στην καρτέλα — '
            'παράλειψη (δεν μαντεύουμε ημερομηνία)')
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
    ok_c = _select_dxe_combo(driver, TYPE_COMBO_BASE_ID, DESCRIPTION_TEXT, log)
    log(f'  {"✓" if ok_c else "⚠"} Περιγραφή: {DESCRIPTION_TEXT}')

    # ── Ώρες νέας εγγραφής (DXEditor4) — ίδια τιμή με «Διαθέσιμες ώρες
    # μονάδας» που διαβάσαμε παραπάνω, όπως στο editor.py ───────────────
    if person_hours:
        try:
            hours_inp = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.ID, f'ctl00_ContentData_gridEmplDet_{HOURS_FIELD_ID_SUFFIX}_I')))
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

    # ── Κατηγορία (αν υπάρχει ξεχωριστό 2ο combo) ───────────────────────
    extra_bases = _find_extra_combo_base_ids(driver, TYPE_COMBO_BASE_ID, log)
    category_text = None
    if extra_bases:
        category_text = _open_combo_and_pick_first(driver, extra_bases[0], log)
        if category_text:
            log(f'  ✓ Κατηγορία (1η επιλογή): {category_text}')
        else:
            log('  ⚠ Βρέθηκε 2ο combo αλλά η επιλογή απέτυχε — έλεγξε χειροκίνητα')
    else:
        log('  ℹ Δεν βρέθηκε ξεχωριστό combo κατηγορίας — πιθανόν ενσωματωμένο στην περιγραφή')

    # ── Ημερομηνίες περιόδου ─────────────────────────────────────────────
    # «Από»: η ημ. έναρξης τοποθέτησης ΤΟΥ ΚΑΘΕ ατόμου (dtDutyStartDate).
    # «Έως»: σταθερή τιμή DUTY_TO (21/6/2027), όπως ζητήθηκε.
    try:
        _set_dxe_value(driver, DATE_FROM_ID, person_duty_from)
        log(f'  ✓ Ημ. από: {person_duty_from}')
    except Exception as e:
        log(f'  ⚠ Ημ. από: {e}')
    try:
        _set_dxe_value(driver, DATE_TO_ID, DUTY_TO)
        log(f'  ✓ Ημ. έως: {DUTY_TO}')
    except Exception as e:
        log(f'  ⚠ Ημ. έως: {e}')

    time.sleep(0.5)

    # ── Χειροκίνητη επιβεβαίωση ΜΟΝΟ στην πρώτη εγγραφή ─────────────────
    if not _first_confirmed:
        log('\n  ⏸  ΕΛΕΓΞΕ ΤΩΡΑ στο παράθυρο του Chrome ότι η περιγραφή, η')
        log('     κατηγορία, οι ώρες και οι ημερομηνίες είναι σωστές πριν συνεχίσω.')
        try:
            input('     Πάτα Enter για να συνεχίσω (Αποδοχή+Αποθήκευση), ή Ctrl+C για διακοπή... ')
        except KeyboardInterrupt:
            raise
        _first_confirmed = True

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


def _ask_for_file():
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askopenfilename(
        title='Επίλεξε το αρχείο Excel (Α.Φ.Μ., Ονομασία Σχολείου, ...)',
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
    print('  Καταχώρηση ΠΑΡΑΛΛΗΛΗΣ ΣΤΗΡΙΞΗΣ / ΣΤΗΡΙΞΗΣ ΑΠΟ Ε.Ε.Π.-Ε.Β.Π.')
    print(f'  Έως: {DUTY_TO}  (η ημ. έναρξης διαβάζεται ξεχωριστά ανά άτομο)')
    print('=' * 65)

    people = load_people(file_path)
    if limit:
        people = people[:limit]
        print(f'  (δοκιμαστικό τρέξιμο — μόνο οι πρώτες {limit} εγγραφές)')

    driver = connect()
    if driver is None:
        print('✗ Αποτυχία σύνδεσης — τέλος.')
        sys.exit(1)

    results = {'ok': [], 'notfound': [], 'ambiguous': [], 'error': []}

    total = len(people)
    try:
        for idx, person in enumerate(people, 1):
            print(f'\n[{idx}/{total}] ΑΦΜ {person["afm"]}  {person["eponymo"]} {person["onoma"]}  '
                  f'—  {person["sxoleio"]}')
            status = process_person(driver, person, print)
            results[status].append(person['afm'])
            time.sleep(0.5)
    except KeyboardInterrupt:
        print('\n\n⚠ Διακόπηκε από τον χρήστη (Ctrl+C). Ό,τι έχει ήδη αποθηκευτεί παραμένει.')
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    print('\n' + '─' * 65)
    print(f'✓ Καταχωρήθηκαν επιτυχώς : {len(results["ok"])}')
    print(f'⚠ Δεν βρέθηκε ταίριασμα σχολείου : {len(results["notfound"])}')
    print(f'⚠ Διφορούμενα (πάνω από μία γραμμές ταίριαξαν) : {len(results["ambiguous"])}')
    print(f'✗ Σφάλματα εκτέλεσης : {len(results["error"])}')
    for key, label in (
        ('notfound', 'Δεν βρέθηκε ταίριασμα'),
        ('ambiguous', 'Διφορούμενα'),
        ('error', 'Σφάλματα'),
    ):
        if results[key]:
            print(f'  {label} — ΑΦΜ: ' + ' | '.join(results[key]))
    print('─' * 65)


if __name__ == '__main__':
    main()
