#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
dioikitiko_ergo_entry.py
=========================
Ε8. Διοικητικό Έργο — Αυτόματη καταχώρηση στο MySchool.

Ροή:
  1. Λήψη (ή εντοπισμός ήδη κατεβασμένου, από άλλον έλεγχο/εργαλείο)
     στατιστικού 2.2 — «Εκτεταμένα Στοιχεία Σχολ. Μον.». Χρησιμοποιείται
     ΜΟΝΟ για να βρεθεί το ακριβές λεκτικό κάθε σχολείου όπως εμφανίζεται
     στο MySchool (όχι για κάποιον κωδικό — δεν καταχωρείται πουθενά).
  2. Ανέβασμα Excel με στήλες Α.Μ., ΟΡΓΑΝΙΚΗ ΘΕΣΗ, ΤΟΠΟΘΕΤΗΣΗ ΓΙΑ ΑΣΚΗΣΗ
     ΔΙΟΙΚΗΤΙΚΟΥ ΕΡΓΟΥ.
  3. Για κάθε εγγραφή:
       • Αν οι δύο στήλες (Οργανική Θέση / Τοποθέτηση) ΔΕΝ συμφωνούν
         (έξυπνη σύγκριση — ίδια κανονικοποίηση με placements.py) →
         αγνοείται, συγκεντρώνεται στις Παρατηρήσεις.
       • Αν συμφωνούν, αναζητείται στο 2.2 το ακριβές λεκτικό MySchool
         του σχολείου· αν δεν βρεθεί (ή είναι διφορούμενο) → αγνοείται
         κι αυτή, συγκεντρώνεται στις Παρατηρήσεις.
       • Ό,τι απομείνει είναι «έτοιμο» για καταχώρηση.
  4. Εκτέλεση: αναζήτηση στο MySchool βάσει Α.Μ. Αν η αναζήτηση επιστρέψει
     πάνω από ένα αποτέλεσμα, επιλέγεται η γραμμή στην οποία η στήλη
     «Φορέας τοποθέτησης» ταιριάζει με το λεκτικό MySchool του βήματος 3
     — αν δεν βρεθεί μονοσήμαντη αντιστοιχία, η εγγραφή αγνοείται και
     συγκεντρώνεται. Στην καρτέλα: «Γραμματειακή Υποστήριξη» +
     01/09/2026–21/06/2027 + Παρατηρήσεις «ΠΔΕ-» → Αποδοχή → Αποθήκευση.
"""

import os
import sys
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Ε8 — Διοικητικό Έργο'
CHECK_DESCRIPTION = ('Αυτόματη καταχώρηση Διοικητικού Έργου (Γραμματειακή Υποστήριξη, '
                      '01/09/2026–21/06/2027) στο MySchool βάσει Α.Μ., με επαλήθευση '
                      'ΟΡΓΑΝΙΚΗΣ ΘΕΣΗΣ / ΤΟΠΟΘΕΤΗΣΗΣ και το 2.2.')
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL   = 'https://app.myschool.sch.gr'
SEARCH_URL = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

WORK_TYPE_TEXT = 'Γραμματειακή Υποστήριξη'   # ίδιο dropdown με την Ε6

# ── Σταθερές τιμές του Ε8 ────────────────────────────────────────────────────
FIXED_DATE_FROM = '01/09/2026'   # DXEditor5 — Ημ. από
FIXED_DATE_TO   = '21/06/2027'   # DXEditor6 — Ημ. έως
FIXED_NOTES     = 'ΠΔΕ-'         # DXEditor7 — Παρατηρήσεις

# Νέα στήλη που προστίθεται στο excel με το ακριβές λεκτικό MySchool (από 2.2)
WORDING_COL = 'Φορέας Τοποθέτησης (MySchool)'


# ── 2.2 — εντοπισμός / λήψη ───────────────────────────────────────────────────

def find_stat2():
    """Ψάχνει ήδη κατεβασμένο 2.1/2.2 στα downloads (reuse από άλλον έλεγχο).
    Επιστρέφει (path, date_label) ή (None, None)."""
    import placements as _pl
    return _pl._auto_find_stat2()


def download_stat2(username, password, dest_dir, browser='chrome', callback=None, force=False):
    """Κατεβάζει το 2.2 μέσω MySchoolDownloader. Επιστρέφει το path (ή None)."""
    from core.downloader import MySchoolDownloader
    log = callback or print
    dl = MySchoolDownloader(
        username=username, password=password, dest_dir=dest_dir,
        callback=log, reports=['2.2'], browser=browser,
        force=(['2.2'] if force else []),
    )
    results = dl.run()
    return results.get('2.2')


def _match_wording(name, lookup):
    """
    Επιστρέφει (orig_name, status): status ∈ {'exact','ambiguous','notfound'}.
    Ίδια λογική αντιστοίχισης με placements.match_school_code, αλλά
    επιστρέφει το πρωτότυπο λεκτικό (όπως στο MySchool) αντί για κωδικό.
    """
    import placements as _pl
    key = _pl._normalize_school_name(name)

    matches = lookup.get(key, [])
    if len(matches) == 1:
        return matches[0][1], 'exact'
    if len(matches) > 1:
        return None, 'ambiguous'

    prefix_matches = [v for k, v in lookup.items() if k.startswith(key + ' ')]
    flat = [item for sub in prefix_matches for item in sub]
    if len(flat) == 1:
        return flat[0][1], 'exact'
    if len(flat) > 1:
        return None, 'ambiguous'

    return None, 'notfound'


# ── Ανάγνωση αρχείου ─────────────────────────────────────────────────────────

def load_excel(file_path, log=print):
    """Διαβάζει το excel/csv και εντοπίζει τις στήλες Α.Μ. / Οργανική Θέση /
    Τοποθέτηση. Επιστρέφει (df, colmap) ή (None, None) αν λείπει στήλη."""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(file_path, dtype=str)
        else:
            df = None
            for enc in ('utf-8-sig', 'utf-8', 'iso-8859-7', 'cp1253'):
                try:
                    df = pd.read_csv(file_path, sep=';', encoding=enc, dtype=str)
                    break
                except Exception:
                    pass
            if df is None:
                df = pd.read_csv(file_path, dtype=str)
    except Exception as e:
        log(f'✗ Σφάλμα ανάγνωσης αρχείου: {e}')
        return None, None

    df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
    cols = set(df.columns)

    def _find(candidates):
        for c in candidates:
            if c in cols:
                return c
        low = {x.lower(): x for x in cols}
        for c in candidates:
            if c.lower() in low:
                return low[c.lower()]
        return None

    am_col = _find(('Α.Μ.', 'ΑΜ', 'Α.Μ', 'Αριθμός Μητρώου', 'ΑΡΙΘΜΟΣ ΜΗΤΡΩΟΥ', 'AM'))
    organiki_col = _find(('ΟΡΓΑΝΙΚΗ ΘΕΣΗ', 'Οργανική Θέση', 'ΟΡΓΑΝΙΚΗ'))
    topothetisi_col = _find((
        'ΤΟΠΟΘΕΤΗΣΗ ΓΙΑ ΑΣΚΗΣΗ ΔΙΟΙΚΗΤΙΚΟΥ ΕΡΓΟΥ',
        'Τοποθέτηση για Άσκηση Διοικητικού Έργου',
        'ΤΟΠΟΘΕΤΗΣΗ ΔΙΟΙΚΗΤΙΚΟΥ ΕΡΓΟΥ',
        'ΤΟΠΟΘΕΤΗΣΗ',
    ))

    missing = []
    if not am_col: missing.append('Α.Μ.')
    if not organiki_col: missing.append('ΟΡΓΑΝΙΚΗ ΘΕΣΗ')
    if not topothetisi_col: missing.append('ΤΟΠΟΘΕΤΗΣΗ ΓΙΑ ΑΣΚΗΣΗ ΔΙΟΙΚΗΤΙΚΟΥ ΕΡΓΟΥ')
    if missing:
        log('✗ Δεν βρέθηκαν στήλες: ' + ', '.join(missing))
        log('  Διαθέσιμες: ' + str(list(df.columns)))
        return None, None

    log(f'  Στήλες: Α.Μ.={am_col!r} | Οργανική={organiki_col!r} | Τοποθέτηση={topothetisi_col!r}')
    return df, {'am': am_col, 'organiki': organiki_col, 'topothetisi': topothetisi_col}


def prepare_records(file_path, stat2_path, log=print):
    """
    Πλήρες βήμα προετοιμασίας — διαβάζει το excel, ελέγχει ΟΡΓΑΝΙΚΗ ΘΕΣΗ vs
    ΤΟΠΟΘΕΤΗΣΗ, αντιστοιχίζει λεκτικό MySchool μέσω 2.2.

    Επιστρέφει dict:
      ready    -> [{'am','organiki','topothetisi','wording'}, ...]  έτοιμες
      mismatch -> [{'am','organiki','topothetisi'}, ...]             διαφορά στηλών
      notfound -> [{'am','organiki','topothetisi'}, ...]             άγνωστο στο 2.2
      df_out   -> pandas DataFrame (αντίγραφο εισόδου + νέα στήλη λεκτικού)
    ή None αν απέτυχε η ανάγνωση του αρχείου.
    """
    import placements as _pl

    df, colmap = load_excel(file_path, log=log)
    if df is None:
        return None

    lookup = {}
    if stat2_path:
        lookup = _pl.build_school_lookup(stat2_path)
        log(f'  2.2: {len(lookup)} σχολεία στο λεξικό')
    else:
        log('  ⚠ Δεν βρέθηκε αρχείο 2.2 — όλες οι εγγραφές θα θεωρηθούν άγνωστες.')

    df[WORDING_COL] = ''

    ready, mismatch, notfound = [], [], []
    for idx, row in df.iterrows():
        am = str(row.get(colmap['am'], '')).strip()
        if not am or am in ('nan', 'None', ''):
            continue

        organiki    = str(row.get(colmap['organiki'], '')).strip()
        topothetisi = str(row.get(colmap['topothetisi'], '')).strip()
        rec = {'am': am, 'organiki': organiki, 'topothetisi': topothetisi}

        organiki_ok = organiki and organiki not in ('nan', 'None')
        same = organiki_ok and (
            _pl._normalize_school_name(organiki) == _pl._normalize_school_name(topothetisi))
        if not same:
            mismatch.append(rec)
            continue

        wording, status = _match_wording(organiki, lookup)
        if status != 'exact':
            rec['_lookup_status'] = status
            notfound.append(rec)
            continue

        rec['wording'] = wording
        df.at[idx, WORDING_COL] = wording
        ready.append(rec)

    log(f'  ✓ Έτοιμες: {len(ready)}  |  ⚠ Διαφορά στηλών: {len(mismatch)}  |  '
        f'✗ Άγνωστο σχολείο (2.2): {len(notfound)}')

    return {'ready': ready, 'mismatch': mismatch, 'notfound': notfound, 'df_out': df}


def export_augmented(df_out, out_path):
    """Αποθηκεύει το επεξεργασμένο excel (με τη νέα στήλη λεκτικού MySchool)."""
    df_out.to_excel(out_path, index=False)
    return out_path


# ── Βοηθητικές Selenium ───────────────────────────────────────────────────────

STRIKE_INTERVAL = 0.3


def _send_keys_slow(element, text, delay=STRIKE_INTERVAL):
    for char in str(text):
        element.send_keys(char)
        time.sleep(delay)


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


def _select_dxe_combo(driver, base_id, text):
    """Επιλέγει τιμή από DevExpress ComboBox με πληκτρολόγηση + TAB."""
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
    except Exception:
        return False


# ── Σύνδεση (καλείται από το UI) ─────────────────────────────────────────────

def connect(log=print):
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


# ── Εκτέλεση (καλείται από το UI) ────────────────────────────────────────────

def run(ctx, driver, callback=None):
    """
    ctx αναμένεται να περιέχει το αποτέλεσμα του prepare_records():
      ctx['ready'], ctx['mismatch'], ctx['notfound']  (λίστες dict)
    (τα mismatch/notfound έχουν ήδη αγνοηθεί πριν φτάσουμε εδώ — απλά
    εμφανίζονται στην τελική αναφορά).
    """
    log = callback or print

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    ready    = ctx.get('ready') or []
    mismatch = ctx.get('mismatch') or []
    notfound = ctx.get('notfound') or []

    total = len(ready)
    if not total:
        log('Καμία έτοιμη εγγραφή για καταχώρηση.')
    else:
        log(f'  {total} εγγραφές προς καταχώρηση  |  Ημ. από: {FIXED_DATE_FROM}  |  '
            f'Ημ. έως: {FIXED_DATE_TO}')

    ok = 0
    failed = []
    not_found_myschool = []
    ambiguous_myschool = []

    for idx, record in enumerate(ready, 1):
        am      = record['am']
        wording = record.get('wording', '')
        log(f'\n[{idx}/{total}] Α.Μ.: {am}  |  {wording}')

        # ── Σελίδα αναζήτησης ─────────────────────────────────────────────
        driver.get(SEARCH_URL)
        time.sleep(2)

        # ── Συμπλήρωση Α.Μ. ───────────────────────────────────────────────
        try:
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
        except Exception as e:
            log(f'  ✗ Α.Μ. field: {e}')
            failed.append(am)
            continue

        # ── Αναζήτηση ─────────────────────────────────────────────────────
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}')
            failed.append(am)
            continue

        hours = ''

        # ── Εύρεση εγγραφής — αν πάνω από μία, ταιριάζουμε με το λεκτικό
        #    MySchool (Φορέας Τοποθέτησης) από το 2.2 ───────────────────────
        edit_links = driver.find_elements(
            By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log('  ✗ Κανένα αποτέλεσμα')
            not_found_myschool.append(am)
            continue

        target = None
        if len(edit_links) == 1:
            target = edit_links[0]
        else:
            matched = []
            for link in edit_links:
                try:
                    row = link.find_element(By.XPATH, './ancestor::tr[1]')
                    if wording and wording.upper() in row.text.upper():
                        matched.append(link)
                except Exception:
                    pass
            if len(matched) == 1:
                target = matched[0]
                log(f'  Βρέθηκε μονοσήμαντα: {wording}')
            else:
                log(f'  ⚠ {len(edit_links)} αποτελέσματα, {len(matched)} ταίριαξαν με '
                    f'"{wording}" — αγνοείται')
                ambiguous_myschool.append(am)
                continue

        # ── Άνοιγμα καρτέλας ──────────────────────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Καρτέλα: {e}')
            failed.append(am)
            continue

        # ── Ώρες από καρτέλα (μετά άνοιγμα) ────────────────────────────
        try:
            hours_el = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.ID, 'ctl00_ContentData_txtAvailableHoursForUnit_I')))
            hours = hours_el.get_attribute('value').strip()
            log(f'  Διαθέσιμες ώρες: {hours}')
        except Exception as e:
            log(f'  ⚠ Ώρες: {e}')

        # ── Σταυρός Προσθήκης ─────────────────────────────────────────────
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
            log(f'  ✗ Σταυρός: {e}')
            failed.append(am)
            continue

        # ── Dropdown "Γραμματειακή Υποστήριξη" ───────────────────────────
        combo_base = ('ctl00_ContentData_gridEmplDet_'
                      'editnew_2_cmbWorkHoursDetailsType')
        try:
            ok_c = _select_dxe_combo(driver, combo_base, WORK_TYPE_TEXT)
            log(f'  {"✓" if ok_c else "⚠"} Τύπος: {WORK_TYPE_TEXT}')
        except Exception as e:
            log(f'  ⚠ Dropdown: {e}')

        # ── Ώρες φόρμα (DXEditor4) ─────────────────────────────────────────
        if hours:
            try:
                hours_inp = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.ID, 'ctl00_ContentData_gridEmplDet_DXEditor4_I')))
                driver.execute_script('arguments[0].click(); arguments[0].focus();', hours_inp)
                time.sleep(0.5)
                driver.execute_script("""
                    var el = arguments[0];
                    var val = arguments[1];
                    el.value = '';
                    el.value = val;
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                    aspxEValueChanged('ctl00_ContentData_gridEmplDet_DXEditor4');
                """, hours_inp, hours)
                log(f'  ✓ Ώρες: {hours}')
            except Exception as e:
                log(f'  ⚠ Ώρες φόρμα: {e}')

        # ── Ημ. από (DXEditor5) — σταθερή 01/09/2026 ───────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor5_I', FIXED_DATE_FROM)
            log(f'  ✓ Ημ. από: {FIXED_DATE_FROM}')
        except Exception as e:
            log(f'  ⚠ Ημ. από: {e}')

        # ── Ημ. έως (DXEditor6) — σταθερή 21/06/2027 ────────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor6_I', FIXED_DATE_TO)
            log(f'  ✓ Ημ. έως: {FIXED_DATE_TO}')
        except Exception as e:
            log(f'  ⚠ Ημ. έως: {e}')

        # ── Παρατηρήσεις (DXEditor7) — σταθερό «ΠΔΕ-» ──────────────────────
        try:
            _set_dxe_value(driver,
                'ctl00_ContentData_gridEmplDet_DXEditor7_I', FIXED_NOTES)
            log(f'  ✓ Παρατηρήσεις: {FIXED_NOTES}')
        except Exception as e:
            log(f'  ⚠ Παρατηρήσεις: {e}')

        time.sleep(0.5)

        # ── Αποδοχή (πράσινο τικ) ────────────────────────────────────────
        try:
            accept_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//img[@alt="Αποδοχή"]')))
            driver.execute_script('arguments[0].click();', accept_btn)
            time.sleep(2)
            log('  ✓ Αποδοχή')
        except Exception as e:
            log(f'  ✗ Αποδοχή: {e}')
            failed.append(am)
            continue

        # ── Αποθήκευση ────────────────────────────────────────────────────
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
            failed.append(am)

    # ── Τελική αναφορά (συμπεριλαμβανομένων ό,τι αγνοήθηκε πριν την
    #    εκτέλεση, από το prepare_records) ───────────────────────────────────
    log('\n' + '─'*50)
    log(f'Καταχωρήθηκαν: {ok}   |   Σφάλματα εκτέλεσης: {len(failed)}')
    log(f'Δεν βρέθηκαν στο MySchool: {len(not_found_myschool)}   |   '
        f'Διφορούμενα (Φορέας Τοποθέτησης): {len(ambiguous_myschool)}')
    log('\nΠΑΡΑΤΗΡΗΣΕΙΣ (αγνοήθηκαν πριν την εκτέλεση):')
    log(f'  • Διαφορά ΟΡΓΑΝΙΚΗΣ ΘΕΣΗΣ / ΤΟΠΟΘΕΤΗΣΗΣ: {len(mismatch)}')
    if mismatch:
        log('    Α.Μ.: ' + ' | '.join(r['am'] for r in mismatch))
    log(f'  • Άγνωστο/διφορούμενο σχολείο στο 2.2: {len(notfound)}')
    if notfound:
        log('    Α.Μ.: ' + ' | '.join(r['am'] for r in notfound))
    if failed:
        log('  • Σφάλματα εκτέλεσης — Α.Μ.: ' + ' | '.join(failed))
    if not_found_myschool:
        log('  • Δεν βρέθηκαν στο MySchool — Α.Μ.: ' + ' | '.join(not_found_myschool))
    if ambiguous_myschool:
        log('  • Διφορούμενα στο MySchool — Α.Μ.: ' + ' | '.join(ambiguous_myschool))
    log('─'*50)
