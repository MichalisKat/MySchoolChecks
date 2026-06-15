#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
termination.py
==============
Αυτόματος τερματισμός τοποθετήσεων εκπαιδευτικών στο MySchool.

Ροή:
1. Επιλογή Excel/CSV (στήλη: Α.Μ.)
2. Σύνδεση στο MySchool
3. Για κάθε Α.Μ.:
   α) Αναζήτηση στη λίστα εκπαιδευτικών
   β) Αν δεν βρεθεί → επόμενος (καταγραφή)
   γ) Αν 1 αποτέλεσμα → κλικ γρανάζι
   δ) Αν πολλαπλά → popup επιλογής
4. Στην καρτέλα: έλεγχος dtDutyStopDate
   - 21/6/2026 → αφήνεται ως έχει
   - 31/8/2026 → αλλάζεται αυτόματα σε 21/6/2026
   - Άλλο      → ρωτάει χρήστη
5. Scroll πάνω → Αποθήκευση → επόμενος
6. Σύνοψη: 2 λίστες (αλλαχτήκαν / δεν αλλαχτήκαν)
"""

import os
import sys
import time
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Τερματισμός Τοποθετήσεων'
CHECK_DESCRIPTION = 'Αυτόματος τερματισμός τοποθετήσεων εκπαιδευτικών στο MySchool'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL   = 'https://app.myschool.sch.gr'
SEARCH_URL = BASE_URL + '/Worker.list.myEmplUnit.aspx'
TIME_TO_WAIT = 15

DEFAULT_STOP_DATE = '21/6/2026'   # Προεπιλεγμένη ημ. λήξης
AUTO_REPLACE_DATE = '31/8/2026'   # Αυτόματη αντικατάσταση αυτής

# ID πεδίου Α.Μ. στη σελίδα αναζήτησης εκπαιδευτικών
AM_FIELD_ID = 'ctl00_ContentData_txtRegistryNo_I'


# ── Ανάγνωση αρχείου ─────────────────────────────────────────────────────────

def load_data(file_path, log=print):
    """Φορτώνει Α.Μ. (+ προαιρετικά σχολείο) από Excel/CSV."""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(file_path, dtype=str)
        else:
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

    cols = {c.strip(): c for c in df.columns}

    am_col = None
    for candidate in ('Α.Μ.', 'ΑΜ', 'Α.Μ', 'AM', 'Αρ.Μητρώου', 'ΑΜΕΚΠ', 'Α.Μ. ΕΚΠ/ΚΟΥ'):
        if candidate in cols:
            am_col = cols[candidate]
            break

    school_col = None
    for candidate in ('Ονομασία Σχολείου', 'ΣΧΟΛΕΙΟ', 'Σχολείο', 'ΚΩΔ. ΣΧΟΛΕΙΟΥ'):
        if candidate in cols:
            school_col = cols[candidate]
            break

    name_col = None
    for candidate in ('Επώνυμο', 'ΕΠΩΝΥΜΟ', 'Ονοματεπώνυμο'):
        if candidate in cols:
            name_col = cols[candidate]
            break

    if not am_col:
        log(f'Δεν βρέθηκε στήλη Α.Μ. Διαθέσιμες: {list(cols.keys())}')
        return []

    records = []
    for _, row in df.iterrows():
        am     = str(row.get(am_col, '')).strip()
        school = str(row.get(school_col, '')).strip() if school_col else ''
        name   = str(row.get(name_col,  '')).strip() if name_col  else ''
        if am and am not in ('nan', 'None', ''):
            records.append({'am': am, 'school': school, 'name': name})

    log(f'Φορτώθηκαν {len(records)} εγγραφές')
    return records


# ── Βοηθητικές ────────────────────────────────────────────────────────────────

def _normalize_date(s):
    """Κανονικοποιεί ημερομηνία — αφαιρεί μηδενικά πρόθεμα ΗΗ/ΜΜ."""
    s = (s or '').strip()
    parts = s.split('/')
    if len(parts) == 3:
        try:
            d = int(parts[0])
            m = int(parts[1])
            y = parts[2].strip()
            return f'{d}/{m}/{y}'
        except Exception:
            pass
    return s


def _set_dxe_value(driver, element_id, value):
    """Θέτει τιμή σε DevExpress date/text field."""
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


# ── Σύνδεση ───────────────────────────────────────────────────────────────────

def connect(log=print):
    """Σύνδεση στο MySchool — επιστρέφει driver ή None."""
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


# ── Κύρια εκτέλεση ────────────────────────────────────────────────────────────

def run(ctx, driver, callback=None, ask_user=None):
    """
    Τερματισμός τοποθετήσεων.

    Παράμετροι:
      ctx       : dict με 'file_path' και 'date' (ημ. λήξης)
      driver    : Selenium WebDriver (ήδη συνδεδεμένος)
      callback  : log(msg) — εκτυπώνει στο UI
      ask_user  : ask_user(title, prompt, options) → str|None
                  Αν options=None  → text input (επιστρέφει string ή None)
                  Αν options=list  → επιλογή από λίστα ('1'–'N' ή '0' για παράλειψη)
    """
    log  = callback or print
    _ask = ask_user or _fallback_ask

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    file_path = ctx.get('file_path')
    stop_date = ctx.get('date', DEFAULT_STOP_DATE)

    records = load_data(file_path, log=log)
    if not records:
        return

    total = len(records)
    log(f'  {total} εγγραφές  |  Ημ. Λήξης: {stop_date}')

    stop_norm    = _normalize_date(stop_date)
    replace_norm = _normalize_date(AUTO_REPLACE_DATE)

    not_found   = []   # 0 αποτελέσματα
    already_ok  = []   # Ημ. ήδη σωστή
    modified_ok = []   # Αλλαχτήκαν & αποθηκεύτηκαν
    failed      = []   # Τεχνικά σφάλματα
    skipped     = []   # Παραλείφθηκαν από χρήστη

    for rec_idx, record in enumerate(records, 1):
        am     = record['am']
        school = record.get('school', '')
        name   = record.get('name', '')
        label  = f'{name} ({am})' if name else am
        log(f'\n[{rec_idx}/{total}] Α.Μ.: {am}  {("| " + school) if school else ""}')

        # ── Πλοήγηση στη σελίδα αναζήτησης ───────────────────────────────
        try:
            driver.get(SEARCH_URL)
            time.sleep(2)
        except Exception as e:
            log(f'  ✗ Πλοήγηση: {e}')
            failed.append(label)
            continue

        # ── Συμπλήρωση Α.Μ. ───────────────────────────────────────────────
        try:
            am_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.ID, AM_FIELD_ID)))
            driver.execute_script("""
                var el  = arguments[0];
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
            failed.append(label)
            continue

        # ── Κουμπί αναζήτησης ─────────────────────────────────────────────
        try:
            search_link = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.hint_search')))
            driver.execute_script('arguments[0].click();', search_link)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}')
            failed.append(label)
            continue

        # ── Εύρεση γρανάζι (Διόρθωση) ────────────────────────────────────
        edit_links = driver.find_elements(
            By.XPATH, '//a[.//img[@alt="Διόρθωση"]]')
        log(f'  {len(edit_links)} αποτέλεσμα(-τα)')

        if not edit_links:
            log(f'  — Δεν βρέθηκε')
            not_found.append(label)
            continue

        # ── Επιλογή εγγραφής ──────────────────────────────────────────────
        if len(edit_links) == 1:
            target = edit_links[0]
        else:
            rows_text = []
            for lnk in edit_links:
                try:
                    row = lnk.find_element(By.XPATH, './ancestor::tr[1]')
                    rows_text.append(row.text.strip().replace('\n', '  '))
                except Exception:
                    rows_text.append(f'Εγγραφή {len(rows_text) + 1}')

            choice = _ask(
                'Πολλαπλά Σχολεία',
                f'Α.Μ. {am}\nΒρέθηκαν {len(edit_links)} αποτελέσματα.\nΕπίλεξε (0 = παράλειψη):',
                rows_text
            )

            if choice is None or str(choice).strip() == '0':
                log(f'  ⏭ Παράλειψη από χρήστη')
                skipped.append(label)
                continue

            try:
                pick = int(str(choice).strip()) - 1
                if not (0 <= pick < len(edit_links)):
                    raise ValueError('εκτός ορίων')
                target = edit_links[pick]
                log(f'  Επιλέχθηκε: {rows_text[pick][:60]}')
            except (ValueError, TypeError):
                log(f'  ✗ Άκυρη επιλογή: {choice}')
                skipped.append(label)
                continue

        # ── Άνοιγμα καρτέλας (κλικ γρανάζι) ─────────────────────────────
        try:
            driver.execute_script('arguments[0].click();', target)
            time.sleep(3)
            log('  Καρτέλα ανοιχτή')
        except Exception as e:
            log(f'  ✗ Άνοιγμα καρτέλας: {e}')
            failed.append(label)
            continue

        # ── Ανάγνωση ημ. λήξης (dtDutyStopDate) ──────────────────────────
        try:
            stop_field = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.presence_of_element_located(
                    (By.ID, 'ctl00_ContentData_dtDutyStopDate_I')))
            current_val  = (stop_field.get_attribute('value') or '').strip()
            current_norm = _normalize_date(current_val)
            log(f'  Ημ. Λήξης: "{current_val}"')
        except Exception as e:
            log(f'  ✗ Δεν βρέθηκε πεδίο Ημ. Λήξης: {e}')
            failed.append(label)
            continue

        # ── Λογική ελέγχου / τροποποίησης ────────────────────────────────
        needs_save = False

        if current_norm == stop_norm:
            log(f'  ✓ Ήδη {stop_date} — χωρίς αλλαγή')
            already_ok.append(label)
            continue

        elif current_norm == replace_norm:
            ok = _set_dxe_value(driver,
                'ctl00_ContentData_dtDutyStopDate_I', stop_date)
            if ok:
                time.sleep(0.5)
                log(f'  ✓ Αλλαγή: {current_val} → {stop_date}')
                needs_save = True
            else:
                log(f'  ✗ Αποτυχία εγγραφής τιμής')
                failed.append(label)
                continue

        else:
            choice = _ask(
                'Απροσδόκητη Ημερομηνία',
                (f'Α.Μ. {am}\n'
                 f'Ημ. Λήξης: "{current_val}"\n\n'
                 f'Τι να κάνω;'),
                [
                    'Παράλειψη',
                    f'Αλλαγή σε {stop_date}',
                    'Εισαγωγή νέας ημ/νίας',
                ]
            )

            if choice is None or str(choice).strip() in ('0', '1'):
                log(f'  ⏭ Παράλειψη — {current_val}')
                skipped.append(label)
                continue

            elif str(choice).strip() == '2':
                _set_dxe_value(driver,
                    'ctl00_ContentData_dtDutyStopDate_I', stop_date)
                time.sleep(0.5)
                log(f'  ✓ Αλλαγή: {current_val} → {stop_date}')
                needs_save = True

            elif str(choice).strip() == '3':
                new_date = _ask(
                    'Νέα Ημερομηνία',
                    f'Α.Μ. {am}\nΕισάγετε ημερομηνία (π.χ. 21/6/2026):',
                    None
                )
                if not new_date or not str(new_date).strip():
                    log(f'  ⏭ Παράλειψη (χωρίς ημερομηνία)')
                    skipped.append(label)
                    continue
                new_date = str(new_date).strip()
                _set_dxe_value(driver,
                    'ctl00_ContentData_dtDutyStopDate_I', new_date)
                time.sleep(0.5)
                log(f'  ✓ Αλλαγή: {current_val} → {new_date}')
                needs_save = True

            else:
                log(f'  ⏭ Παράλειψη')
                skipped.append(label)
                continue

        if not needs_save:
            continue

        # ── Scroll στην κορυφή ────────────────────────────────────────────
        driver.execute_script('window.scrollTo(0, 0);')
        time.sleep(0.5)

        # ── Αποθήκευση ────────────────────────────────────────────────────
        try:
            save_btn = WebDriverWait(driver, TIME_TO_WAIT).until(
                EC.element_to_be_clickable(
                    (By.ID, 'ctl00_ContentData_btnSave')))
            driver.execute_script('arguments[0].click();', save_btn)
            time.sleep(3)
            log('  ✓ Αποθήκευση')
            modified_ok.append(label)
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}')
            failed.append(label)

    # ── Σύνοψη ────────────────────────────────────────────────────────────────
    unchanged = not_found + already_ok + failed + skipped

    log(f'\n{"═" * 55}')
    log(f'  ΑΠΟΤΕΛΕΣΜΑΤΑ ΤΕΡΜΑΤΙΣΜΟΥ ΤΟΠΟΘΕΤΗΣΕΩΝ')
    log(f'{"─" * 55}')
    log(f'  ✓  Αλλαχτήκαν & αποθηκεύτηκαν  : {len(modified_ok)}')
    log(f'  —  Δεν αλλαχτήκαν (σύνολο)      : {len(unchanged)}')
    log(f'       εκ των οποίων:')
    log(f'       • Ήδη σωστή ημερομηνία      : {len(already_ok)}')
    log(f'       • Δεν βρέθηκαν              : {len(not_found)}')
    log(f'       • Σφάλματα                  : {len(failed)}')
    log(f'       • Παραλείφθηκαν             : {len(skipped)}')
    log(f'{"═" * 55}')

    # ── Αναλυτικές λίστες (μετά από ερώτηση) ─────────────────────────────────
    show = _ask(
        'Αναλυτικές Λίστες',
        'Θέλεις να δεις αναλυτικές λίστες;',
        [
            f'Ναι — εμφάνιση και των δύο λιστών',
            f'Μόνο αυτοί που αλλαχτήκαν ({len(modified_ok)})',
            f'Μόνο αυτοί που δεν αλλαχτήκαν ({len(unchanged)})',
            'Όχι',
        ]
    )

    show = str(show).strip() if show else '4'

    if show in ('1', '2'):
        log(f'\n✓ ΑΛΛΑΧΤΗΚΑΝ ({len(modified_ok)}):')
        if modified_ok:
            for entry in modified_ok:
                log(f'   {entry}')
        else:
            log('   (κανένας)')

    if show in ('1', '3'):
        log(f'\n— ΔΕΝ ΑΛΛΑΧΤΗΚΑΝ ({len(unchanged)}):')
        if unchanged:
            for entry in unchanged:
                log(f'   {entry}')
        else:
            log('   (κανένας)')


# ── Fallback ask (για terminal / testing) ─────────────────────────────────────

def _fallback_ask(title, prompt, options=None):
    """Fallback: χρήση terminal input όταν δεν υπάρχει UI callback."""
    print(f'\n{"=" * 50}')
    print(f'[{title}]')
    print(prompt)
    if options:
        for i, opt in enumerate(options, 1):
            print(f'  {i}. {opt}')
        print('  0. Παράλειψη')
        return input('Επιλογή: ').strip()
    else:
        return input('Εισαγωγή: ').strip() or None
