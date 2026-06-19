#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os, sys, time
import pandas as pd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CHECK_TITLE       = 'Αλλαγή Λειτουργικότητας'
CHECK_DESCRIPTION = 'Αυτόματη ενημέρωση λειτουργικότητας σχολικών μονάδων'
HAS_EMAIL         = False
CUSTOM_RUN        = True

BASE_URL      = 'https://app.myschool.sch.gr'
SEARCH_URL    = BASE_URL + '/Unit.list.myUnits.aspx'
TIME_TO_WAIT  = 15
CODE_FIELD_ID = 'ctl00_ContentData_txtRegistryNo_I'
LEIT_FIELD_ID = 'ctl00_ContentData_txtFunctionality_I'
SAVE_BTN_ID   = 'ctl00_ContentData_btnSave'

# Υποψήφια ονόματα στηλών
_CODE_CANDIDATES = (
    'Κωδικός ΥΠΑΙΘΑ',
    'ΚΩΔΙΚΟΣ Υ.ΠΑΙ.Θ.Α.',
    'Κωδικός Υπουργείου',
    'ΚΩΔΙΚΟΣ',
    'Κωδικός',
    'KOD', 'CODE',
    'Κωδ.',
    'ΚΩΔ.',
    'κωδικός',
    'ΚΩΔΙΚΟΣ ΣΧΟΛΕΙΟΥ',
    'Κωδικός Σχολείου',
)
_OLD_CANDIDATES = (
    'Λειτουργικότητα (ισχύουσα)',
    'Λειτουργικότητα',
    'ΛΕΙΤΟΥΡΓΙΚΟΤΗΤΑ',
    'Παλιά Λειτουργικότητα',
    'Παλιά Λειτ.',
    'Λειτ. (ισχύουσα)',
    'Λειτουργικότητα\n(ισχύουσα)',
)
_NEW_CANDIDATES = (
    'Νέα Λειτουργικότητα',
    'ΝΕΑ ΛΕΙΤΟΥΡΓΙΚΟΤΗΤΑ',
    'Νέα Λειτ.',
    'Νέα\nΛειτουργικότητα',
    'Λειτουργικότητα 2026-2027',
)


def _read_excel_smart(file_path):
    _KW = ('κωδ', 'cod', 'λειτουργ', 'αριθμ', 'α/α', 'ονομ', 'σχολ', 'ειδ')
    raw = pd.read_excel(file_path, header=None, dtype=str, nrows=15)
    header_row = 0
    for i, row in raw.iterrows():
        non_null = [v for v in row if pd.notna(v) and str(v) not in ('nan', 'None', '')]
        # Πρέπει να έχει τουλάχιστον 3 μη-κενά κελιά ώστε να είναι πραγματικό header
        if len(non_null) < 3:
            continue
        vals = ' '.join(str(v).lower() for v in non_null)
        if any(kw in vals for kw in _KW):
            header_row = i
            break
    return pd.read_excel(file_path, header=header_row, dtype=str)


def load_data(file_path, log=print):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext in ('.xlsx', '.xls'):
            df = _read_excel_smart(file_path)
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
        return []

    # Normalise headers: strip + flatten newlines
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

    code_col = _find(_CODE_CANDIDATES)
    old_col  = _find(_OLD_CANDIDATES)
    new_col  = _find(_NEW_CANDIDATES)

    missing = []
    if not code_col: missing.append('Κωδικός')
    if not old_col:  missing.append('Παλιά Λειτουργικότητα')
    if not new_col:  missing.append('Νέα Λειτουργικότητα')

    if missing:
        log('✗ Δεν βρέθηκαν υποχρεωτικές στήλες: ' + ', '.join(missing))
        log('  Διαθέσιμες: ' + str(list(df.columns)))
        return []

    log(f'  Στήλες: Κωδ.={code_col!r} | Παλιά={old_col!r} | Νέα={new_col!r}')

    name_col = None
    for c in ('Ονομασία', 'ΟΝΟΜΑΣΙΑ', 'Σχολείο', 'ΣΧΟΛΕΙΟ', 'Ονομ.'):
        if c in cols:
            name_col = c
            break

    records = []
    for _, row in df.iterrows():
        code = str(row.get(code_col, '')).strip()
        if code.endswith('.0'): code = code[:-2]
        raw_old = row.get(old_col, '')
        raw_new = row.get(new_col, '')
        old_leit = str(raw_old).strip() if raw_old and str(raw_old) not in ('nan','None','') else ''
        new_leit = str(raw_new).strip() if raw_new and str(raw_new) not in ('nan','None','') else ''
        onoma    = str(row.get(name_col, '')).strip() if name_col else ''
        for attr in ('old_leit', 'new_leit'):
            try:
                locals()[attr]  # just to trigger
            except Exception:
                pass
        try: old_leit = str(int(float(old_leit)))
        except (ValueError, TypeError): pass
        try: new_leit = str(int(float(new_leit)))
        except (ValueError, TypeError): pass
        if not code or code in ('nan','None',''): continue
        if not new_leit or new_leit in ('nan','None','','-'): continue
        records.append({'code': code, 'old_leit': old_leit, 'new_leit': new_leit, 'onoma': onoma})

    log(f'Φορτώθηκαν {len(records)} εγγραφές')
    return records


def _set_dxe_value(driver, eid, val):
    js = (
        "var inp=document.getElementById(arguments[0]);"
        "if(!inp)return false;"
        "inp.value=arguments[1];"
        "inp.dispatchEvent(new Event('change',{bubbles:true}));"
        "inp.dispatchEvent(new Event('input',{bubbles:true}));"
        "var b=arguments[0].replace(/_I$/,'');"
        "if(typeof aspxETextChanged==='function')aspxETextChanged(b);"
        "if(typeof aspxEValueChanged==='function')aspxEValueChanged(b);"
        "return true;"
    )
    return driver.execute_script(js, eid, val)


def _type_into_field(driver, el, val):
    driver.execute_script(
        "var e=arguments[0],v=arguments[1];e.focus();e.value='';"
        "for(var i=0;i<v.length;i++){"
        "var c=v[i];e.value+=c;"
        "e.dispatchEvent(new KeyboardEvent('keydown',{key:c,bubbles:true}));"
        "e.dispatchEvent(new KeyboardEvent('keypress',{key:c,bubbles:true}));"
        "e.dispatchEvent(new KeyboardEvent('keyup',{key:c,bubbles:true}));"
        "e.dispatchEvent(new Event('input',{bubbles:true}));}"
        "e.dispatchEvent(new Event('change',{bubbles:true}));",
        el, val)


def connect(log=print):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import config as _cfg

    opts = webdriver.ChromeOptions()
    opts.add_argument('--window-size=1400,900')
    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_experimental_option('excludeSwitches', ['enable-logging'])
    try:
        log('Αυτόματη εύρεση ChromeDriver...')
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
        except Exception as e1:
            log(f'webdriver-manager: {e1} -- δοκιμάζω χωρίς service...')
            driver = webdriver.Chrome(options=opts)
    except Exception as e:
        log(f'✗ Chrome: {e}')
        return None
    try:
        log('Σύνδεση στο MySchool...')
        driver.get(BASE_URL)
        time.sleep(2)
        if 'sso.sch.gr' in driver.current_url or 'login' in driver.current_url.lower():
            uf = WebDriverWait(driver,15).until(EC.presence_of_element_located((By.CSS_SELECTOR,'#username,input[name="username"]' )))
            uf.clear(); uf.send_keys(_cfg.MYSCHOOL_USER)
            pf = driver.find_element(By.CSS_SELECTOR,'#password,input[name="password"],input[type="password"]')
            pf.clear(); pf.send_keys(_cfg.MYSCHOOL_PASS)
            driver.find_element(By.CSS_SELECTOR,'button[type="submit"],input[type="submit"]').click()
            time.sleep(3)
        # Αλλαγή ακαδημαϊκού έτους σε 2026-2027
        log('Αλλαγή έτους σε 2026-2027...')
        driver.get(BASE_URL + '/Default.aspx'); time.sleep(2)
        try:
            yr_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ctl00_ContentData_cmbActiveAcadYear_I')))
            driver.execute_script(
                "var f=arguments[0]; f.value='2026-2027';"
                "f.dispatchEvent(new Event('change',{bubbles:true}));"
                "aspxETextChanged('ctl00_ContentData_cmbActiveAcadYear');",
                yr_field)
            time.sleep(2)
            log('  Έτος ΟΚ')
        except Exception as ey:
            log(f'  ⚠ Αλλαγή έτους: {ey}')
        driver.get(SEARCH_URL); time.sleep(3)
        log('Σύνδεση ΟΚ')
        return driver
    except Exception as e:
        log(f'✗ {e}')
        try: driver.quit()
        except Exception: pass
        return None


def run(ctx, driver, callback=None):
    log = callback or print
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    records = load_data(ctx.get('file_path'), log=log)
    if not records: return
    total = len(records)
    log(f'  {total} σχολεία προς ενημέρωση')
    modified_ok=[]; already_ok=[]; not_found=[]; multi_found=[]; mismatch=[]; failed=[]

    for idx, rec in enumerate(records, 1):
        code=rec['code']; old_leit=rec['old_leit']; new_leit=rec['new_leit']; onoma=rec['onoma']
        label=(onoma+' ('+code+')') if onoma else code
        log(f'\n[{idx}/{total}] {code}  {onoma}')
        try:
            driver.get(SEARCH_URL); time.sleep(2)
        except Exception as e:
            log(f'  ✗ Πλοήγηση: {e}'); failed.append(label); continue
        try:
            cf=WebDriverWait(driver,TIME_TO_WAIT).until(EC.presence_of_element_located((By.ID,CODE_FIELD_ID)))
            _type_into_field(driver,cf,code); time.sleep(1)
        except Exception as e:
            log(f'  ✗ Πεδίο κωδικού: {e}'); failed.append(label); continue
        try:
            # Try multiple selectors for search button; fallback to Enter
            _searched = False
            for _sel in ('a[href*="FillGrid"]', 'a.btn',
                         'a.hint_search', 'input[type="submit"]',
                         'a[id*="Search"]', '[id*="btnSearch"]'):
                try:
                    sb = WebDriverWait(driver, 4).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, _sel)))
                    driver.execute_script('arguments[0].click();', sb)
                    _searched = True
                    break
                except Exception:
                    pass
            if not _searched:
                from selenium.webdriver.common.keys import Keys
                driver.find_element(By.ID, CODE_FIELD_ID).send_keys(Keys.RETURN)
            time.sleep(3)
        except Exception as e:
            log(f'  ✗ Αναζήτηση: {e}'); failed.append(label); continue

        edits=driver.find_elements(By.XPATH,'//a[.//img[@alt="Διόρθωση"]]'  )
        log(f'  {len(edits)} αποτελέσματα')
        if len(edits)==0: log(f'  — Δεν βρέθηκε {code}'); not_found.append(label); continue
        if len(edits)>1:  log(f'  ⚠ {len(edits)} αποτελέσματα για {code}'); multi_found.append(label); continue

        try:
            driver.execute_script('arguments[0].click();',edits[0]); time.sleep(3)
        except Exception as e:
            log(f'  ✗ Άνοιγμα καρτέλας: {e}'); failed.append(label); continue
        try:
            lf=WebDriverWait(driver,TIME_TO_WAIT).until(EC.presence_of_element_located((By.ID,LEIT_FIELD_ID)))
            cur=(lf.get_attribute('value') or '').strip()
            try: cur_n=str(int(float(cur)))
            except: cur_n=cur
            log(f'  Τρέχουσα: {cur!r}')
        except Exception as e:
            log(f'  ✗ Πεδίο λειτουργικότητας: {e}'); failed.append(label); continue

        if cur_n==new_leit: log(f'  ✔ Ηδη {new_leit}'); already_ok.append(label); continue
        if old_leit and cur_n!=old_leit:
            log(f'  ⚠ Παλιά: {old_leit} | Βρέθηκε: {cur_n}')
            mismatch.append(label+' ['+old_leit+'>'+cur_n+']'); continue

        try:
            from selenium.webdriver.common.keys import Keys
            lf2 = driver.find_element(By.ID, LEIT_FIELD_ID)
            lf2.click()
            time.sleep(0.2)
            # Σβήνουμε με Ctrl+A + Delete για να καθαρίσει το DevExpress state
            lf2.send_keys(Keys.CONTROL + 'a')
            lf2.send_keys(Keys.DELETE)
            time.sleep(0.2)
            lf2.send_keys(new_leit)
            time.sleep(0.2)
            # Tab: trigger blur/aspxEValueChanged εσωτερικά
            lf2.send_keys(Keys.TAB)
            time.sleep(0.5)
            # Επαλήθευση ότι η τιμή καταχωρήθηκε
            lf3 = driver.find_element(By.ID, LEIT_FIELD_ID)
            actual = (lf3.get_attribute('value') or '').strip()
            try: actual_n = str(int(float(actual)))
            except: actual_n = actual
            if actual_n != new_leit:
                log(f'  ⚠ Τιμή δεν κρατήθηκε ({actual_n}), ξαναπροσπαθώ...')
                lf3.click(); lf3.send_keys(Keys.CONTROL+'a'); lf3.send_keys(new_leit)
                lf3.send_keys(Keys.TAB); time.sleep(0.5)
            log(f'  ✔ {cur} → {new_leit}')
        except Exception as e:
            log(f'  ✗ Εγγραφή τιμής: {e}'); failed.append(label); continue

        driver.execute_script('window.scrollTo(0,0);'); time.sleep(0.5)
        try:
            sv=WebDriverWait(driver,TIME_TO_WAIT).until(
                EC.element_to_be_clickable((By.ID, SAVE_BTN_ID)))
            # Anchor με __doPostBack: ValidatePage() + doPostBack
            valid = driver.execute_script(
                "return (typeof ValidatePage!=='function') || ValidatePage();")
            if not valid:
                log('  ✗ Αποθήκευση: ValidatePage() απέτυχε'); failed.append(label); continue
            driver.execute_script(
                "__doPostBack('ctl00$ContentData$btnSave','');")
            time.sleep(4)
            log('  ✔ Αποθήκευση'); modified_ok.append(label)
        except Exception as e:
            log(f'  ✗ Αποθήκευση: {e}'); failed.append(label)

    unchanged=not_found+already_ok+multi_found+mismatch+failed
    log('\n'+'='*50)
    log(f'  ✔  Αλλαχτήκαν: {len(modified_ok)}')
    log(f'  —  Αναλλοίωτα: {len(unchanged)}')
    if not_found:   log(f'     • Δεν βρέθηκαν: {len(not_found)}')
    if multi_found: log(f'     • Πολλαπλά: {len(multi_found)}')
    if mismatch:    log(f'     • Διαφορά τιμής: {len(mismatch)}')
    if already_ok:  log(f'     • Ηδη σωστή: {len(already_ok)}')
    if failed:      log(f'     • Σφάλματα: {len(failed)}')
    log('='*50)
    if modified_ok:
        log(f'\n✔ ΑΛΛΑΧΤΗΚΑΝ ({len(modified_ok)}):'+ ' '.join(modified_ok))
    if not_found:
        log('✗ ΔΕΝ ΒΡΕΘΗΚΑΝ: '+ ' | '.join(not_found))
    if multi_found:
        log('⚠ ΠΟΛΛΑΠΛΑ: '+ ' | '.join(multi_found))
    if mismatch:
        log('⚠ ΔΙΑΦΟΡΑ: '+ ' | '.join(mismatch))
    if already_ok:
        log('— ΗΔΗ ΣΩΣΤΑ: '+ ' | '.join(already_ok))
    if failed:
        log('✗ ΣΦΑΛΜΑ: '+ ' | '.join(failed))
