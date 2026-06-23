"""
core/downloader.py
══════════════════
Αυτόματο κατέβασμα αρχείων από το MySchool μέσω Selenium.

Αρχεία που κατεβαίνουν:
  topoth — Τοποθετήσεις εκπαιδευτικών
  2.1    — Κατάλογος σχολικών μονάδων
  4.1    — Οργανικές τοποθετήσεις εκπαιδευτικών
  4.2    — Αποσπασμένοι εκπαιδευτικοί
  4.8    — Ωράριο εκπαιδευτικών
  4.9    — Παρόντες εκπαιδευτικοί
  4.11   — Λεπτομέρειες μείωσης ωραρίου
  4.12   — Λεπτομέρειες συμπλήρωσης ωραρίου
  4.16   — Αιτιολόγηση απουσίας εκπαιδευτικών
  4.20   — Άδειες Άνευ Αποδοχών
  4.21   — Άδειες (πλην ΑΑ)
  8.2    — Επιβεβαίωση δεδομένων σχολείων
  4.26   — Αδυνατούντες ανά ειδικότητα (4.26 ΠΔΕ / 4.27 ΔΔΕ)
  5.3    — Αποτύπωση νηπιαγωγείων
  5.4    — Αποτύπωση δημοτικών

Χρήση:
  from core.downloader import MySchoolDownloader
  dl = MySchoolDownloader(username, password, dest_dir, callback=print)
  results = dl.run()   # dict {report_id: path_or_None}
"""

import os, time, shutil, glob, zipfile
from datetime import datetime
from pathlib import Path


# ── Ρυθμίσεις ανά αρχείο ────────────────────────────────────────────────────
# (report_id, label, url_path, fname_base, wait_search, wait_dl, direct_export)
# direct_export=True → σελίδα χωρίς αναζήτηση/grid, απευθείας εξαγωγή
REPORTS = [
    ('topoth', 'Τοποθετήσεις εκπαιδευτικών',    '/Worker.list.myEmplUnit.aspx',                                               'Topothetiseis',       30, 90, False, 'a.hint_search', '#ctl00_ContentData_gridResults_StatusBar_btnExport'),
    ('2.1',    'Κατάλογος σχολείων',             '/Statistics/Management.stat.infoUnits.aspx?parentId=3',                     'gridResults',         30, 60, False),
    ('2.2',    'Εκτεταμένα Στοιχεία Σχολ. Μον.', '/Statistics/Management.stat.infoAdvUnits.aspx?parentId=3',                  'stat2_2',             30, 90, False),
    # 3.1: Πριν την αναζήτηση πρέπει να τσεκαριστούν τα checkboxes ομαδοποίησης
    ('3.1',  'Κατανομή μαθητών ανά τάξη',      '/Statistics/Management.stat.sumStudGroupGP.aspx?parentId=4',                 'stat3_1',             60, 120, False, None, None,
             ['Είδος σχολείου', 'Τύπος σχολείου', 'Σχολική Μονάδα', 'Τάξη']),
    ('4.1',  'Οργανικές τοποθετήσεις',         '/Statistics/Management.stat.wrkCatalogue.aspx?parentId=5',                   'stat4_1',             60, 60, False),
    ('4.2',  'Αποσπασμένοι εκπαιδευτικοί',     '/Statistics/Management.stat.wrkDetachedCatalogue.aspx?parentId=5',           'stat4_2',             60, 60, False),
    ('4.8',  'Ωράριο εκπαιδευτικών',           '/Statistics/Management.stat.TchHoursCatalogue.aspx?parentId=5',              '4.8_Ωραριο',          30, 40, False),
    ('4.9',  'Παρόντες εκπαιδευτικοί',         '/Statistics/Management.stat.TchHoursCatalogueUnqWrk.aspx?parentId=5',       '4.9_Παροντες',        69, 46, False),
    ('4.11', 'Μείωση ωραρίου',                  '/Statistics/Management.stat.MeiwseisCatalogue.aspx?parentId=5',              '4.11_Meiwseis',       30, 40, False),
    ('4.12', 'Συμπλήρωση ωραρίου',             '/Statistics/Management.stat.SymplirwseisCatalogue.aspx?parentId=5',          '4.12_Symplirwseis',   30, 40, False),
    ('4.16', 'Αιτιολόγηση απουσίας',           '/Statistics/Management.stat.wrkAbsenteesFromUnitCatalogue.aspx?parentId=5', 'stat4_16',            60, 60, False),
    ('4.20', 'Άδειες Άνευ Αποδοχών',           '/Statistics/Management.stat.NoCalcSrvLeaveStats.aspx?parentId=5',            '4.20_Adeies_AA',      30, 40, False),
    ('4.21', 'Άδειες (πλην ΑΑ)',               '/Statistics/Management.stat.LeavesPerPFU.aspx?parentId=5',                   '4.21_Adeies',         30, 40, False, None, None, None,
             ['ctl00_cntStats_cbpSearch_sus_cbmPanel_cmbCalendarYear_LBI1T1',
              'ctl00_cntStats_cbpSearch_sus_cbmPanel_cmbCalendarYear_LBI2T1']),
    ('8.2',  'Επιβεβαίωση δεδομένων',          '/Statistics/Management.stat.LastConfirmDateUnits.aspx?parentId=9',           '8.2_Epivevaiwsi',     30, 60, False),
    ('4.26', 'Αδυνατούντες ανά ειδικότητα',    '/Statistics/Management.stat.wrkInCapable.aspx?parentId=5',                    '4.26_Adynatountes',   30, 60, False),
    ('5.3',  'Αποτύπωση νηπιαγωγείων (2026-27)', '/Statistics/Management.stat.ZHTHSHPE60.aspx?parentId=6',                      'stat5_3',             30, 60, False, None, None, None, None, '2026-2027'),
    ('5.4',  'Αποτύπωση δημοτικών (2026-27)',  '/Statistics/Management.stat.ZHTHSHPE70.aspx?parentId=6',                      'stat5_4',             30, 60, False, None, None, None, None, '2026-2027'),
]

# Mapping: report_id → prefix ονόματος αποθηκευμένου αρχείου
FILE_PREFIX_MAP = {
    'topoth': 'Topothetiseis',
    '2.1'  : 'gridResults',
    '2.2'  : 'stat2_2',
    '3.1'  : 'stat3_1',
    '4.1'  : 'stat4_1',
    '4.2'  : 'stat4_2',
    '4.8' : '4.8_',
    '4.9' : '4.9_',
    '4.11': '4.11_',
    '4.12': '4.12_',
    '4.16': 'stat4_16',
    '4.20': '4.20_',
    '4.21': '4.21_',
    '8.2' : '8.2_',
    '4.26': '4.26_',
    '5.3' : 'stat5_3',
    '5.4' : 'stat5_4',
}

BASE_URL = 'https://app.myschool.sch.gr'
SSO_URL  = 'https://sso.sch.gr'


class MySchoolDownloader:
    """
    Selenium-based downloader για MySchool.

    Παράμετροι:
        username    : username MySchool (SSO)
        password    : κωδικός MySchool
        dest_dir    : φάκελος αποθήκευσης (δημιουργείται αυτόματα)
        callback    : callable(message) για progress reporting (default: print)
        reports     : list of report_ids να κατεβούν (default: όλα)
        browser     : 'chrome' (default) ή 'firefox'
    """

    def __init__(self, username, password, dest_dir,
                 callback=None, reports=None, browser='chrome', force=None):
        self.username  = username
        self.password  = password
        self.dest_dir  = dest_dir
        self.callback  = callback or print
        self.reports   = reports   # None = όλα
        self.browser   = (browser or 'chrome').lower().strip()  # 'chrome' ή 'firefox'
        self.force     = set(force) if force else set()  # rids για αναγκαστική λήψη

    def _log(self, msg):
        self.callback(msg)
        # Αποθήκευση log σε αρχείο — γράφει μέσα στον φάκελο λήψης (πάντα υπάρχει)
        from datetime import datetime as _dt
        line = f'[{_dt.now().strftime("%H:%M:%S")}] {msg}\n'
        # 1. Μέσα στον φάκελο της σημερινής λήψης
        try:
            if self.dest_dir and os.path.isdir(self.dest_dir):
                with open(os.path.join(self.dest_dir, 'run_log.txt'), 'a', encoding='utf-8') as f:
                    f.write(line)
        except Exception:
            pass
        # 2. Στον γονικό φάκελο downloads (fallback)
        try:
            parent = os.path.normpath(os.path.join(self.dest_dir, '..'))
            if os.path.isdir(parent):
                with open(os.path.join(parent, 'download_log.txt'), 'a', encoding='utf-8') as f:
                    f.write(line)
        except Exception:
            pass

    def _set_academic_year(self, driver, wait, school_year):
        """
        Αλλάζει το ενεργό σχολικό έτος στη σελίδα Default.aspx του MySchool.
        Χρησιμοποιεί το DevExpress combobox cmbActiveAcadYear.
        """
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException

        INPUT_ID = 'ctl00_ContentData_cmbActiveAcadYear_I'
        CTL_NAME = 'ctl00_ContentData_cmbActiveAcadYear'

        self._log(f'  Ορισμός σχολικού έτους: {school_year}')
        try:
            driver.get(BASE_URL + '/Default.aspx')
            time.sleep(2)

            inp = wait.until(EC.presence_of_element_located((By.ID, INPUT_ID)))
            current = inp.get_attribute('value') or ''
            if current.strip() == school_year.strip():
                self._log(f'  Σχολικό έτος ήδη σωστό: {current} — παράλειψη.')
                return

            driver.execute_script(
                "arguments[0].value = arguments[1];"
                "aspxETextChanged(arguments[2]);",
                inp, school_year, CTL_NAME
            )
            time.sleep(3)

            inp = driver.find_element(By.ID, INPUT_ID)
            new_val = inp.get_attribute('value') or ''
            if new_val.strip() == school_year.strip():
                self._log(f'  ✓ Σχολικό έτος ορίστηκε: {new_val}')
            else:
                self._log(f'  ⚠ Αναμενόμενο: {school_year} | Τρέχον: {new_val}')

        except TimeoutException:
            self._log('  ⚠ Combobox ακαδημαϊκού έτους δεν βρέθηκε.')
        except Exception as e:
            self._log(f'  ⚠ Αδυναμία αλλαγής έτους: {e}')

    def run(self):
        """
        Εκτελεί το κατέβασμα. Επιστρέφει dict {report_id: filepath_or_None}.
        """
        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.common.exceptions import TimeoutException, WebDriverException
            from selenium.webdriver.common.action_chains import ActionChains
        except ImportError:
            raise ImportError(
                'Η βιβλιοθήκη selenium δεν είναι εγκατεστημένη.\n'
                'Εκτέλεσε: pip install selenium'
            )

        # Δημιουργία φακέλου λήψης
        os.makedirs(self.dest_dir, exist_ok=True)
        self._log(f'Φάκελος λήψης: {self.dest_dir}')

        dest_resolved = str(Path(self.dest_dir).resolve())
        driver  = None
        results = {r[0]: None for r in REPORTS if self.reports is None or r[0] in self.reports}

        try:
            if self.browser == 'firefox':
                # ── Firefox ──────────────────────────────────────────
                self._log('Εκκίνηση Firefox...')
                from selenium.webdriver.firefox.service import Service as FirefoxService
                from selenium.webdriver.firefox.options import Options as FirefoxOptions

                ff_options = FirefoxOptions()
                ff_options.set_preference('browser.download.folderList', 2)
                ff_options.set_preference('browser.download.dir', dest_resolved)
                ff_options.set_preference('browser.download.useDownloadDir', True)
                ff_options.set_preference('browser.download.manager.showWhenStarting', False)
                ff_options.set_preference('browser.helperApps.neverAsk.saveToDisk',
                    'application/vnd.ms-excel,'
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,'
                    'text/csv,application/csv,application/octet-stream')
                ff_options.set_preference('pdfjs.disabled', True)

                # Προσπάθεια 1: Selenium built-in (4.6+) — χωρίς internet
                try:
                    driver = webdriver.Firefox(options=ff_options)
                    self._log('  GeckoDriver OK (Selenium built-in).')
                except Exception as _builtin_err:
                    self._log(f'  Selenium built-in απέτυχε: {_builtin_err}')

                    # Προσπάθεια 2: webdriver-manager
                    try:
                        from webdriver_manager.firefox import GeckoDriverManager
                        _gecko_path = GeckoDriverManager().install()
                        driver = webdriver.Firefox(
                            service=FirefoxService(_gecko_path), options=ff_options)
                        self._log('  GeckoDriver OK (webdriver-manager).')
                    except Exception as _ff_err:
                        raise RuntimeError(
                            f'Δεν ήταν δυνατή η εκκίνηση του Firefox.\n\n'
                            f'Βεβαιώσου ότι ο Firefox είναι εγκατεστημένος\n'
                            f'και ότι υπάρχει σύνδεση internet για αυτόματη λήψη GeckoDriver.'
                        ) from _ff_err

            else:
                # ── Chrome (default) ──────────────────────────────────
                self._log('Εκκίνηση Chrome...')
                from selenium.webdriver.chrome.service import Service as ChromeService

                options = webdriver.ChromeOptions()
                prefs = {
                    'download.default_directory'      : dest_resolved,
                    'download.prompt_for_download'     : False,
                    'download.directory_upgrade'       : True,
                    'safebrowsing.enabled'             : True,
                    'profile.default_content_setting_values.automatic_downloads': 1,
                    'credentials_enable_service'       : False,
                    'profile.password_manager_enabled' : False,
                }
                options.add_experimental_option('prefs', prefs)
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                options.add_argument('--window-size=1000,700')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')

                # Εύρεση τοπικού bundled driver (fallback)
                _base = os.path.dirname(os.path.abspath(__file__))
                _drivers_dir = os.path.normpath(
                    os.path.join(_base, '..', '..', 'drivers', 'chromedriver-win64'))
                _candidates = [
                    os.path.join(_drivers_dir, 'chromedriver.exe'),
                    os.path.normpath(os.path.join(_base, '..', '..', 'drivers', 'chromedriver.exe')),
                    os.path.normpath(os.path.join(_base, '..', 'drivers', 'chromedriver-win64', 'chromedriver.exe')),
                    os.path.normpath(os.path.join(_base, '..', 'drivers', 'chromedriver.exe')),
                ]
                _local = next((p for p in _candidates if os.path.isfile(p)), None)

                self._log('  Αυτόματη εύρεση/λήψη ChromeDriver...')

                # Προσπάθεια 1: Selenium built-in driver manager (Selenium 4.6+)
                # Δεν χρειάζεται εξωτερική βιβλιοθήκη — το Selenium κατεβάζει μόνο του
                try:
                    driver = webdriver.Chrome(options=options)
                    self._log('  ChromeDriver OK (Selenium built-in).')
                except Exception as _builtin_err:
                    self._log(f'  Selenium built-in απέτυχε: {_builtin_err}')

                    # Προσπάθεια 2: webdriver-manager
                    try:
                        from webdriver_manager.chrome import ChromeDriverManager
                        _wdm_path = ChromeDriverManager().install()
                        driver = webdriver.Chrome(service=ChromeService(_wdm_path), options=options)
                        self._log('  ChromeDriver OK (webdriver-manager).')
                    except Exception as _wdm_err:
                        self._log(f'  webdriver-manager απέτυχε: {_wdm_err}')

                        # Προσπάθεια 3: τοπικό bundled driver
                        if _local:
                            self._log('  Δοκιμάζω τοπικό driver...')
                            try:
                                driver = webdriver.Chrome(service=ChromeService(_local), options=options)
                                self._log('  Τοπικός driver OK.')
                            except Exception as _local_err:
                                raise RuntimeError(
                                    f'Δεν ήταν δυνατή η εκκίνηση του ChromeDriver.\n\n'
                                    f'Βεβαιώσου ότι ο Chrome είναι εγκατεστημένος και ενημερωμένος\n'
                                    f'και ότι υπάρχει σύνδεση internet για αυτόματη λήψη driver.'
                                ) from _local_err
                        else:
                            raise RuntimeError(
                                f'Δεν ήταν δυνατή η εκκίνηση του ChromeDriver.\n\n'
                                f'Βεβαιώσου ότι ο Chrome είναι εγκατεστημένος και ενημερωμένος\n'
                                f'και ότι υπάρχει σύνδεση internet για αυτόματη λήψη driver.'
                            ) from _wdm_err

            wait = WebDriverWait(driver, 20)

            # ── Login ────────────────────────────────────────────────
            self._log('Σύνδεση στο MySchool...')
            driver.get(BASE_URL)
            time.sleep(2)
            self._log(f'  URL μετά φόρτωση: {driver.current_url}')

            # SSO redirect
            if 'sso.sch.gr' in driver.current_url or 'login' in driver.current_url.lower():
                self._log('  Σελίδα SSO...')
                try:
                    user_field = wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[type="text"], input[name="username"], #username')
                    ))
                    user_field.clear()
                    user_field.send_keys(self.username)

                    pass_field = driver.find_element(
                        By.CSS_SELECTOR, 'input[type="password"], input[name="password"], #password'
                    )
                    pass_field.clear()
                    pass_field.send_keys(self.password)

                    # Submit
                    submit = driver.find_element(
                        By.CSS_SELECTOR, 'button[type="submit"], input[type="submit"]'
                    )
                    submit.click()
                    time.sleep(3)
                    self._log(f'  URL μετά login: {driver.current_url}')
                    self._log('  Login ολοκληρώθηκε.')
                except TimeoutException:
                    raise RuntimeError('Δεν βρέθηκε η φόρμα login του SSO.')
            else:
                self._log('  Ήδη συνδεδεμένος.')

            # Έλεγχος επιτυχίας login
            if 'sso.sch.gr' in driver.current_url and 'error' in driver.current_url.lower():
                raise RuntimeError('Λανθασμένα στοιχεία σύνδεσης MySchool.')

            # ── Κατέβασμα κάθε αρχείου ──────────────────────────────
            target_reports = [
                r for r in REPORTS
                if self.reports is None or r[0] in self.reports
            ]

            for report_entry in target_reports:
                rid, label, url_path, fname_base, wait_search, wait_dl = report_entry[:6]
                direct_export      = report_entry[6] if len(report_entry) > 6 else False
                custom_search      = report_entry[7] if len(report_entry) > 7 else None
                custom_export      = report_entry[8] if len(report_entry) > 8 else None
                pre_search_labels  = report_entry[9]  if len(report_entry) > 9  else None
                pre_search_js      = report_entry[10] if len(report_entry) > 10 else None
                report_year        = report_entry[11] if len(report_entry) > 11 else None

                self._log(f'[{rid}] {label}...')
                try:
                    # Αλλαγή σχολικού έτους αν απαιτείται για αυτό το report
                    if report_year:
                        self._set_academic_year(driver, wait, report_year)

                    # Έλεγχος αν το αρχείο υπάρχει ήδη στον φάκελο (same-day reuse)
                    existing = [
                        f for f in glob.glob(os.path.join(self.dest_dir, f'{fname_base}*'))
                        if not f.endswith('.tmp') and not f.endswith('.crdownload')
                    ]
                    if existing and rid not in self.force:
                        results[rid] = existing[0]
                        self._log(f'  [{rid}] Ήδη υπάρχει → {os.path.basename(existing[0])} — παράλειψη.')
                        continue
                    elif existing and rid in self.force:
                        for old in existing:
                            try: os.remove(old)
                            except Exception: pass
                        self._log(f'  [{rid}] Επαναλήψη λήψης (διαγράφηκε παλιό αρχείο).')

                    # Πλοήγηση απευθείας στο URL της αναφοράς
                    report_url = BASE_URL + url_path
                    self._log(f'  [{rid}] Φόρτωση: {report_url}')
                    driver.get(report_url)
                    time.sleep(3)
                    self._log(f'  [{rid}] URL: {driver.current_url}')

                    # Scroll στο κάτω μέρος για να φορτωθούν τυχόν lazy elements
                    driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
                    time.sleep(1)
                    driver.execute_script('window.scrollTo(0, 0);')
                    time.sleep(1)

                    # Debug: εκτύπωση όλων των κουμπιών/inputs/links στη σελίδα
                    elems = driver.find_elements(By.XPATH,
                        '//button | //input[@type="button"] | //input[@type="submit"] | //a')
                    self._log(f'  [{rid}] Στοιχεία στη σελίδα:')
                    for el in elems:
                        try:
                            tag  = el.tag_name
                            eid  = el.get_attribute('id') or ''
                            enam = el.get_attribute('name') or ''
                            eval_ = el.get_attribute('value') or ''
                            etxt = (el.text or '').strip()[:50]
                            ehref = el.get_attribute('href') or ''
                            if eid or eval_ or etxt or ehref:
                                self._log(f'    <{tag}> id={eid!r} name={enam!r} value={eval_!r} text={etxt!r} href={ehref[:60]!r}')
                        except Exception:
                            pass

                    # Κλείσιμο popup αν υπάρχει
                    try:
                        popup_btn = driver.find_element(By.CSS_SELECTOR, 'input[name*="popupControl"][name*="btnClose"]')
                        if popup_btn.is_displayed():
                            self._log(f'  [{rid}] Κλείσιμο popup...')
                            driver.execute_script('arguments[0].click();', popup_btn)
                            time.sleep(1)
                    except Exception:
                        pass

                    if not direct_export:
                        # ── Τσεκάρισμα checkboxes ομαδοποίησης (πριν αναζήτηση) ──────
                        if pre_search_labels:
                            self._log(f'  [{rid}] Τσεκάρισμα επιλογών ομαδοποίησης...')
                            for lbl_text in pre_search_labels:
                                try:
                                    # Βρες το <label> με το κείμενο
                                    lbl_el = WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located((By.XPATH,
                                            f'//label[normalize-space(text())="{lbl_text}"]'))
                                    )
                                    lbl_for = lbl_el.get_attribute('for')
                                    if lbl_for:
                                        # label έχει for= → click το input
                                        chk = driver.find_element(By.ID, lbl_for)
                                        driver.execute_script('arguments[0].click();', chk)
                                    else:
                                        # DevExpress/ASP.NET: ψάξε input ή dx-span στον ίδιο container
                                        parent_el = lbl_el.find_element(By.XPATH, '..')
                                        try:
                                            chk = parent_el.find_element(By.XPATH,
                                                './/input[@type="checkbox"] | '
                                                'preceding-sibling::*[contains(@class,"dxI") or '
                                                'contains(@class,"dx-checkbox")]')
                                            driver.execute_script('arguments[0].click();', chk)
                                        except Exception:
                                            # Fallback: click the label itself
                                            driver.execute_script('arguments[0].click();', lbl_el)
                                    time.sleep(0.4)
                                    self._log(f'  [{rid}]   ✓ "{lbl_text}"')
                                except Exception as _ck_err:
                                    self._log(f'  [{rid}]   ⚠ "{lbl_text}" δεν βρέθηκε: {_ck_err}')

                        # JS clicks (π.χ. DevExpress ListBox items — επιλογή ετών)
                        if pre_search_js:
                            self._log(f'  [{rid}] JS pre-search clicks...')
                            for element_id in pre_search_js:
                                try:
                                    el = WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located((By.ID, element_id))
                                    )
                                    driver.execute_script('arguments[0].click();', el)
                                    time.sleep(0.3)
                                    self._log(f'  [{rid}]   ✓ clicked #{element_id}')
                                except Exception as _js_err:
                                    self._log(f'  [{rid}]   ⚠ #{element_id}: {_js_err}')

                        # Κουμπί Αναζήτησης
                        search_clicked = False
                        if custom_search:
                            # Custom CSS selector (π.χ. topoth: 'a.hint_search')
                            try:
                                el = WebDriverWait(driver, wait_search).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, custom_search))
                                )
                                self._log(f'  [{rid}] Αναζήτηση (custom: {custom_search})...')
                                driver.execute_script('arguments[0].click();', el)
                                search_clicked = True
                            except TimeoutException:
                                self._log(f'  [{rid}] Custom κουμπί αναζήτησης δεν βρέθηκε.')
                        else:
                            # Standard ASP.NET submit button
                            try:
                                WebDriverWait(driver, wait_search).until(
                                    EC.presence_of_element_located((
                                        By.NAME, 'ctl00$cntStats$btnSubmit'
                                    ))
                                )
                                self._log(f'  [{rid}] Αναζήτηση...')
                                driver.execute_script(
                                    '$("input[name=\'ctl00$cntStats$btnSubmit\']").click();'
                                )
                                search_clicked = True
                            except TimeoutException:
                                self._log(f'  [{rid}] Κουμπί αναζήτησης δεν βρέθηκε — προσπάθεια εξαγωγής...')

                        # Περίμενε να φορτώσουν τα αποτελέσματα (πρώτη γραμμή grid)
                        try:
                            self._log(f'  [{rid}] Αναμονή αποτελεσμάτων...')
                            WebDriverWait(driver, wait_search).until(
                                EC.presence_of_element_located((
                                    By.XPATH,
                                    '//*[contains(@id,"DXDataRow0") or contains(@id,"gridResults_DXDataRow0")]'
                                ))
                            )
                            self._log(f'  [{rid}] Αποτελέσματα φορτώθηκαν.')
                        except TimeoutException:
                            self._log(f'  [{rid}] ΠΡΟΣΟΧΗ: Grid δεν φορτώθηκε εντός {wait_search}s.')
                    else:
                        self._log(f'  [{rid}] Άμεση εξαγωγή (χωρίς αναζήτηση).')
                        time.sleep(2)

                    # Κουμπί Εξαγωγής
                    # Πρώτα δοκιμάζουμε γνωστά ονόματα, μετά fallback σε οποιοδήποτε κουμπί εξαγωγής
                    try:
                        # Custom export selector (π.χ. topoth)
                        if custom_export:
                            try:
                                export_el = WebDriverWait(driver, wait_dl).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, custom_export))
                                )
                                self._log(f'  [{rid}] Εξαγωγή (custom: {custom_export})...')
                                before_files = set(os.listdir(self.dest_dir))
                                driver.execute_script('arguments[0].click();', export_el)
                                # Αναμονή αρχείου
                                WebDriverWait(driver, wait_dl).until(
                                    lambda d: bool(
                                        set(os.listdir(self.dest_dir)) - before_files - {
                                            f for f in set(os.listdir(self.dest_dir)) - before_files
                                            if f.endswith('.crdownload') or f.endswith('.tmp')
                                        }
                                    )
                                )
                                time.sleep(2)
                                new_files = {f for f in set(os.listdir(self.dest_dir)) - before_files
                                             if not f.endswith('.crdownload') and not f.endswith('.tmp')}
                                if new_files:
                                    raw = os.path.join(self.dest_dir, sorted(new_files)[-1])
                                    if raw.endswith('.zip'):
                                        raw = self._extract_zip(raw)
                                    ext   = os.path.splitext(raw)[1]
                                    final = os.path.join(self.dest_dir, fname_base + ext)
                                    if raw != final:
                                        shutil.move(raw, final)
                                    results[rid] = final
                                    self._log(f'  [{rid}] OK → {os.path.basename(final)}')
                                else:
                                    self._log(f'  [{rid}] ΠΡΟΣΟΧΗ: Δεν βρέθηκε νέο αρχείο.')
                                continue  # επόμενο report
                            except TimeoutException:
                                self._log(f'  [{rid}] Custom export δεν βρέθηκε — fallback...')

                        KNOWN_EXPORT_NAMES = [
                            'ctl00$cntStats$btnCSVExport',
                            'ctl00$cntStats$btnToExcel',
                            'ctl00$cntContent$btnExport',
                            'ctl00$cntContent$btnCSVExport',
                            'ctl00$cntContent$btnToExcel',
                            'ctl00$cntContent$btnDownload',
                        ]
                        EXPORT_KEYWORDS = ('export', 'excel', 'csv', 'download',
                                           'εξαγωγη', 'εξαγωγή', 'ληψη', 'λήψη')

                        def _find_export_button(d):
                            # 1. Δοκίμασε γνωστά ονόματα (input/button)
                            for name in KNOWN_EXPORT_NAMES:
                                if d.find_elements(By.NAME, name):
                                    return ('name', name)
                            # 2. Fallback: οποιοδήποτε input/button με export/excel/csv στο name ή value
                            for el in d.find_elements(By.XPATH,
                                    '//input[@type="submit" or @type="button"] | //button'):
                                try:
                                    el_name  = (el.get_attribute('name')  or '').lower()
                                    el_value = (el.get_attribute('value') or '').lower()
                                    el_text  = (el.text or '').lower()
                                    combined = el_name + ' ' + el_value + ' ' + el_text
                                    if any(kw in combined for kw in EXPORT_KEYWORDS):
                                        return ('element', el)
                                except Exception:
                                    pass
                            # 3. ASP.NET LinkButton: <a> tag με export-related text ή id
                            #    Π.χ. <a href="javascript:__doPostBack(...)">Εξαγωγή</a>
                            NAV_SKIP = ('logout', 'signout', 'default.aspx', 'login',
                                        'home', 'javascript:void', '#')
                            for el in d.find_elements(By.XPATH, '//a'):
                                try:
                                    el_text = (el.text or '').strip().lower()
                                    el_id   = (el.get_attribute('id')   or '').lower()
                                    el_href = (el.get_attribute('href') or '').lower()
                                    # Αγνόησε nav links
                                    if any(skip in el_href for skip in NAV_SKIP):
                                        continue
                                    # Έλεγξε text ή id για export-related λέξεις
                                    combined = el_text + ' ' + el_id
                                    if any(kw in combined for kw in EXPORT_KEYWORDS):
                                        return ('element', el)
                                    # Επίσης έλεγξε αν το __doPostBack ID περιέχει export λέξεις
                                    if '__dopostback' in el_href and any(kw in el_href for kw in EXPORT_KEYWORDS):
                                        return ('element', el)
                                except Exception:
                                    pass
                            return None

                        found = WebDriverWait(driver, wait_dl).until(_find_export_button)

                        if found:
                            kind, target = found
                            if kind == 'name':
                                export_name = target
                                self._log(f'  [{rid}] Εξαγωγή ({export_name})...')
                                before_files = set(os.listdir(self.dest_dir))
                                driver.execute_script(
                                    f'$("input[name=\'{export_name}\']").click();'
                                )
                            else:
                                # Fallback element — click directly (input, button ή <a> LinkButton)
                                el_tag   = target.tag_name
                                el_name  = target.get_attribute('name')  or ''
                                el_id    = target.get_attribute('id')    or ''
                                el_value = target.get_attribute('value') or ''
                                el_text  = (target.text or '').strip()
                                self._log(f'  [{rid}] Εξαγωγή fallback <{el_tag}>'
                                          f' id={el_id!r} name={el_name!r}'
                                          f' value={el_value!r} text={el_text!r}...')
                                before_files = set(os.listdir(self.dest_dir))
                                driver.execute_script('arguments[0].click();', target)
                        else:
                            self._log(f'  [{rid}] Κανένα κουμπί εξαγωγής βρέθηκε — παράλειψη.')
                            results[rid] = None
                            continue

                        # Περίμενε να εμφανιστεί νέο αρχείο (όχι .crdownload/.tmp)
                        self._log(f'  [{rid}] Αναμονή αρχείου...')
                        WebDriverWait(driver, wait_dl).until(
                            lambda d: bool(
                                set(os.listdir(self.dest_dir)) - before_files - {
                                    f for f in set(os.listdir(self.dest_dir)) - before_files
                                    if f.endswith('.crdownload') or f.endswith('.tmp')
                                }
                            )
                        )
                        time.sleep(2)  # μικρή αναμονή για να ολοκληρωθεί η εγγραφή

                        # Βρες το νέο αρχείο
                        new_files = set(os.listdir(self.dest_dir)) - before_files
                        new_files = {f for f in new_files
                                     if not f.endswith('.crdownload')
                                     and not f.endswith('.tmp')}

                        if new_files:
                            raw = os.path.join(self.dest_dir, sorted(new_files)[-1])

                            # Αποσυμπίεση αν ZIP
                            if raw.endswith('.zip'):
                                raw = self._extract_zip(raw)

                            # Μετονομασία με ξεκάθαρο όνομα
                            ext = os.path.splitext(raw)[1]
                            final = os.path.join(self.dest_dir, fname_base + ext)
                            if raw != final:
                                shutil.move(raw, final)
                            results[rid] = final
                            self._log(f'  [{rid}] OK → {os.path.basename(final)}')
                        else:
                            self._log(f'  [{rid}] ΠΡΟΣΟΧΗ: Δεν βρέθηκε νέο αρχείο.')

                    except TimeoutException:
                        self._log(f'  [{rid}] Κουμπί εξαγωγής ή αρχείο δεν βρέθηκε.')

                except Exception as e:
                    self._log(f'  [{rid}] ΣΦΑΛΜΑ: {e}')
                    results[rid] = None

            # Σύνοψη
            ok   = sum(1 for v in results.values() if v)
            fail = len(results) - ok
            self._log(f'\nΚατεβηκαν: {ok}/{len(results)} αρχεία' +
                      (f' | Αποτυχίες: {fail}' if fail else ''))

        except Exception as e:
            self._log(f'ΚΡΙΤΙΚΟ ΣΦΑΛΜΑ: {e}')
            raise
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

        return results

    def _extract_zip(self, zip_path):
        """Αποσυμπιέζει ZIP και επιστρέφει το εσωτερικό αρχείο."""
        extract_dir = os.path.dirname(zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zf:
            names = zf.namelist()
            zf.extractall(extract_dir)
        os.remove(zip_path)
        if names:
            return os.path.join(extract_dir, names[0])
        return zip_path


# ── Βοηθητικές για διαχείριση φακέλου downloads ─────────────────────────────

def get_downloads_dir(base_dir):
    """
    Επιστρέφει τον φάκελο downloads της σημερινής ημέρας (YYYYMMDD).
    Αν υπάρχει ήδη από προηγούμενη εκτέλεση σήμερα, τον επαναχρησιμοποιεί
    ώστε να μην ξανακατεβούν αρχεία που υπάρχουν ήδη.
    """
    today = datetime.now().strftime('%Y%m%d')
    return os.path.join(base_dir, 'downloads', today)


def cleanup_old_downloads(base_dir, keep=3):
    """
    Κρατά μόνο τους τελευταίους `keep` φακέλους downloads (default: 3).
    Διαγράφει τους παλαιότερους.
    """
    dl_base = os.path.join(base_dir, 'downloads')
    if not os.path.exists(dl_base):
        return
    folders = sorted([
        os.path.join(dl_base, d)
        for d in os.listdir(dl_base)
        if os.path.isdir(os.path.join(dl_base, d))
    ], reverse=True)

    for old_folder in folders[keep:]:
        try:
            shutil.rmtree(old_folder)
        except Exception:
            pass



def find_latest_downloads(base_dir):
    """
    Ψάχνει τον φάκελο downloads της σημερινής μέρας (YYYYMMDD) και επιστρέφει
    dict {report_id: filepath} για τα αρχεία που υπάρχουν.
    """
    today = datetime.now().strftime('%Y%m%d')
    today_dir = os.path.join(base_dir, 'downloads', today)
    if not os.path.isdir(today_dir):
        return {}

    result = {}
    for rid, prefix in FILE_PREFIX_MAP.items():
        matches = glob.glob(os.path.join(today_dir, f'{prefix}*'))
        if matches:
            result[rid] = max(matches, key=os.path.getmtime)

    return result


def downloads_info(base_dir):
    """
    Επιστρέφει πληροφορίες για τα διαθέσιμα downloads:
    (timestamp_str, dict {report_id: filename}, age_minutes)
    ή None αν δεν υπάρχουν.
    """
    dl_base = os.path.join(base_dir, 'downloads')
    if not os.path.exists(dl_base):
        return None

    folders = sorted([
        d for d in os.listdir(dl_base)
        if os.path.isdir(os.path.join(dl_base, d))
    ], reverse=True)

    if not folders:
        return None

    latest_name = folders[0]
    latest_path = os.path.join(dl_base, latest_name)

    # Ηλικία σε λεπτά (υπολογισμός με βάση την ημερομηνία YYYYMMDD)
    try:
        ts      = datetime.strptime(latest_name, '%Y%m%d')
        age_min = int((datetime.now() - ts).total_seconds() / 60)
        ts_str  = ts.strftime('%d/%m/%Y')
    except ValueError:
        # Fallback για παλιές ονομασίες YYYYMMDD_HHMMSS
        try:
            ts      = datetime.strptime(latest_name, '%Y%m%d_%H%M%S')
            age_min = int((datetime.now() - ts).total_seconds() / 60)
            ts_str  = ts.strftime('%d/%m/%Y %H:%M')
        except ValueError:
            age_min = 0
            ts_str  = latest_name

    # Αρχεία που βρέθηκαν
    found = {}
    for rid, prefix in FILE_PREFIX_MAP.items():
        matches = glob.glob(os.path.join(latest_path, f'{prefix}*'))
        if matches:
            found[rid] = os.path.basename(matches[0])

    return ts_str, found, age_min
