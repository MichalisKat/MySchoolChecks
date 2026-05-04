"""
smeae/downloader.py
═══════════════════
Κατέβασμα στατιστικών ΣΜΕΑΕ/ΕΕΑ από το MySchool.
10 αρχεία .xls — 9 ΣΜΕΑΕ + Βασικά Στοιχεία Σχολικών Μονάδων.

Χρησιμοποιεί Chrome με auto-download (χωρίς PyAutoGUI/watchdog),
ίδιο pattern με core/downloader.py.
"""
import os, time, glob, shutil
from pathlib import Path

BASE_URL = 'https://app.myschool.sch.gr'

# (num, label, url_path, fname_base, wait_submit_sec, wait_dl_sec)
SMEAE_REPORTS = [
    (1, 'Συγκεντρωτικά ΕΕΑ',
     '/Statistics/Management.stat.StudentEEA.aspx?parentId=11',
     '1. Συγκεντρωτικά Στοιχεία μαθητών-μαθητριών με ΕΕΑ',
     25, 40),
    (2, 'ΕΕΑ — Εκπαιδευτικός Τάξης',
     '/Statistics/Management.stat.StudentEEASupport.aspx?parentId=11',
     '2. Συγκεντρωτικά Στοιχεία μαθητών-μαθητριών με ΕΕΑ που υποστηρίζονται από τον-την εκπαιδευτικό της τάξης',
     25, 50),
    (3, 'Τμήματα Ένταξης — κοινό/εξειδικευμένο',
     '/Statistics/Management.stat.studentInAccessionGroupWithSpecProgram.aspx?parentId=11',
     '3. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών σε Τμήματα Ένταξης με κοινό και εξειδικευμένο πρόγραμμα',
     25, 40),
    (4, 'Τμήματα Ένταξης — διευρυμένο ωράριο',
     '/Statistics/Management.stat.studentInAccessionGroupWithExtendedProgram.aspx?parentId=11',
     '4. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών σε Τμήματα Ένταξης διευρυμένου ωραρίου',
     25, 40),
    (5, 'Παράλληλη Στήριξη',
     '/Statistics/Management.stat.studentInParallelSupport.aspx?parentId=11',
     '5. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών σε Παράλληλη Στήριξη',
     25, 40),
    (6, 'Ειδικό Βοηθητικό Προσωπικό',
     '/Statistics/Management.stat.studentSupportedBySpecial.aspx?parentId=11',
     '6. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών με Ειδικό Βοηθητικό Προσωπικό',
     25, 40),
    (7, 'Σχολικός Νοσηλευτής',
     '/Statistics/Management.stat.studentSupportedBySchoolNurse.aspx?parentId=11',
     '7. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών με Σχολικό Νοσηλευτή',
     25, 40),
    (8, 'Ειδικός Βοηθός (οικογένεια)',
     '/Statistics/Management.stat.studentSupportedBySpecialAssistant.aspx?parentId=11',
     '8. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών με ειδικό βοηθό (που διαθέτει η οικογένεια)',
     25, 40),
    (9, "Κατ' οίκον Διδασκαλία",
     '/Statistics/Management.stat.studentSupportedAtHome.aspx?parentId=11',
     "9. Συγκεντρωτικά Στοιχεία Ειδικών Εκπαιδευτικών Αναγκών Μαθητών που υποστηρίζονται Κατ' οίκον",
     25, 40),
    (10, 'Βασικά Στοιχεία Σχολικών Μονάδων',
     '/Statistics/Management.stat.infoUnits.aspx?parentId=3',
     '10. Βασικά Στοιχεία Σχολικών Μονάδων',
     25, 40),
]


class SmeaeDownloader:
    """
    Κατεβάζει τα 9 στατιστικά ΣΜΕΑΕ από MySchool.
    Χρησιμοποιεί Chrome με auto-download (χωρίς PyAutoGUI),
    ίδιο pattern με MySchoolDownloader.
    """

    def __init__(self, username, password, dest_dir, callback=None):
        self.username = username
        self.password = password
        self.dest_dir = str(Path(dest_dir).resolve())
        self.callback = callback or print

    def _log(self, msg):
        self.callback(msg)

    def run(self):
        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.common.exceptions import TimeoutException
            from selenium.webdriver.chrome.service import Service as ChromeService
        except ImportError:
            raise ImportError('Η βιβλιοθήκη selenium δεν είναι εγκατεστημένη.')

        os.makedirs(self.dest_dir, exist_ok=True)
        self._log(f'Φάκελος λήψης ΣΜΕΑΕ: {self.dest_dir}')

        options = webdriver.ChromeOptions()
        prefs = {
            'download.default_directory': self.dest_dir,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True,
            'profile.default_content_setting_values.automatic_downloads': 1,
            'credentials_enable_service': False,
            'profile.password_manager_enabled': False,
        }
        options.add_experimental_option('prefs', prefs)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument('--window-size=1000,700')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        driver  = None
        results = {}

        try:
            self._log('Εκκίνηση Chrome...')
            _base = os.path.dirname(os.path.abspath(__file__))
            _candidates = [
                os.path.normpath(os.path.join(_base, '..', '..', 'drivers', 'chromedriver-win64', 'chromedriver.exe')),
                os.path.normpath(os.path.join(_base, '..', '..', 'drivers', 'chromedriver.exe')),
                os.path.normpath(os.path.join(_base, '..', 'drivers', 'chromedriver-win64', 'chromedriver.exe')),
                os.path.normpath(os.path.join(_base, '..', 'drivers', 'chromedriver.exe')),
            ]
            _local = next((p for p in _candidates if os.path.isfile(p)), None)

            if _local:
                self._log(f'  Χρήση τοπικού driver: {_local}')
                driver = webdriver.Chrome(service=ChromeService(_local), options=options)
            else:
                # Fallback: Selenium Manager (απαιτεί internet) — ίδιο με core/downloader.py
                self._log('  Αυτόματη αναζήτηση driver (απαιτεί internet)...')
                try:
                    driver = webdriver.Chrome(options=options)
                except Exception as e:
                    chrome_ver = ''
                    try:
                        import winreg
                        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r'Software\Google\Chrome\BLBeacon')
                        chrome_ver = winreg.QueryValueEx(key, 'version')[0]
                        winreg.CloseKey(key)
                    except Exception:
                        pass
                    ver_msg = f' (Chrome {chrome_ver})' if chrome_ver else ''
                    raise RuntimeError(
                        f'Δεν βρέθηκε Chrome WebDriver{ver_msg}.\n\n'
                        f'Κατέβασε το chromedriver από:\n'
                        f'https://storage.googleapis.com/chrome-for-testing-public/'
                        f'{chrome_ver or "VERSION"}/win64/chromedriver-win64.zip\n\n'
                        f'και τοποθέτησέ το στον φάκελο:\n'
                        f'{os.path.normpath(os.path.join(_base, "..", "drivers", "chromedriver-win64"))}'
                    ) from e

            wait = WebDriverWait(driver, 20)

            # ── Login ────────────────────────────────────────────────────────
            self._log('Σύνδεση στο MySchool...')
            driver.get(BASE_URL)
            time.sleep(2)

            if 'sso.sch.gr' in driver.current_url or 'login' in driver.current_url.lower():
                try:
                    user_field = wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR,
                         'input[type="text"], input[name="username"], #username')
                    ))
                    user_field.clear()
                    user_field.send_keys(self.username)
                    pass_field = driver.find_element(
                        By.CSS_SELECTOR,
                        'input[type="password"], input[name="password"], #password'
                    )
                    pass_field.clear()
                    pass_field.send_keys(self.password)
                    submit = driver.find_element(
                        By.CSS_SELECTOR, 'button[type="submit"], input[type="submit"]'
                    )
                    submit.click()
                    time.sleep(3)
                    self._log('  Login ολοκληρώθηκε.')
                except TimeoutException:
                    raise RuntimeError('Δεν βρέθηκε η φόρμα login του SSO.')

            if 'sso.sch.gr' in driver.current_url and 'error' in driver.current_url.lower():
                raise RuntimeError('Λανθασμένα στοιχεία σύνδεσης MySchool.')

            # ── Λήψη κάθε αρχείου ────────────────────────────────────────────
            for (num, label, url_path, fname_base, wait_submit, wait_dl) in SMEAE_REPORTS:
                self._log(f'[{num}/10] {label}...')
                try:
                    # Παράλειψη αν υπάρχει ήδη
                    existing = [
                        f for f in glob.glob(os.path.join(self.dest_dir, f'{num}.*'))
                        if not f.endswith(('.tmp', '.crdownload'))
                    ]
                    if existing:
                        results[num] = existing[0]
                        self._log(f'  Υπάρχει ήδη — παράλειψη.')
                        continue

                    driver.get(BASE_URL + url_path)
                    time.sleep(2)

                    # Κουμπί Υποβολής
                    try:
                        WebDriverWait(driver, wait_submit).until(
                            EC.presence_of_element_located(
                                (By.ID, 'ctl00_cntStats_btnSubmit_CD'))
                        )
                        driver.find_element(By.ID, 'ctl00_cntStats_btnSubmit_CD').click()
                        self._log('  Αναζήτηση...')
                    except TimeoutException:
                        self._log('  Κουμπί υποβολής δεν βρέθηκε — δοκιμή εξαγωγής...')

                    # Αναμονή πλέγματος αποτελεσμάτων
                    try:
                        WebDriverWait(driver, wait_submit).until(
                            EC.presence_of_element_located(
                                (By.ID, 'ctl00_cntStats_gridResults_DXDataRow0'))
                        )
                    except TimeoutException:
                        self._log('  Αποτελέσματα δεν φόρτωσαν εγκαίρως — δοκιμή εξαγωγής...')

                    # Κουμπί Εξαγωγής σε Excel
                    before = set(os.listdir(self.dest_dir))
                    try:
                        export_btn = WebDriverWait(driver, wait_dl).until(
                            EC.element_to_be_clickable(
                                (By.ID, 'ctl00_cntStats_btnToExcel_CD'))
                        )
                        export_btn.click()
                        self._log('  Εξαγωγή...')
                    except TimeoutException:
                        self._log('  Κουμπί εξαγωγής δεν βρέθηκε.')
                        results[num] = None
                        continue

                    # Αναμονή νέου αρχείου
                    try:
                        WebDriverWait(driver, wait_dl).until(
                            lambda d: bool(
                                set(os.listdir(self.dest_dir)) - before - {
                                    f for f in set(os.listdir(self.dest_dir)) - before
                                    if f.endswith(('.crdownload', '.tmp'))
                                }
                            )
                        )
                    except TimeoutException:
                        self._log('  Timeout — αρχείο δεν εμφανίστηκε.')
                        results[num] = None
                        continue

                    time.sleep(1)
                    new_files = {
                        f for f in set(os.listdir(self.dest_dir)) - before
                        if not f.endswith(('.crdownload', '.tmp'))
                    }

                    if new_files:
                        raw = os.path.join(self.dest_dir, sorted(new_files)[-1])
                        ext = os.path.splitext(raw)[1] or '.xls'
                        final = os.path.join(self.dest_dir, fname_base + ext)
                        if os.path.abspath(raw) != os.path.abspath(final):
                            shutil.move(raw, final)
                        results[num] = final
                        self._log(f'  OK → {os.path.basename(final)}')
                    else:
                        self._log('  Δεν βρέθηκε νέο αρχείο.')
                        results[num] = None

                except Exception as e:
                    self._log(f'  ΣΦΑΛΜΑ: {e}')
                    results[num] = None

            ok = sum(1 for v in results.values() if v)
            self._log(f'\nΚατεβήκαν: {ok}/10 αρχεία')

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
