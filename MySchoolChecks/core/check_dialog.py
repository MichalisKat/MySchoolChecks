"""
core/check_dialog.py
═════════════════════
Γενικό (generic) 3-tab Toplevel dialog για τους «απλούς» ελέγχους της
κεντρικής λίστας (main.py — LauncherApp): Λήψη / Εκτέλεση / Αποστολή.

Ίδιο μοτίβο με smeae/dialog.py::SmeaeDialog, αλλά παραμετρικό ως προς το
check_module — γράφεται μία φορά και χρησιμοποιείται από όλους τους
ελέγχους (adies, adies_aneu, analipsi, apontes_xwris_adeia,
arnhtika_ypoloipa, dioikitiko_ergo, forma_82, orario_diafora, ypoloipa).

  Tab 1 «⬇ Λήψη»      : κατεβάζει ΜΟΝΟ τα REQUIRED_REPORTS του ελέγχου
                        (μέσω core.downloader.report_ids_from_required +
                        MySchoolDownloader(reports=...)).
  Tab 2 «▶ Εκτέλεση»  : ask_inputs() → process() → αποθήκευση Excel, χωρίς
                        popup — η σύνοψη γράφεται στο text pane του tab
                        (core.framework.execute_check).
  Tab 3 «✉ Αποστολή»  : επιλογή Δοκιμαστική/Κανονική αποστολή, ενεργή μόνο
                        αφού έχει τρέξει επιτυχώς η Εκτέλεση
                        (core.framework.send_from_exec_result).

Για ελέγχους με CUSTOM_RUN = True (σήμερα: dioikitiko_ergo, ypoloipa) η
Εκτέλεση καλεί απευθείας mod.run(config) όπως έκανε πάντα το main.py — αυτά
τα modules διαχειρίζονται ήδη μόνα τους ολόκληρη τη ροή αποστολής
(περιλαμβανομένου ενός popup επιλογής Test/Κανονική μέσα στο run()), οπότε
το tab «Αποστολή» παραμένει κλειδωμένο με επεξηγηματικό μήνυμα.
"""
import os, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox


# ── Πρότυπα email (local_settings.json — data/local_settings.json) ─────────
# Ίδιο path/format με ό,τι ήδη διαβάζει core/framework.py (execute_check) όταν
# ψάχνει custom template — βλ. εκεί για το 'email_templates' key. Το tab
# «✉ Αποστολή» (CheckRunDialog._open_email_editor) διαβάζει/γράφει εδώ.
def _local_settings_path():
    import sys as _sys
    if getattr(_sys, 'frozen', False):
        _exe_dir = os.path.dirname(_sys.executable)
        if 'program files' in _exe_dir.lower():
            base = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'MySchoolChecks')
        else:
            base = _exe_dir
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, 'data', 'local_settings.json')


def _load_local_settings():
    import json
    path = _local_settings_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_local_settings(data):
    import json
    path = _local_settings_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class CheckRunDialog(tk.Toplevel):

    def __init__(self, parent, config_ref, docs_base, C, check_module):
        super().__init__(parent)
        self._cfg   = config_ref
        self._base  = docs_base          # π.χ. Documents\MySchoolChecks
        self._C     = C
        self._mod   = check_module
        self._exec_result   = None
        self._is_custom_run = getattr(check_module, 'CUSTOM_RUN', False)

        title = getattr(check_module, 'CHECK_TITLE', '?')
        self.title(f'Έλεγχος — {title}')
        self.configure(bg=C['bg'])
        self.resizable(False, False)
        self.grab_set()
        self.transient(parent)

        app_ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'app.ico')
        if os.path.exists(app_ico):
            try:
                self.iconbitmap(app_ico)
            except Exception:
                pass

        self._build()
        self.update_idletasks()
        self.geometry('640x620')
        pw = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        ph = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f'+{pw}+{ph}')

    # ── Widget helpers ──────────────────────────────────────────────────────

    def _log_append(self, widget, msg):
        def _do():
            try:
                widget.configure(state='normal')
                widget.insert(tk.END, msg + '\n')
                widget.see(tk.END)
                widget.configure(state='disabled')
            except Exception:
                pass
        self.after(0, _do)

    def _make_log(self, parent, height=10):
        log = scrolledtext.ScrolledText(
            parent, height=height, font=('Consolas', 8),
            relief='solid', bd=1, state='disabled',
            bg='#F5F5F5', wrap=tk.WORD)
        log.pack(fill='both', expand=True, pady=(6, 0))
        return log

    def _run_btn(self, parent, text, cmd):
        C = self._C
        btn = tk.Button(
            parent, text=text,
            bg=C['btn_bg'], fg=C['btn_fg'],
            font=('Arial', 10, 'bold'),
            relief='flat', padx=16, pady=6,
            cursor='hand2', command=cmd)
        btn.pack(side='right', pady=(8, 0))
        return btn

    def _section_lbl(self, parent, text):
        C = self._C
        tk.Label(parent, text=text, bg=C['bg'], fg=C['hdr_bg'],
                 font=('Arial', 9, 'bold')).pack(anchor='w', pady=(0, 3))

    def _redirect_print(self, on_log):
        """Επιστρέφει (redirected_print, restore) — redirect του builtins.print
        στο log widget του τρέχοντος tab, όσο διαρκεί μια εργασία."""
        import builtins
        orig = builtins.print

        def _redirected(*args, **kwargs):
            try:
                text = ' '.join(str(a) for a in args)
            except Exception:
                text = ''
            if text.strip():
                on_log(text)

        def _install():
            builtins.print = _redirected

        def _restore():
            builtins.print = orig

        return _install, _restore

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        C = self._C
        title = getattr(self._mod, 'CHECK_TITLE', '?')

        hdr = tk.Frame(self, bg=C['hdr_bg'], pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text=f'📝  {title}',
                 bg=C['hdr_bg'], fg=C['hdr_fg'],
                 font=('Arial', 11, 'bold'), wraplength=600,
                 justify='center').pack()

        style = ttk.Style()
        style.configure('TNotebook', background=C['bg'])
        style.configure('TNotebook.Tab', background=C['bg2'], foreground=C['desc'],
                        font=('Arial', 9, 'bold'), padding=(10, 5))
        style.map('TNotebook.Tab',
                  background=[('selected', C['hdr_sub']), ('active', C['sel_bg'])],
                  foreground=[('selected', C['hdr_bg']),  ('active', C['hdr_bg'])])

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=12, pady=10)

        t1 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
        t2 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
        nb.add(t1, text='  ⬇ Λήψη  ')
        nb.add(t2, text='  ▶ Εκτέλεση  ')
        self._build_download(t1)
        self._build_execute(t2)

        # Έλεγχοι με CUSTOM_SPLIT_TAB (π.χ. tmimata_genikis) εκθέτουν ένα
        # επιπλέον tab «✂ Διαχωρισμός» — ΠΡΙΝ την «Αποστολή» (χωρίζει το
        # συγκεντρωτικό αρχείο αποτελεσμάτων σε ένα Excel ανά σχολείο, ίδια
        # λογική με το tab Διαχωρισμός του ελέγχου Ε.Ε.Α. — smeae/dialog.py).
        # Δεν εξαρτάται από το αν έχει τρέξει η Εκτέλεση μέσα σε αυτή τη
        # σύνοδο — ανιχνεύει το πιο πρόσφατο αρχείο από τον δίσκο.
        custom_split = getattr(self._mod, 'CUSTOM_SPLIT_TAB', None)
        self._has_generic_split = False
        if callable(custom_split):
            t3 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
            nb.add(t3, text='  ✂ Διαχωρισμός  ')
            try:
                custom_split(t3, self._cfg)
            except Exception as e:
                import traceback
                traceback.print_exc()
                tk.Label(t3, text=f'✗ Σφάλμα κατά τη δημιουργία του tab: {e}',
                         bg=C['bg'], fg='#B00020', font=('Arial', 9),
                         wraplength=560, justify='left').pack(anchor='w', pady=8)
        elif self._generic_split_applicable():
            # Ίδια λογική για ΟΛΟΥΣ τους «απλούς» ελέγχους (χωρίς CUSTOM_RUN,
            # χωρίς CUSTOM_SPLIT_TAB): tab «✂ Διαχωρισμός» πριν την
            # «Αποστολή» — χωρίζει το συνολικό αρχείο σε ένα Excel ανά
            # σχολείο (φάκελος «split»), ώστε να φαίνεται τι θα σταλεί σε
            # κάθε σχολείο πριν πατηθεί «Αποστολή» (βλ.
            # core/framework.py::split_exec_result). Προς το παρόν δεν
            # εξαιρεί κανένα σχολείο — μόνο σπάει το αρχείο.
            t3 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
            nb.add(t3, text='  ✂ Διαχωρισμός  ')
            self._build_generic_split(t3)
            self._has_generic_split = True

        # NO_SEND_TAB: ελέγχοι αμιγώς εσωτερικοί (π.χ. «Διαφορές AK-AL») δεν
        # παίρνουν καθόλου tab «✉ Αποστολή» — δεν αφορούν αποστολή σε
        # σχολεία/τρίτους, μόνο εσωτερική προβολή αποτελεσμάτων.
        self._has_send_tab = not getattr(self._mod, 'NO_SEND_TAB', False)
        if self._has_send_tab:
            t_send = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
            nb.add(t_send, text='  ✉ Αποστολή  ')
            self._build_send(t_send)

            # Η Αποστολή είναι κλειδωμένη μέχρι να τρέξει επιτυχώς η Εκτέλεση
            self._lock_send_tab(True)

        foot = tk.Frame(self, bg=C['bg2'], pady=10)
        foot.pack(fill='x')
        tk.Button(foot, text='Κλείσιμο',
                  bg=C['bg2'], fg=C['desc'],
                  font=('Arial', 10), relief='flat', padx=12, pady=5,
                  cursor='hand2', command=self.destroy).pack(side='right', padx=16)

    # ── Tab 1: Λήψη ──────────────────────────────────────────────────────────

    def _build_download(self, body):
        C = self._C

        # Έλεγχοι με CUSTOM_DOWNLOAD (π.χ. tmimata_genikis — χρειάζεται
        # override σχολικού έτους) χρησιμοποιούν τη δική τους function αντί
        # για τον γενικό μηχανισμό λήψης παρακάτω. Βλ. _build_custom_download.
        custom_dl = getattr(self._mod, 'CUSTOM_DOWNLOAD', None)
        if callable(custom_dl):
            self._build_custom_download(body, custom_dl)
            return

        from core.downloader import report_ids_from_required, FILE_PREFIX_MAP
        import glob as _glob

        required = getattr(self._mod, 'REQUIRED_REPORTS', [])
        rids = report_ids_from_required(required)
        self._dl_rids = rids

        self._section_lbl(body, 'Απαιτούμενα στατιστικά:')
        txt = ('\n'.join(f'  •  {r}' for r in required) if required
               else '  (δεν έχουν οριστεί σε αυτόν τον έλεγχο)')
        tk.Label(body, text=txt, bg=C['bg'], fg=C['desc'],
                 font=('Arial', 9), justify='left', anchor='w',
                 wraplength=560).pack(fill='x', pady=(0, 8))

        # Έλεγχος αν υπάρχουν ήδη σήμερα
        today_dir = os.path.join(self._base, 'downloads', datetime.now().strftime('%Y%m%d'))
        have = set()
        if rids and os.path.isdir(today_dir):
            for rid in rids:
                prefix  = FILE_PREFIX_MAP.get(rid, rid)
                matches = [f for f in _glob.glob(os.path.join(today_dir, f'{prefix}*'))
                           if not f.endswith(('.tmp', '.crdownload'))]
                if matches:
                    have.add(rid)

        if rids:
            complete = (have == set(rids))
            status_txt = (f'✓ Σήμερα υπάρχουν ήδη {len(have)}/{len(rids)} αρχεία.' if have
                          else 'Δεν έχουν κατέβει ακόμα σήμερα.')
            fg = C['status_ok'] if complete else C['desc']
            tk.Label(body, text=status_txt, bg=C['bg'], fg=fg,
                     font=('Arial', 8, 'italic')).pack(anchor='w', pady=(0, 6))

        self._dl_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        n = len(rids)
        label = f'⬇  Λήψη {n} Αρχεί{"ου" if n == 1 else "ων"}' if n else '⬇  Λήψη'
        self._dl_btn = self._run_btn(br, label, self._start_download)
        if not rids:
            self._dl_btn.configure(state='disabled')

    def _start_download(self):
        C   = self._C
        usr = getattr(self._cfg, 'MYSCHOOL_USER', '').strip()
        pw  = getattr(self._cfg, 'MYSCHOOL_PASS', '').strip()
        if not usr or not pw:
            messagebox.showwarning('Προσοχή',
                'Συμπλήρωσε username και κωδικό MySchool στις Ρυθμίσεις (⚙).',
                parent=self)
            return

        rids = self._dl_rids
        n = len(rids)
        btn_label = f'⬇  Λήψη {n} Αρχεί{"ου" if n == 1 else "ων"}'

        # Έλεγχος αν υπάρχουν ήδη αρχεία (πιθανόν από άλλον έλεγχο που
        # χρησιμοποιεί το ίδιο στατιστικό, π.χ. 2.1) ΠΡΙΝ ξεκινήσει η λήψη —
        # ρωτάμε αν θέλει επαναλήψη αντί να γίνεται σιωπηλή επαναχρησιμοποίηση.
        from core.downloader import get_downloads_dir, find_existing_reports
        dest = get_downloads_dir(self._base)
        existing_rids = find_existing_reports(dest, rids)
        force = []
        if existing_rids:
            redo = messagebox.askyesno(
                'Αρχεία υπάρχουν ήδη',
                f'Βρέθηκαν ήδη {len(existing_rids)}/{n} αρχεία σήμερα '
                '(πιθανόν από άλλον έλεγχο).\n\nΝα γίνει λήψη εκ νέου;',
                parent=self)
            if redo:
                force = existing_rids
            elif len(existing_rids) == n:
                # Όλα τα απαραίτητα αρχεία υπάρχουν ήδη και δεν θέλουμε νέα
                # λήψη — δεν υπάρχει λόγος να γίνει login στο MySchool.
                messagebox.showinfo('Λήψη',
                    f'Όλα τα {n} αρχεία υπάρχουν ήδη σήμερα — δεν χρειάζεται νέα λήψη.',
                    parent=self)
                return

        self._dl_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')

        def on_log(msg):
            self._log_append(self._dl_log, msg)

        def task():
            try:
                from core.downloader import MySchoolDownloader, cleanup_old_downloads
                dl = MySchoolDownloader(
                    usr, pw, dest, callback=on_log, reports=rids, force=force,
                    browser=getattr(self._cfg, 'BROWSER', 'chrome'))
                results = dl.run()
                ok = sum(1 for v in results.values() if v)
                cleanup_old_downloads(self._base, keep=1)
                self.after(0, lambda: [
                    self._dl_btn.configure(state='normal', bg=C['btn_bg'], text=btn_label),
                    messagebox.showinfo('Λήψη',
                        f'Ολοκληρώθηκε: {ok}/{n} αρχεία κατεβήκαν.', parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._dl_btn.configure(state='normal', bg=C['btn_bg'], text=btn_label),
                    messagebox.showerror('Σφάλμα Λήψης', m, parent=self)
                ])

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 1β: Λήψη για ελέγχους με CUSTOM_DOWNLOAD (π.χ. tmimata_genikis) ──
    #    Η λήψη γίνεται εδώ (πατώντας «⬇ Λήψη»), ΟΧΙ μέσα στο tab «Εκτέλεση» —
    #    το «Εκτέλεση» απλά ψάχνει τα ήδη κατεβασμένα αρχεία της ημέρας.

    def _build_custom_download(self, body, custom_dl):
        C = self._C
        required = getattr(self._mod, 'REQUIRED_REPORTS', [])

        self._section_lbl(body, 'Απαιτούμενα στατιστικά:')
        txt = ('\n'.join(f'  •  {r}' for r in required) if required
               else '  (δεν έχουν οριστεί σε αυτόν τον έλεγχο)')
        tk.Label(body, text=txt, bg=C['bg'], fg=C['desc'],
                 font=('Arial', 9), justify='left', anchor='w',
                 wraplength=560).pack(fill='x', pady=(0, 8))

        self._dl_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._dl_btn = self._run_btn(br, '⬇  Λήψη', lambda: self._start_custom_download(custom_dl))

    def _start_custom_download(self, custom_dl):
        C = self._C

        # Έλεγχος αν υπάρχουν ήδη αρχεία (πιθανόν από άλλον έλεγχο που
        # χρησιμοποιεί το ίδιο στατιστικό, π.χ. 3.1) ΠΡΙΝ ξεκινήσει η λήψη —
        # μόνο για ελέγχους που εκθέτουν CUSTOM_DOWNLOAD_INFO (π.χ.
        # tmimata_genikis) ώστε να ξέρουμε ΠΟΥ και ΤΙ να ελέγξουμε.
        force = []
        info_fn = getattr(self._mod, 'CUSTOM_DOWNLOAD_INFO', None)
        if callable(info_fn):
            try:
                dest_dir, rids = info_fn()
                from core.downloader import find_existing_reports
                existing_rids = find_existing_reports(dest_dir, rids)
                if existing_rids:
                    redo = messagebox.askyesno(
                        'Αρχεία υπάρχουν ήδη',
                        f'Βρέθηκαν ήδη {len(existing_rids)}/{len(rids)} αρχεία σήμερα '
                        '(πιθανόν από άλλον έλεγχο).\n\nΝα γίνει λήψη εκ νέου;',
                        parent=self)
                    if redo:
                        force = existing_rids
                    elif len(existing_rids) == len(rids):
                        # Όλα τα απαραίτητα αρχεία υπάρχουν ήδη και δεν θέλουμε
                        # νέα λήψη — δεν υπάρχει λόγος να γίνει login στο MySchool.
                        messagebox.showinfo('Λήψη',
                            f'Όλα τα {len(rids)} αρχεία υπάρχουν ήδη σήμερα — '
                            'δεν χρειάζεται νέα λήψη.', parent=self)
                        return
            except Exception:
                pass

        self._dl_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')
        try:
            self._dl_log.configure(state='normal')
            self._dl_log.delete('1.0', tk.END)
            self._dl_log.configure(state='disabled')
        except Exception:
            pass

        def on_log(msg):
            self._log_append(self._dl_log, msg)

        def task():
            try:
                try:
                    custom_dl(self._cfg, log=on_log, force=force)
                except TypeError:
                    # Έλεγχοι με CUSTOM_DOWNLOAD που δεν δέχονται force=
                    # (παλιότερη υπογραφή) — τρέχει κανονικά χωρίς αυτό.
                    custom_dl(self._cfg, log=on_log)
                self.after(0, lambda: [
                    self._dl_btn.configure(state='normal', bg=C['btn_bg'], text='⬇  Λήψη'),
                    messagebox.showinfo('Λήψη', 'Η λήψη ολοκληρώθηκε.', parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._dl_btn.configure(state='normal', bg=C['btn_bg'], text='⬇  Λήψη'),
                    messagebox.showerror('Σφάλμα Λήψης', m, parent=self)
                ])

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 2: Εκτέλεση ──────────────────────────────────────────────────────

    def _build_execute(self, body):
        C = self._C
        tk.Label(body,
                 text='Επεξεργάζεται τα κατεβασμένα στατιστικά και παράγει το αρχείο Excel.\n'
                      'Δεν στέλνει email — αυτό επιλέγεται στο επόμενο tab.',
                 bg=C['bg'], fg=C['desc'], font=('Arial', 8),
                 justify='left').pack(anchor='w', pady=(0, 6))

        self._ex_log = self._make_log(body, height=18)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._ex_btn = self._run_btn(br, '▶  Εκτέλεση Ελέγχου', self._start_execute)

    def _start_execute(self):
        C = self._C
        self._ex_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')
        try:
            self._ex_log.configure(state='normal')
            self._ex_log.delete('1.0', tk.END)
            self._ex_log.configure(state='disabled')
        except Exception:
            pass

        def on_log(msg):
            self._log_append(self._ex_log, msg)

        install_redirect, restore_print = self._redirect_print(on_log)

        def task():
            install_redirect()
            try:
                if self._is_custom_run:
                    self._mod.run(self._cfg)
                    self._exec_result = {'status': 'custom_run_done'}
                    on_log('\n✓ Ολοκληρώθηκε.')
                else:
                    from core.framework import execute_check
                    result = execute_check(self._mod, self._cfg)
                    self._exec_result = result
                    if result['status'] == 'missing_files':
                        req = result.get('required') or []
                        msg = 'Δεν βρέθηκαν τα απαραίτητα αρχεία.'
                        if req:
                            msg += '\nΑπαιτούνται: ' + ', '.join(req)
                        msg += '\n\nΚατέβασέ τα πρώτα από το tab «⬇ Λήψη».'
                        on_log('\n⚠ ' + msg)
                    elif result['status'] == 'empty':
                        on_log('\n✓ Δεν βρέθηκαν εγγραφές που χρήζουν προσοχής.\n'
                               'Ο έλεγχος ολοκληρώθηκε χωρίς θέματα.')
                    elif result['status'] == 'error':
                        on_log('\n✗ ' + result.get('message', 'Σφάλμα.'))
                    elif result['status'] == 'ok':
                        on_log('\n' + result['summary'])
            except Exception as e:
                import traceback
                err = f'{e}\n\n{traceback.format_exc()[-800:]}'
                on_log('\n✗ Σφάλμα: ' + err)
                self._exec_result = {'status': 'error', 'message': str(e)}
            finally:
                restore_print()

                def _done():
                    self._ex_btn.configure(state='normal', bg=C['btn_bg'],
                                            text='▶  Εκτέλεση Ελέγχου')
                    if self._exec_result and self._exec_result.get('status') == 'ok' \
                            and not self._is_custom_run:
                        # Παράθυρο αποτελεσμάτων μετά την Εκτέλεση — ίδιο στυλ
                        # με τον έλεγχο Τμημάτων (tmimata_genikis), για όλους
                        # τους «απλούς» ελέγχους (ανεξάρτητα από το αν έχουν
                        # tab «Αποστολή» — π.χ. εσωτερικούς ελέγχους με
                        # NO_SEND_TAB).
                        try:
                            self._show_generic_results_popup(self._exec_result)
                        except Exception:
                            import traceback
                            traceback.print_exc()
                    if not self._has_send_tab:
                        pass
                    elif self._is_custom_run:
                        custom_send = getattr(self._mod, 'CUSTOM_SEND_TAB', None)
                        if callable(custom_send):
                            # Ο έλεγχος εκθέτει δικό του σημείο αποστολής (π.χ.
                            # tmimata_genikis) — χτίζουμε τη φόρμα αποστολής
                            # ΑΠΕΥΘΕΙΑΣ μέσα στο tab «Αποστολή» (χωρίς popup).
                            self._send_hint.configure(text='')
                            container = self._custom_send_body
                            # Καθαρίζουμε τυχόν προηγούμενη φόρμα, ώστε να μην
                            # διπλασιάζονται τα widgets αν ξανατρέξει η Εκτέλεση.
                            for child in list(container.winfo_children()):
                                child.destroy()
                            try:
                                custom_send(container, self._cfg)
                            except Exception as e:
                                import traceback
                                traceback.print_exc()
                                tk.Label(container, text=f'✗ Σφάλμα: {e}',
                                         bg=C['bg'], fg='#B00020',
                                         font=('Arial', 9)).pack(anchor='w', pady=8)
                        else:
                            self._lock_send_tab(True)
                            self._send_hint.configure(
                                text='Αυτός ο έλεγχος διαχειρίζεται μόνος του την αποστολή '
                                     'email μέσα στο tab «Εκτέλεση» (θα σου ζητήσει τις '
                                     'επιλογές εκεί) — το tab «Αποστολή» δεν χρησιμοποιείται.')
                    else:
                        can_send = bool(self._exec_result
                                        and self._exec_result.get('status') == 'ok'
                                        and self._exec_result.get('has_email'))
                        if getattr(self, '_has_generic_split', False):
                            try:
                                self._split_btn.configure(
                                    state='normal' if can_send else 'disabled',
                                    bg=C['btn_bg'] if can_send else C['btn_dis'])
                            except Exception:
                                pass
                            try:
                                self._detect_split_file()
                            except Exception:
                                pass
                            self._split_hint.configure(
                                text=('Πάτησε «✂ Διαχωρισμός ανά Σχολείο» για να δεις τι '
                                      'αρχείο θα σταλεί σε κάθε σχολείο.' if can_send else
                                      'Τρέξε πρώτα την «▶ Εκτέλεση» — μετά ενεργοποιείται '
                                      'ο διαχωρισμός εδώ.'))
                        self._lock_send_tab(not can_send)
                self.after(0, _done)

        threading.Thread(target=task, daemon=True).start()

    def _show_generic_results_popup(self, result):
        """
        Παράθυρο αποτελεσμάτων μετά από επιτυχή Εκτέλεση («status: ok») —
        ίδιο στυλ με το παράθυρο αποτελεσμάτων του ελέγχου Τμημάτων
        (checks/tmimata_genikis.py::_show_results_dialog): σύνοψη + κουμπί
        «Άνοιγμα Excel». Καλείται από _start_execute._done() στο main thread
        (μέσω self.after), οπότε μπορεί να φτιάξει widgets απευθείας.
        """
        title    = result.get('title', '?')
        summary  = result.get('summary', '')
        path_all = result.get('path_all')

        win = tk.Toplevel(self)
        win.title(f'Αποτελέσματα — {title}')
        win.configure(bg='#FFF8E1')
        win.resizable(True, True)
        win.attributes('-topmost', True)
        # Το CheckRunDialog κάνει ήδη grab_set() στον εαυτό του — χωρίς δικό
        # της grab, η νέα Toplevel δεν λαμβάνει clicks μέχρι να κλείσει το
        # CheckRunDialog (ίδιο fix με tmimata_genikis._show_results_dialog).
        win.transient(self)
        win.grab_set()
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f'520x380+{sw//2-260}+{sh//2-190}')

        hdr = tk.Frame(win, bg='#E65100', pady=8)
        hdr.pack(fill='x')
        tk.Label(hdr, text=f'⚠  {title}',
                 bg='#E65100', fg='white',
                 font=('Arial', 11, 'bold')).pack()

        txt = scrolledtext.ScrolledText(win, font=('Arial', 9), wrap='word',
                                         relief='flat', bg='#FFF8E1', height=10)
        txt.pack(fill='both', expand=True, padx=14, pady=8)
        txt.insert('1.0', summary)
        txt.config(state='disabled')

        btn_f = tk.Frame(win, bg='#FFF8E1')
        btn_f.pack(pady=(0, 12))

        def _open_excel():
            if not path_all:
                messagebox.showinfo('Άνοιγμα Excel', 'Δεν υπάρχει αρχείο.', parent=win)
                return
            try:
                os.startfile(os.path.normpath(path_all))
            except Exception as e:
                messagebox.showerror('Σφάλμα',
                                      f'Δεν ήταν δυνατό το άνοιγμα του αρχείου:\n{e}',
                                      parent=win)

        tk.Button(btn_f, text='📄 Άνοιγμα Excel',
                  bg='#E65100', fg='white',
                  font=('Arial', 9, 'bold'), relief='flat',
                  padx=14, pady=5, cursor='hand2',
                  command=_open_excel).pack(side='left', padx=4)
        tk.Button(btn_f, text='Κλείσιμο',
                  bg='#E8EDF3', fg='#333333',
                  font=('Arial', 9), relief='flat',
                  padx=14, pady=5, cursor='hand2',
                  command=win.destroy).pack(side='left', padx=4)

    # ── Tab «✂ Διαχωρισμός» (γενικό, για ελέγχους χωρίς CUSTOM_RUN) ──────────

    def _generic_split_applicable(self):
        """
        True όταν αυτός ο έλεγχος πρέπει να πάρει το γενικό tab
        «✂ Διαχωρισμός» (χωρίζει το αρχείο σε ένα Excel ανά σχολείο πριν την
        «Αποστολή»): έλεγχοι χωρίς CUSTOM_RUN, με HAS_EMAIL, που ΔΕΝ είναι
        TEST_ONLY (δεν αφορούν σχολεία). Ελέγχους με ειδική πηγή δεδομένων
        για τον διαχωρισμό (π.χ. apontes_xwris_adeia — ζητά επεξεργασμένο
        αρχείο) τους καλύπτει επίσης το ίδιο γενικό tab, μέσω
        CUSTOM_SPLIT_SOURCE (core/framework.py::split_exec_result). Το
        παλιό custom_full_send hook εξαιρείται εδώ μόνο σαν προστασία, σε
        περίπτωση που κάποιο μελλοντικό module το ξαναχρησιμοποιήσει χωρίς
        να περάσει από το γενικό tab.
        """
        if self._is_custom_run:
            return False
        if not getattr(self._mod, 'HAS_EMAIL', False):
            return False
        if getattr(self._mod, 'TEST_ONLY', False):
            return False
        if callable(getattr(self._mod, 'custom_full_send', None)):
            return False
        return True

    def _build_generic_split(self, body):
        """
        Tab «✂ Διαχωρισμός» για τους «απλούς» ελέγχους (χωρίς CUSTOM_RUN) —
        χωρίζει το συνολικό αρχείο αποτελεσμάτων σε ένα Excel ανά σχολείο
        μέσα σε φάκελο «split», ΠΡΙΝ την «Αποστολή» — ώστε να φαίνεται τι
        αρχείο θα σταλεί σε κάθε σχολείο. Προς το παρόν δεν εξαιρεί κανένα
        σχολείο (θα προστεθεί αργότερα φιλτράρισμα όπου χρειάζεται).
        Ξεκλειδώνεται μόλις τρέξει επιτυχώς η Εκτέλεση (βλ. _start_execute/_done).
        """
        C = self._C
        tk.Label(body,
                 text='Χωρίζει το συνολικό αρχείο αποτελεσμάτων σε ένα Excel ανά '
                      'σχολείο (φάκελος «split»), ώστε να δεις τι αρχείο θα '
                      'σταλεί σε κάθε σχολείο πριν την αποστολή. Στο tab '
                      '«✉ Αποστολή» θα σταλούν emails μόνο σε σχολεία που έχουν '
                      'ατομικό αρχείο εδώ.',
                 bg=C['bg'], fg=C['desc'], font=('Arial', 8),
                 wraplength=560, justify='left', anchor='w').pack(fill='x', pady=(0, 8))

        self._split_hint = tk.Label(
            body,
            text='Τρέξε πρώτα την «▶ Εκτέλεση» — μετά ενεργοποιείται ο διαχωρισμός εδώ.',
            bg=C['bg'], fg=C['desc'], font=('Arial', 8, 'italic'),
            justify='left', wraplength=560)
        self._split_hint.pack(anchor='w', pady=(0, 8))

        # Ανίχνευση/εμφάνιση του αρχείου προς διαχωρισμό — ίδια λογική με
        # το tab Διαχωρισμός του ελέγχου Τμημάτων Γενικής (βλ.
        # checks/tmimata_genikis.py::build_split_tab / _find_latest_output).
        # Εδώ η «ανίχνευση» κοιτάζει το αποτέλεσμα της τελευταίας
        # Εκτέλεσης μέσα σε αυτή τη σύνοδο (self._exec_result['path_all']).
        self._split_info_lbl = tk.Label(
            body, text='', bg=C['bg'], fg=C['desc'],
            font=('Arial', 8, 'italic'), anchor='w',
            justify='left', wraplength=560)
        self._split_info_lbl.pack(fill='x', pady=(0, 4))

        self._split_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._split_btn = self._run_btn(br, '✂  Διαχωρισμός ανά Σχολείο',
                                         self._start_generic_split)
        self._split_btn.configure(state='disabled', bg=C['btn_dis'])
        tk.Button(br, text='Ανίχνευση αρχείου', bg=C['bg2'], fg=C['desc'],
                  font=('Arial', 9), relief='flat', padx=10, pady=5,
                  cursor='hand2',
                  command=self._detect_split_file).pack(side='left', padx=(8, 0))

        self._detect_split_file()

    def _detect_split_file(self):
        """
        Ενημερώνει το `self._split_info_lbl` με το αρχείο που θα χωριστεί —
        το συγκεντρωτικό αρχείο αποτελεσμάτων (`path_all`) της τελευταίας
        επιτυχούς Εκτέλεσης σε αυτή τη σύνοδο. Καλείται στο χτίσιμο του tab,
        μετά από κάθε Εκτέλεση, και όταν πατηθεί το κουμπί «Ανίχνευση
        αρχείου».
        """
        path_all = None
        if self._exec_result and self._exec_result.get('status') == 'ok':
            path_all = self._exec_result.get('path_all')
        if path_all and os.path.isfile(path_all):
            self._split_info_lbl.configure(
                text=f'✓ {os.path.basename(path_all)}', fg='#2E7D32')
        else:
            self._split_info_lbl.configure(
                text='Δεν βρέθηκε αρχείο αποτελεσμάτων — τρέξε πρώτα την «▶ Εκτέλεση».',
                fg='#B00020')

    def _start_generic_split(self):
        C = self._C
        if not self._exec_result or self._exec_result.get('status') != 'ok':
            messagebox.showwarning('Προσοχή', 'Τρέξε πρώτα την Εκτέλεση.', parent=self)
            return

        self._split_btn.configure(state='disabled', bg=C['btn_dis'], text='Διαχωρισμός...')
        try:
            self._split_log.configure(state='normal')
            self._split_log.delete('1.0', tk.END)
            self._split_log.configure(state='disabled')
        except Exception:
            pass

        def on_log(msg):
            self._log_append(self._split_log, msg)

        install_redirect, restore_print = self._redirect_print(on_log)

        def task():
            install_redirect()
            try:
                from core.framework import split_exec_result
                n = split_exec_result(self._exec_result, log=on_log)
                self.after(0, lambda: [
                    self._split_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='✂  Διαχωρισμός ανά Σχολείο'),
                    messagebox.showinfo('Διαχωρισμός', f'Ολοκληρώθηκε! {n} αρχεία.',
                                         parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._split_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='✂  Διαχωρισμός ανά Σχολείο'),
                    messagebox.showerror('Σφάλμα Διαχωρισμού', m, parent=self)
                ])
            finally:
                restore_print()

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 3: Αποστολή ──────────────────────────────────────────────────────

    def _build_send(self, body):
        C = self._C

        # Έλεγχοι με CUSTOM_SEND_TAB (π.χ. tmimata_genikis) διαχειρίζονται τη
        # δική τους λογική αποστολής (ανά σχολείο, με βάση τις αποκλίσεις) —
        # δεν ταιριάζει το γενικό Δοκιμαστική/Κανονική radio button παρακάτω.
        custom_send = getattr(self._mod, 'CUSTOM_SEND_TAB', None)
        if self._is_custom_run and callable(custom_send):
            self._build_custom_send(body, custom_send)
            return

        self._send_mode = tk.StringVar(value='test')
        _from_email = getattr(self._cfg, 'FROM_EMAIL', '') or '...'

        # Επεξεργασία προτύπου email — μετακομισμένο εδώ από το κουμπί ✏ που
        # υπήρχε παλιότερα δίπλα σε κάθε έλεγχο στο tab «Έλεγχοι» (main.py).
        if getattr(self._mod, 'HAS_EMAIL', False):
            tmpl_row = tk.Frame(body, bg=C['bg'])
            tmpl_row.pack(fill='x', pady=(0, 8))
            tk.Button(tmpl_row, text='✏  Πρότυπο Email',
                      bg=C['bg2'], fg=C['hdr_bg'],
                      font=('Arial', 9), relief='flat', cursor='hand2',
                      padx=10, pady=4,
                      activebackground=C['sel_bg'],
                      command=self._open_email_editor).pack(side='left')

        self._send_hint = tk.Label(
            body,
            text='Τρέξε πρώτα την «▶ Εκτέλεση» — μετά ενεργοποιείται η αποστολή εδώ.',
            bg=C['bg'], fg=C['desc'], font=('Arial', 8, 'italic'),
            justify='left', wraplength=560)
        self._send_hint.pack(anchor='w', pady=(0, 8))

        opt = tk.Frame(body, bg=C['bg'])
        opt.pack(fill='x', pady=(0, 6))
        tk.Radiobutton(opt, text=f'Δοκιμαστική αποστολή (test) — στο: {_from_email}',
                       variable=self._send_mode, value='test',
                       bg=C['bg'], selectcolor=C['sel_bg'],
                       activebackground=C['bg'], font=('Arial', 9)).pack(anchor='w')
        tk.Radiobutton(opt, text='Κανονική αποστολή σε όλα τα σχολεία',
                       variable=self._send_mode, value='full',
                       bg=C['bg'], selectcolor=C['sel_bg'],
                       activebackground=C['bg'], font=('Arial', 9)).pack(anchor='w')

        self._send_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._send_btn = self._run_btn(br, '✉  Αποστολή', self._start_send)

    def _build_custom_send(self, body, custom_send):
        """Tab «Αποστολή» για ελέγχους με CUSTOM_SEND_TAB (π.χ. tmimata_genikis)
        — η φόρμα αποστολής του ελέγχου χτίζεται ΑΠΕΥΘΕΙΑΣ μέσα στο tab
        (χωρίς ξεχωριστό popup), μόλις τρέξει επιτυχώς η Εκτέλεση
        (βλ. _start_execute/_done)."""
        C = self._C
        self._send_hint = tk.Label(
            body,
            text='Τρέξε πρώτα την «▶ Εκτέλεση» — μετά ενεργοποιείται η αποστολή εδώ.',
            bg=C['bg'], fg=C['desc'], font=('Arial', 8, 'italic'),
            justify='left', wraplength=560)
        self._send_hint.pack(anchor='w', pady=(0, 10))

        # Container στον οποίο ο έλεγχος (custom_send) χτίζει τη δική του
        # φόρμα αποστολής απευθείας — γεμίζει μόλις ολοκληρωθεί η Εκτέλεση.
        self._custom_send_body = tk.Frame(body, bg=C['bg'])
        self._custom_send_body.pack(fill='both', expand=True)

    def _lock_send_tab(self, locked):
        state = 'disabled' if locked else 'normal'
        try:
            self._send_btn.configure(state=state)
        except Exception:
            pass
        if locked:
            self._send_hint.configure(
                text='Τρέξε πρώτα την «▶ Εκτέλεση» — μετά ενεργοποιείται η αποστολή εδώ.')
        else:
            self._send_hint.configure(
                text='Επίλεξε τρόπο αποστολής και πάτησε «Αποστολή».')

    def _start_send(self):
        C = self._C
        if not self._exec_result or self._exec_result.get('status') != 'ok':
            messagebox.showwarning('Προσοχή', 'Τρέξε πρώτα την Εκτέλεση.', parent=self)
            return
        if not getattr(self._cfg, 'FROM_PASSWORD', '').strip():
            messagebox.showwarning('Προσοχή',
                'Ο κωδικός email δεν έχει οριστεί.\nΠήγαινε στις Ρυθμίσεις (⚙).',
                parent=self)
            return

        test_mode = (self._send_mode.get() == 'test')

        if not test_mode:
            ok = messagebox.askyesno('Επιβεβαίωση Αποστολής',
                'Θα αποσταλούν emails σε ΟΛΑ τα σχολεία που έχουν αποκλίσεις.\n\nΣυνέχεια;',
                parent=self)
            if not ok:
                return

        self._send_btn.configure(state='disabled', bg=C['btn_dis'], text='Αποστολή...')

        def on_log(msg):
            self._log_append(self._send_log, msg)

        install_redirect, restore_print = self._redirect_print(on_log)

        def task():
            install_redirect()
            try:
                from core.framework import send_from_exec_result
                send_from_exec_result(self._exec_result, test_mode)
                self.after(0, lambda: [
                    self._send_btn.configure(state='normal', bg=C['btn_bg'], text='✉  Αποστολή'),
                    messagebox.showinfo('Αποστολή', 'Η αποστολή ολοκληρώθηκε!', parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._send_btn.configure(state='normal', bg=C['btn_bg'], text='✉  Αποστολή'),
                    messagebox.showerror('Σφάλμα Αποστολής', m, parent=self)
                ])
            finally:
                restore_print()

        threading.Thread(target=task, daemon=True).start()

    # ── Πρότυπο Email (μετακομισμένο από main.py — πρώην κουμπί ✏ δίπλα σε
    #    κάθε έλεγχο στο tab «Έλεγχοι») ──────────────────────────────────────

    def _get_default_email_body(self):
        """Επιστρέφει το default body text του ελέγχου, χωρίς υπογραφή."""
        body_t = getattr(self._mod, 'EMAIL_BODY', '')
        try:
            full = body_t('') if callable(body_t) else body_t
            sig  = self._cfg.email_signature()
            if sig and full.endswith(sig):
                return full[:-len(sig)]
            return full
        except Exception:
            return ''

    def _open_email_editor(self):
        """Dialog επεξεργασίας email template για αυτόν τον έλεγχο."""
        C = self._C
        mod_name  = self._mod.__name__.split('.')[-1]
        title_str = getattr(self._mod, 'CHECK_TITLE', mod_name)

        settings  = _load_local_settings()
        templates = settings.get('email_templates', {})
        custom    = templates.get(mod_name)

        if custom:
            cur_subject = custom.get('subject', '')
            cur_body    = custom.get('body', '')
        else:
            cur_subject = getattr(self._mod, 'EMAIL_SUBJECT', '')
            cur_body    = self._get_default_email_body()

        dlg = tk.Toplevel(self)
        dlg.title(f'Πρότυπο Email — {title_str}')
        dlg.configure(bg=C['bg'])
        dlg.resizable(True, False)
        dlg.grab_set()
        dlg.transient(self)

        pad = dict(padx=14, pady=5)

        tk.Label(dlg, text='Θέμα:', bg=C['bg'], fg=C['hdr_bg'],
                 font=('Arial', 9, 'bold'), anchor='w').pack(fill='x', **pad)

        subj_var = tk.StringVar(value=cur_subject)
        tk.Entry(dlg, textvariable=subj_var, font=('Arial', 9),
                 width=60).pack(fill='x', padx=14, pady=(0, 8))

        tk.Label(dlg, text='Κείμενο email:', bg=C['bg'], fg=C['hdr_bg'],
                 font=('Arial', 9, 'bold'), anchor='w').pack(fill='x', **pad)

        txt = tk.Text(dlg, font=('Arial', 9), width=60, height=10,
                      wrap='word', relief='solid', bd=1)
        txt.pack(fill='x', padx=14, pady=(0, 4))
        txt.insert('1.0', cur_body)

        tk.Label(dlg, text='Η υπογραφή σας προστίθεται αυτόματα στο τέλος.',
                 bg=C['bg'], fg=C['desc'], font=('Arial', 8),
                 anchor='w').pack(fill='x', padx=14, pady=(0, 10))

        def _save():
            new_subj = subj_var.get().strip()
            new_body = txt.get('1.0', 'end-1c')
            s = _load_local_settings()
            s.setdefault('email_templates', {})[mod_name] = {
                'subject': new_subj,
                'body':    new_body,
            }
            _save_local_settings(s)
            dlg.destroy()
            messagebox.showinfo('Αποθήκευση', 'Το πρότυπο email αποθηκεύτηκε.',
                                parent=self)

        def _reset():
            if messagebox.askyesno('Επαναφορά', 'Να επανέλθει το προεπιλεγμένο κείμενο;',
                                   parent=dlg):
                s = _load_local_settings()
                s.get('email_templates', {}).pop(mod_name, None)
                _save_local_settings(s)
                dlg.destroy()

        btn_row = tk.Frame(dlg, bg=C['bg'])
        btn_row.pack(pady=(0, 12))

        tk.Button(btn_row, text='Αποθήκευση',
                  bg=C['btn_bg'], fg=C['btn_fg'],
                  font=('Arial', 9, 'bold'), relief='flat',
                  padx=14, pady=5, cursor='hand2',
                  command=_save).pack(side='left', padx=4)

        tk.Button(btn_row, text='Επαναφορά προεπιλογής',
                  bg=C['bg2'], fg=C['hdr_bg'],
                  font=('Arial', 9), relief='flat',
                  padx=14, pady=5, cursor='hand2',
                  command=_reset).pack(side='left', padx=4)

        tk.Button(btn_row, text='Άκυρο',
                  bg=C['bg2'], fg=C['desc'],
                  font=('Arial', 9), relief='flat',
                  padx=14, pady=5, cursor='hand2',
                  command=dlg.destroy).pack(side='left', padx=4)

        dlg.update_idletasks()
        w, h = dlg.winfo_width(), dlg.winfo_height()
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        dlg.geometry(f'+{x}+{y}')
