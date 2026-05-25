"""
smeae/dialog.py
═══════════════
Tkinter Toplevel dialog για την ενσωματωμένη λειτουργικότητα ΣΜΕΑΕ.
4 tabs: Λήψη / Σύγκριση / Διαχωρισμός / Email
"""
import os, glob, threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from datetime import datetime


class SmeaeDialog(tk.Toplevel):

    def __init__(self, parent, config_ref, base_dir, C):
        super().__init__(parent)
        self._cfg      = config_ref
        self._base_dir = base_dir
        self._C        = C

        self.title('Έλεγχος Ε.Ε.Α. — Στατιστικά')
        self.configure(bg=C['bg'])
        self.resizable(False, False)
        self.grab_set()
        self.transient(parent)

        ico = os.path.join(base_dir, 'app.ico')
        if os.path.exists(ico):
            try:
                self.iconbitmap(ico)
            except Exception:
                pass

        self._diff_file = None   # ανιχνευμένο αρχείο για split

        self._build()
        self.update_idletasks()
        self.geometry('600x640')
        pw = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        ph = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f'+{pw}+{ph}')

    # ── Βοηθητικά paths ──────────────────────────────────────────────────────

    def _dl_dir(self, year):
        return os.path.join(self._base_dir, 'smeae_downloads', year)

    def _out_dir(self, year):
        return os.path.join(self._base_dir, 'smeae_output', year)

    def _split_dir(self, year):
        return os.path.join(self._base_dir, 'smeae_output', year, 'split')

    def _mappings_path(self):
        return os.path.join(self._base_dir, 'data', 'smeae_column_mappings.json')

    def _guess_year(self):
        now = datetime.now()
        return f'{now.year}-{now.year + 1}' if now.month >= 9 else f'{now.year - 1}-{now.year}'

    # ── Widget helpers ────────────────────────────────────────────────────────

    def _log_append(self, widget, msg):
        def _do():
            widget.configure(state='normal')
            widget.insert(tk.END, msg + '\n')
            widget.see(tk.END)
            widget.configure(state='disabled')
        self.after(0, _do)

    def _make_log(self, parent):
        log = scrolledtext.ScrolledText(
            parent, height=12, font=('Consolas', 8),
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

    def _year_field(self, parent, var):
        C = self._C
        row = tk.Frame(parent, bg=C['bg'])
        row.pack(anchor='w', pady=(4, 10))
        tk.Entry(row, textvariable=var, width=12,
                 font=('Arial', 10), relief='solid', bd=1).pack(side='left')
        tk.Label(row, text='  (π.χ. 2024-2025)',
                 bg=C['bg'], fg=C['footer'],
                 font=('Arial', 8)).pack(side='left')

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        C = self._C

        # Header
        hdr = tk.Frame(self, bg='#0F6E56', pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='📊  Έλεγχος Ε.Ε.Α. — Στατιστικά Ειδικών Εκπαιδευτικών Αναγκών',
                 bg='#0F6E56', fg='white',
                 font=('Arial', 11, 'bold')).pack()

        # Notebook
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
        t3 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
        t4 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
        nb.add(t1, text='  ⬇ Λήψη  ')
        nb.add(t2, text='  🔍 Σύγκριση  ')
        nb.add(t3, text='  ✂ Διαχωρισμός  ')
        nb.add(t4, text='  ✉ Email  ')

        self._build_download(t1)
        self._build_compare(t2)
        self._build_split(t3)
        self._build_email(t4)

        # Footer
        foot = tk.Frame(self, bg=C['bg2'], pady=10)
        foot.pack(fill='x')
        tk.Button(foot, text='Κλείσιμο',
                  bg=C['bg2'], fg=C['desc'],
                  font=('Arial', 10), relief='flat', padx=12, pady=5,
                  cursor='hand2', command=self.destroy).pack(side='right', padx=16)

    # ── Tab 1: Λήψη ──────────────────────────────────────────────────────────

    def _build_download(self, body):
        C = self._C
        self._dl_year = tk.StringVar(value=self._guess_year())
        self._section_lbl(body, 'Σχολικό έτος:')
        self._year_field(body, self._dl_year)
        tk.Label(body,
                 text='Κατεβάζει 10 αρχεία από το MySchool:\n'
                      '• 9 στατιστικά ΣΜΕΑΕ/ΕΕΑ\n'
                      '• Βασικά Στοιχεία Σχολικών Μονάδων (για αποστολή email)\n'
                      'Χρησιμοποιεί τα credentials από τις Ρυθμίσεις (⚙).',
                 bg=C['bg'], fg=C['desc'],
                 font=('Arial', 8), justify='left').pack(anchor='w', pady=(0, 6))
        self._dl_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._dl_btn = self._run_btn(br, '⬇  Λήψη 10 Αρχείων', self._start_download)

    def _start_download(self):
        C   = self._C
        usr = getattr(self._cfg, 'MYSCHOOL_USER', '').strip()
        pw  = getattr(self._cfg, 'MYSCHOOL_PASS', '').strip()
        if not usr or not pw:
            messagebox.showwarning('Προσοχή',
                'Συμπλήρωσε username και κωδικό MySchool στις Ρυθμίσεις (⚙).',
                parent=self)
            return

        year = self._dl_year.get().strip()
        if not year:
            messagebox.showwarning('Προσοχή', 'Εισάγετε σχολικό έτος.', parent=self)
            return

        dest = self._dl_dir(year)
        self._dl_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')

        def on_log(msg):
            self._log_append(self._dl_log, msg)

        def task():
            try:
                from smeae.downloader import SmeaeDownloader
                dl = SmeaeDownloader(usr, pw, dest, callback=on_log)
                results = dl.run()
                ok = sum(1 for v in results.values() if v)
                self.after(0, lambda: [
                    self._dl_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='⬇  Λήψη 10 Αρχείων'),
                    messagebox.showinfo(
                        'Λήψη ΣΜΕΑΕ',
                        f'Ολοκληρώθηκε: {ok}/10 αρχεία κατεβήκαν.',
                        parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._dl_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='⬇  Λήψη 10 Αρχείων'),
                    messagebox.showerror('Σφάλμα Λήψης', m, parent=self)
                ])

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 2: Σύγκριση ──────────────────────────────────────────────────────

    def _build_compare(self, body):
        C = self._C
        self._cmp_year = tk.StringVar(value=self._guess_year())
        self._section_lbl(body, 'Σχολικό έτος (ίδιο με λήψη):')
        self._year_field(body, self._cmp_year)
        tk.Label(body,
                 text='Χρησιμοποιεί αυτόματα:\n'
                      '• master: αρχείο "1. ..." από smeae_downloads/{έτος}/\n'
                      '• slave: αρχεία "2."-"9." από τον ίδιο φάκελο\n'
                      '• αρχείο αντιστοιχίας: data/smeae_column_mappings.json\n'
                      'Αποτελέσματα → smeae_output/{έτος}/',
                 bg=C['bg'], fg=C['desc'],
                 font=('Arial', 8), justify='left').pack(anchor='w', pady=(0, 6))
        self._cmp_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._cmp_btn = self._run_btn(br, '🔍  Εκτέλεση Σύγκρισης', self._start_compare)

    def _start_compare(self):
        C    = self._C
        year = self._cmp_year.get().strip()
        if not year:
            messagebox.showwarning('Προσοχή', 'Εισάγετε σχολικό έτος.', parent=self)
            return

        dl_dir  = self._dl_dir(year)
        out_dir = self._out_dir(year)
        jpath   = self._mappings_path()

        if not os.path.isdir(dl_dir):
            messagebox.showwarning('Προσοχή',
                f'Φάκελος λήψης δεν βρέθηκε:\n{dl_dir}\n\nΤρέξτε πρώτα τη Λήψη.',
                parent=self)
            return
        if not os.path.exists(jpath):
            messagebox.showwarning('Προσοχή',
                f'Αρχείο αντιστοιχίας στηλών δεν βρέθηκε:\n{jpath}',
                parent=self)
            return

        self._cmp_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')

        def on_log(msg):
            self._log_append(self._cmp_log, msg)

        def task():
            try:
                import json as _json
                from smeae.compare import (open_xlsx, match_columns,
                                            compare_xlsx, write_to_excel)
                with open(jpath, 'r', encoding='utf-8') as f:
                    col_maps = _json.load(f)

                all_xls = sorted(
                    glob.glob(os.path.join(dl_dir, '*.xls')) +
                    glob.glob(os.path.join(dl_dir, '*.xlsx'))
                )
                master_files = [f for f in all_xls
                                if os.path.basename(f).startswith('1.')]
                # Αποκλεισμός αρχείου 10. (Βασικά Στοιχεία Σχολικών Μονάδων)
                # — έχει τελείως διαφορετική δομή από τα ΕΕΑ στατιστικά
                slave_files  = [f for f in all_xls
                                if not os.path.basename(f).startswith('1.')
                                and not os.path.basename(f).startswith('10.')]

                if not master_files:
                    raise RuntimeError(
                        'Αρχείο master (1. ...) δεν βρέθηκε στον φάκελο λήψης.')
                if not slave_files:
                    raise RuntimeError(
                        'Αρχεία slave (2.-9. ...) δεν βρέθηκαν στον φάκελο λήψης.')

                master_file = master_files[0]
                on_log(f'Master: {os.path.basename(master_file)}')
                on_log(f'Slave αρχεία: {len(slave_files)}')
                on_log('Φόρτωση master...')

                dfm = open_xlsx(master_file)
                dfm.dropna(how='any', inplace=True)
                on_log(f'  {len(dfm)} εγγραφές master.')

                # Διαγραφή παλιού αρχείου αποτελεσμάτων — κάθε run ξεκινά fresh
                import glob as _gl2
                from datetime import datetime as _dt2
                _today = _dt2.now().strftime('%Y%m%d')
                for _old in _gl2.glob(os.path.join(out_dir, f'differences_{year}_{_today}.xlsx')):
                    try:
                        os.remove(_old)
                        on_log(f'  (Διαγράφηκε παλιό αρχείο: {os.path.basename(_old)})')
                    except Exception:
                        pass

                for sf in slave_files:
                    sname = os.path.basename(sf)
                    on_log(f'\nΣύγκριση: {sname}')
                    dfs = open_xlsx(sf)
                    dfs.dropna(how='any', inplace=True)
                    on_log(f'  {len(dfs)} εγγραφές slave.')

                    match_idx = col_maps['slaves'].get(sname, {})
                    col_l1, col_l2 = match_columns(dfm, dfs, match_idx)
                    diffs   = compare_xlsx(master_file, sf, dfm, dfs, col_l1, col_l2)
                    outfile = write_to_excel(diffs, sf, out_dir, year)
                    on_log(f'  → {os.path.basename(outfile)}')

                # Βρες το τελευταίο αρχείο αποτελεσμάτων
                import glob as _glob
                _diff_files = sorted(_glob.glob(os.path.join(out_dir, 'differences_*.xlsx')))
                _diff_last  = _diff_files[-1] if _diff_files else None
                if _diff_last:
                    on_log(f'\n📄 Αρχείο: {_diff_last}')
                on_log('\n✓ Σύγκριση ολοκληρώθηκε!')

                def _done():
                    self._cmp_btn.configure(state='normal', bg=C['btn_bg'],
                                            text='🔍  Εκτέλεση Σύγκρισης')
                    if _diff_last:
                        ans = messagebox.askyesno('Σύγκριση',
                            'Η σύγκριση ολοκληρώθηκε!\n\nΆνοιγμα αρχείου αποτελεσμάτων;',
                            parent=self)
                        if ans:
                            os.startfile(_diff_last)
                    else:
                        messagebox.showinfo('Σύγκριση', 'Η σύγκριση ολοκληρώθηκε!',
                                            parent=self)
                self.after(0, _done)
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._cmp_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='🔍  Εκτέλεση Σύγκρισης'),
                    messagebox.showerror('Σφάλμα Σύγκρισης', m, parent=self)
                ])

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 3: Διαχωρισμός ───────────────────────────────────────────────────

    def _build_split(self, body):
        C = self._C
        self._split_year = tk.StringVar(value=self._guess_year())
        self._section_lbl(body, 'Σχολικό έτος:')
        self._year_field(body, self._split_year)
        tk.Label(body,
                 text='Ανιχνεύει αυτόματα το τελευταίο αρχείο differences_*.xlsx\n'
                      'από smeae_output/{έτος}/  →  αποθηκεύει σε .../split/',
                 bg=C['bg'], fg=C['desc'],
                 font=('Arial', 8), justify='left').pack(anchor='w', pady=(0, 4))

        self._split_info = tk.Label(body, text='', bg=C['bg'],
                                     fg=C['status_ok'], font=('Arial', 8, 'italic'))
        self._split_info.pack(anchor='w', pady=(0, 4))

        self._split_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._split_btn = self._run_btn(br, '✂  Διαχωρισμός ανά Σχολείο',
                                         self._start_split)
        tk.Button(br, text='Ανίχνευση αρχείου',
                  bg=C['bg2'], fg=C['desc'],
                  font=('Arial', 9), relief='flat', padx=10, pady=6,
                  cursor='hand2',
                  command=self._detect_diff).pack(side='right', padx=(0, 6), pady=(8, 0))

    def _detect_diff(self):
        C    = self._C
        year = self._split_year.get().strip()
        out  = self._out_dir(year)
        files = sorted(glob.glob(os.path.join(out, 'differences_*.xlsx')), reverse=True)
        if files:
            self._diff_file = files[0]
            self._split_info.configure(
                text=f'✓ {os.path.basename(files[0])}', fg=C['status_ok'])
        else:
            self._diff_file = None
            self._split_info.configure(
                text='Δεν βρέθηκε αρχείο differences_*.xlsx', fg=C['status_err'])

    def _start_split(self):
        C    = self._C
        year = self._split_year.get().strip()
        if not self._diff_file:
            self._detect_diff()
        if not self._diff_file:
            messagebox.showwarning('Προσοχή',
                'Δεν βρέθηκε αρχείο αποτελεσμάτων.\n'
                'Τρέξτε πρώτα τη Σύγκριση ή πατήστε "Ανίχνευση αρχείου".',
                parent=self)
            return

        split_out = self._split_dir(year)
        diff_file = self._diff_file
        self._split_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')

        def on_log(msg):
            self._log_append(self._split_log, msg)

        def task():
            try:
                from smeae.compare import split_xlsx
                # Redirect print to log
                import builtins
                _orig = builtins.print
                builtins.print = lambda *a, **kw: on_log(' '.join(str(x) for x in a))
                try:
                    split_xlsx(diff_file, split_out, year)
                finally:
                    builtins.print = _orig
                on_log(f'✓ Αρχεία ανά σχολείο → {split_out}')
                self.after(0, lambda: [
                    self._split_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='✂  Διαχωρισμός ανά Σχολείο'),
                    messagebox.showinfo('Διαχωρισμός', 'Ολοκληρώθηκε!', parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._split_btn.configure(
                        state='normal', bg=C['btn_bg'],
                        text='✂  Διαχωρισμός ανά Σχολείο'),
                    messagebox.showerror('Σφάλμα Διαχωρισμού', m, parent=self)
                ])

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 4: Email ──────────────────────────────────────────────────────────

    def _build_email(self, body):
        C = self._C
        self._email_year = tk.StringVar(value=self._guess_year())
        self._section_lbl(body, 'Σχολικό έτος:')
        self._year_field(body, self._email_year)

        tk.Label(body,
                 text='Το αρχείο "Βασικά Στοιχεία Σχολικών Μονάδων" ανιχνεύεται\n'
                      'αυτόματα από τον φάκελο λήψης (αρχείο 10. ...).',
                 bg=C['bg'], fg=C['desc'],
                 font=('Arial', 8), justify='left').pack(anchor='w', pady=(0, 8))

        # Επιλογή τρόπου αποστολής
        self._email_mode = tk.StringVar(value='schools')
        opt = tk.Frame(body, bg=C['bg'])
        opt.pack(fill='x', pady=(0, 6))
        _from_email = getattr(self._cfg, 'FROM_EMAIL', '') or '...'
        tk.Radiobutton(opt, text='Σχολεία (ένα email ανά σχολείο)',
                       variable=self._email_mode, value='schools',
                       bg=C['bg'], selectcolor=C['sel_bg'],
                       activebackground=C['bg'],
                       font=('Arial', 9)).pack(anchor='w')
        tk.Radiobutton(opt, text=f'Test mode — συνολικό αρχείο στο: {_from_email}',
                       variable=self._email_mode, value='test',
                       bg=C['bg'], selectcolor=C['sel_bg'],
                       activebackground=C['bg'],
                       font=('Arial', 9)).pack(anchor='w')

        self._email_log = self._make_log(body)
        br = tk.Frame(body, bg=C['bg'])
        br.pack(fill='x')
        self._email_btn = self._run_btn(br, '✉  Αποστολή Email', self._start_emails)

    def _start_emails(self):
        C    = self._C
        year = self._email_year.get().strip()
        mode = self._email_mode.get()  # 'schools' ή 'test'
        cfg  = self._cfg

        if not getattr(cfg, 'FROM_PASSWORD', '').strip():
            messagebox.showwarning('Προσοχή',
                'Ο κωδικός email δεν έχει οριστεί.\nΠήγαινε στις Ρυθμίσεις (⚙).',
                parent=self)
            return

        if mode == 'schools':
            # Αποστολή ανά σχολείο — χρειάζεται αρχείο 10 + split φάκελο
            sch_candidates = sorted(glob.glob(os.path.join(self._dl_dir(year), '10.*')))
            sch_path = next(
                (f for f in sch_candidates if not f.endswith(('.tmp', '.crdownload'))), None)
            if not sch_path:
                messagebox.showwarning('Προσοχή',
                    'Δεν βρέθηκε το αρχείο "10. Βασικά Στοιχεία Σχολικών Μονάδων".\n\n'
                    'Τρέξτε πρώτα τη Λήψη (tab ⬇).', parent=self)
                return
            split_out = self._split_dir(year)
            if not os.path.isdir(split_out) or \
               not any(f.endswith('.xlsx') for f in os.listdir(split_out)):
                messagebox.showwarning('Προσοχή',
                    f'Δεν βρέθηκαν αρχεία ανά σχολείο σε:\n{split_out}\n\n'
                    'Τρέξτε πρώτα τον Διαχωρισμό.', parent=self)
                return
            ok = messagebox.askyesno('Επιβεβαίωση Αποστολής',
                'Θα αποσταλούν emails σε ΟΛΑ τα σχολεία που έχουν αποκλίσεις.\n\nΣυνέχεια;',
                parent=self)
            if not ok:
                return
        else:
            # Test mode — στέλνει συνολικό αρχείο στο FROM_EMAIL
            out_dir = self._out_dir(year)
            diff_files = sorted(glob.glob(os.path.join(out_dir, 'differences_*.xlsx')))
            if not diff_files:
                messagebox.showwarning('Προσοχή',
                    f'Δεν βρέθηκε αρχείο αποκλίσεων σε:\n{out_dir}\n\n'
                    'Τρέξτε πρώτα τη Σύγκριση.', parent=self)
                return
            sch_path   = None
            split_out  = None
            diff_file  = diff_files[-1]  # πιο πρόσφατο

        self._email_btn.configure(state='disabled', bg=C['btn_dis'], text='Αποστολή...')

        def on_log(msg):
            self._log_append(self._email_log, msg)

        def task():
            try:
                from smeae.compare import send_emails, send_email_with_attachment
                if mode == 'schools':
                    send_emails(
                        output_dir      = split_out,
                        school_year     = year,
                        email_from      = cfg.FROM_EMAIL,
                        username        = cfg.FROM_EMAIL,
                        password        = cfg.FROM_PASSWORD,
                        smtp_host       = cfg.SMTP_HOST,
                        school_dir_path = sch_path,
                        dry_run         = False,
                        send_only_one   = False,
                        callback        = on_log,
                    )
                else:
                    on_log(f'Test mode — αποστολή συνολικού αρχείου στο {cfg.FROM_EMAIL}')
                    send_email_with_attachment(
                        receiver_email  = cfg.FROM_EMAIL,
                        attachment_path = diff_file,
                        dry_run         = False,
                        sender_email    = cfg.FROM_EMAIL,
                        username        = cfg.FROM_EMAIL,
                        password        = cfg.FROM_PASSWORD,
                        smtp_host       = cfg.SMTP_HOST,
                        first_email     = True,
                        school_name     = 'Test',
                        school_year     = year,
                        callback        = on_log,
                    )
                on_log('\n✓ Αποστολή ολοκληρώθηκε!')
                self.after(0, lambda: [
                    self._email_btn.configure(
                        state='normal', bg=C['btn_bg'], text='✉  Αποστολή Email'),
                    messagebox.showinfo('Email', 'Αποστολή ολοκληρώθηκε!', parent=self)
                ])
            except Exception as e:
                err = str(e)
                self.after(0, lambda m=err: [
                    self._email_btn.configure(
                        state='normal', bg=C['btn_bg'], text='✉  Αποστολή Email'),
                    messagebox.showerror('Σφάλμα Email', m, parent=self)
                ])

        threading.Thread(target=task, daemon=True).start()
