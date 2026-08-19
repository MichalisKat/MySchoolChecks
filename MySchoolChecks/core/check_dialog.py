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
        t3 = tk.Frame(nb, bg=C['bg'], padx=16, pady=12)
        nb.add(t1, text='  ⬇ Λήψη  ')
        nb.add(t2, text='  ▶ Εκτέλεση  ')
        nb.add(t3, text='  ✉ Αποστολή  ')

        self._build_download(t1)
        self._build_execute(t2)
        self._build_send(t3)

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
        self._dl_btn.configure(state='disabled', bg=C['btn_dis'], text='Εκτελείται...')

        def on_log(msg):
            self._log_append(self._dl_log, msg)

        def task():
            try:
                from core.downloader import MySchoolDownloader, get_downloads_dir, cleanup_old_downloads
                dest = get_downloads_dir(self._base)
                dl = MySchoolDownloader(
                    usr, pw, dest, callback=on_log, reports=rids,
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
                    if self._is_custom_run:
                        self._lock_send_tab(True)
                        self._send_hint.configure(
                            text='Αυτός ο έλεγχος διαχειρίζεται μόνος του την αποστολή '
                                 'email μέσα στο tab «Εκτέλεση» (θα σου ζητήσει τις '
                                 'επιλογές εκεί) — το tab «Αποστολή» δεν χρησιμοποιείται.')
                    else:
                        can_send = bool(self._exec_result
                                        and self._exec_result.get('status') == 'ok'
                                        and self._exec_result.get('has_email'))
                        self._lock_send_tab(not can_send)
                self.after(0, _done)

        threading.Thread(target=task, daemon=True).start()

    # ── Tab 3: Αποστολή ──────────────────────────────────────────────────────

    def _build_send(self, body):
        C = self._C
        self._send_mode = tk.StringVar(value='test')
        _from_email = getattr(self._cfg, 'FROM_EMAIL', '') or '...'

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
