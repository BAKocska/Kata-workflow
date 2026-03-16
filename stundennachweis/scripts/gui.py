#!/usr/bin/env python3
"""
Tkinter GUI for the Stundennachweis template generator.

Provides a form-based interface so Kata can generate monthly timesheets
without touching a terminal.
"""

import os
import sys
import math
import calendar
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date

# ---------------------------------------------------------------------------
# Imports from the core module
# ---------------------------------------------------------------------------

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_templates import (
    GERMAN_MONTHS,
    compute_weekdays,
    generate_file,
    read_current_data,
)

GERMAN_DAYS_SHORT = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']

# ---------------------------------------------------------------------------
# Spinning pixelated dachshund
# ---------------------------------------------------------------------------

# Pixel art dachshund (facing left) — B=black, D=dark gray, T=tan, E=eye
_DACHSHUND_COLORS = {
    'B': '#000000',
    'D': '#2e3238',
    'T': '#c8946c',
    'E': '#2a1a0a',
}

_DACHSHUND = [
    '.........BBB..........................',
    '........BDDDBB........................',
    '.......BDDDDBBB.......................',
    '....BBBDBDDBDDBB......................',
    '....BDDDDDDTBDDBB.....................',
    '....BTTTTTDDBDDBB.....................',
    '.....BBBTTTBBBBBBBBBBBBBBBBBBBBB.BB...',
    '........BBBTBBBBDDDDDDDDDDDDDDDDB..B.',
    '...........BDBBDDDDDDDDDDDDDDDDDBB..B',
    '...........BTDDDDDDDDDDDDDDDDDDDDBB.B',
    '............BTDDDDDDDDDDDDDDBDDDDBB.B',
    '............BTDDDDDDDDDDDDDDBDDDDBB.B',
    '............BTDDDDDBDDDDDDTTBTDDDBB..B',
    '.............BTBDDDBBBBTTTTB.BTDBB....',
    '..............BBBTB.BBBBBBB...BTBB....',
    '................BTB............BTB....',
    '...............BTB............BTB.....',
    '...............BB.............BB......',
]


class SpinningDachshund:
    """A pixelated dachshund spinning around a vertical axis."""

    PX = 3

    def __init__(self, canvas):
        self.canvas = canvas
        self.angle = 0.0
        self.rects = []
        self.pixels = []
        for r, row in enumerate(_DACHSHUND):
            for c, ch in enumerate(row):
                if ch in _DACHSHUND_COLORS:
                    self.pixels.append((c, r, _DACHSHUND_COLORS[ch]))
        xs = [p[0] for p in self.pixels]
        self.cx = (min(xs) + max(xs)) / 2
        self._animate()

    def _animate(self):
        px = self.PX
        scale = math.cos(self.angle)
        for r in self.rects:
            self.canvas.delete(r)
        self.rects.clear()
        for x, y, color in self.pixels:
            proj_x = self.cx + (x - self.cx) * scale
            x1 = proj_x * px
            y1 = y * px
            w = px * abs(scale)
            if w < 0.5:
                continue
            if scale >= 0:
                r = self.canvas.create_rectangle(
                    x1, y1, x1 + w, y1 + px, fill=color, outline='')
            else:
                r = self.canvas.create_rectangle(
                    x1 - w, y1, x1, y1 + px, fill=color, outline='')
            self.rects.append(r)
        self.angle += 0.04
        if self.angle >= 2 * math.pi:
            self.angle -= 2 * math.pi
        self.canvas.after(40, self._animate)


def _default_template_path():
    """Locate empty_template.xlsx: next to .exe (frozen) or in data/input/."""
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
    candidate = os.path.join(base, 'data', 'input', 'empty_template.xlsx')
    if os.path.exists(candidate):
        return candidate
    # Fallback: same directory as the executable / script
    candidate = os.path.join(
        os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
        else os.path.dirname(os.path.abspath(__file__)),
        'empty_template.xlsx',
    )
    return candidate if os.path.exists(candidate) else ''


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('Stundennachweis Generator')
        self.root.resizable(False, False)

        self.assignments = []
        self.contact_entries = {}
        self.holiday_vars = {}
        self.progress_queue = queue.Queue()

        # --- Tkinter variables ---
        self.var_data_path = tk.StringVar()
        self.var_template_path = tk.StringVar(value=_default_template_path())
        self.var_output_dir = tk.StringVar()
        self.var_year = tk.IntVar(value=date.today().year)
        self.var_month = tk.IntVar(value=date.today().month)

        pad = dict(padx=8, pady=4)

        # --- Spinning dachshund ---
        art_w = (max(len(row) for row in _DACHSHUND)) * 4
        art_h = len(_DACHSHUND) * 4
        dachshund_canvas = tk.Canvas(root, width=art_w, height=art_h,
                                     highlightthickness=0)
        dachshund_canvas.pack(pady=(8, 0))
        SpinningDachshund(dachshund_canvas)

        # --- Section: Files ---
        frm_files = ttk.LabelFrame(root, text='Files')
        frm_files.pack(fill='x', **pad)

        ttk.Label(frm_files, text='current_data.xlsx:').grid(
            row=0, column=0, sticky='w', padx=4, pady=2)
        ttk.Entry(frm_files, textvariable=self.var_data_path, width=52).grid(
            row=0, column=1, padx=4, pady=2)
        ttk.Button(frm_files, text='Browse', width=8,
                   command=self._browse_data).grid(row=0, column=2, padx=4)
        self.btn_load = ttk.Button(frm_files, text='Load Data', width=10,
                                   command=self._load_data)
        self.btn_load.grid(row=0, column=3, padx=4)

        ttk.Label(frm_files, text='Output folder:').grid(
            row=1, column=0, sticky='w', padx=4, pady=2)
        ttk.Entry(frm_files, textvariable=self.var_output_dir, width=52).grid(
            row=1, column=1, padx=4, pady=2)
        ttk.Button(frm_files, text='Browse', width=8,
                   command=self._browse_output).grid(row=1, column=2, padx=4)

        ttk.Label(frm_files, text='empty_template.xlsx:').grid(
            row=2, column=0, sticky='w', padx=4, pady=2)
        ttk.Entry(frm_files, textvariable=self.var_template_path, width=52).grid(
            row=2, column=1, padx=4, pady=2)
        ttk.Button(frm_files, text='Browse', width=8,
                   command=self._browse_template).grid(row=2, column=2, padx=4)

        # --- Section: Period ---
        frm_period = ttk.LabelFrame(root, text='Period')
        frm_period.pack(fill='x', **pad)

        ttk.Label(frm_period, text='Year:').pack(side='left', padx=(8, 2))
        self.spn_year = ttk.Spinbox(
            frm_period, from_=2020, to=2100, width=6,
            textvariable=self.var_year, command=self._on_period_change)
        self.spn_year.pack(side='left', padx=4)

        ttk.Label(frm_period, text='Month:').pack(side='left', padx=(16, 2))
        month_names = [GERMAN_MONTHS[m] for m in range(1, 13)]
        self.cmb_month = ttk.Combobox(
            frm_period, values=month_names, width=12, state='readonly')
        self.cmb_month.current(self.var_month.get() - 1)
        self.cmb_month.pack(side='left', padx=4)
        self.cmb_month.bind('<<ComboboxSelected>>', self._on_period_change)

        # --- Section: Holidays ---
        frm_holidays = ttk.LabelFrame(root, text='National Holidays')
        frm_holidays.pack(fill='x', **pad)
        self.frm_holiday_grid = ttk.Frame(frm_holidays)
        self.frm_holiday_grid.pack(fill='x', padx=4, pady=4)
        self._rebuild_holiday_grid()

        # --- Section: Contacts ---
        frm_contacts = ttk.LabelFrame(root, text='Auftragsabwicklung mit')
        frm_contacts.pack(fill='both', expand=True, **pad)

        self.contact_canvas = tk.Canvas(frm_contacts, height=180,
                                        highlightthickness=0)
        self.contact_scrollbar = ttk.Scrollbar(
            frm_contacts, orient='vertical', command=self.contact_canvas.yview)
        self.frm_contacts_inner = ttk.Frame(self.contact_canvas)
        self.frm_contacts_inner.bind(
            '<Configure>',
            lambda e: self.contact_canvas.configure(
                scrollregion=self.contact_canvas.bbox('all')))
        self.contact_canvas.create_window(
            (0, 0), window=self.frm_contacts_inner, anchor='nw')
        self.contact_canvas.configure(yscrollcommand=self.contact_scrollbar.set)

        self.contact_canvas.pack(side='left', fill='both', expand=True)
        self.contact_scrollbar.pack(side='right', fill='y')

        self.lbl_contacts_hint = ttk.Label(
            self.frm_contacts_inner,
            text='  Load current_data.xlsx first.',
            foreground='gray')
        self.lbl_contacts_hint.pack(anchor='w', pady=8)

        # --- Section: Generate ---
        frm_gen = ttk.Frame(root)
        frm_gen.pack(fill='x', **pad)

        self.btn_generate = ttk.Button(
            frm_gen, text='Generate', command=self._on_generate)
        self.btn_generate.pack(pady=(0, 4))

        self.progress_bar = ttk.Progressbar(frm_gen, length=400, mode='determinate')
        self.progress_bar.pack(fill='x', padx=4)

        self.lbl_status = ttk.Label(frm_gen, text='Ready.', anchor='w')
        self.lbl_status.pack(fill='x', padx=4, pady=(2, 4))

    # ------------------------------------------------------------------
    # File browsers
    # ------------------------------------------------------------------

    def _browse_data(self):
        p = filedialog.askopenfilename(
            title='Select current_data.xlsx',
            filetypes=[('Excel files', '*.xlsx')])
        if p:
            self.var_data_path.set(p)

    def _browse_output(self):
        p = filedialog.askdirectory(title='Select output folder')
        if p:
            self.var_output_dir.set(p)

    def _browse_template(self):
        p = filedialog.askopenfilename(
            title='Select empty_template.xlsx',
            filetypes=[('Excel files', '*.xlsx')])
        if p:
            self.var_template_path.set(p)

    # ------------------------------------------------------------------
    # Load data
    # ------------------------------------------------------------------

    def _load_data(self):
        path = self.var_data_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror('Error', 'Please select a valid current_data.xlsx.')
            return
        try:
            self.assignments = read_current_data(data_path=path)
        except Exception as exc:
            messagebox.showerror('Error', f'Failed to read data file:\n{exc}')
            return

        # Populate contacts
        for w in self.frm_contacts_inner.winfo_children():
            w.destroy()
        self.contact_entries.clear()

        projects = sorted(set(a['project_name'] for a in self.assignments))
        for i, proj in enumerate(projects):
            ttk.Label(self.frm_contacts_inner, text=proj).grid(
                row=i, column=0, sticky='w', padx=(4, 8), pady=2)
            var = tk.StringVar()
            ttk.Entry(self.frm_contacts_inner, textvariable=var, width=30).grid(
                row=i, column=1, sticky='w', padx=4, pady=2)
            self.contact_entries[proj] = var

        self.contact_canvas.update_idletasks()
        self.contact_canvas.configure(
            scrollregion=self.contact_canvas.bbox('all'))

        self.lbl_status.config(
            text=f'Loaded {len(self.assignments)} rows, '
                 f'{len(projects)} projects.')

    # ------------------------------------------------------------------
    # Period / holiday grid
    # ------------------------------------------------------------------

    def _on_period_change(self, _event=None):
        try:
            self.var_year.set(int(self.spn_year.get()))
        except ValueError:
            pass
        self.var_month.set(self.cmb_month.current() + 1)
        self._rebuild_holiday_grid()

    def _rebuild_holiday_grid(self):
        for w in self.frm_holiday_grid.winfo_children():
            w.destroy()
        self.holiday_vars.clear()

        year = self.var_year.get()
        month = self.var_month.get()
        num_days = calendar.monthrange(year, month)[1]

        cols = 7
        for d in range(1, num_days + 1):
            wd = date(year, month, d).weekday()  # 0=Mon
            var = tk.BooleanVar(value=False)
            self.holiday_vars[d] = var
            label = f'{d:2d}. {GERMAN_DAYS_SHORT[wd]}'
            cb = ttk.Checkbutton(
                self.frm_holiday_grid, text=label, variable=var, width=7)
            r, c = divmod(d - 1, cols)
            cb.grid(row=r, column=c, sticky='w', padx=2, pady=1)
            if wd >= 5:  # Sat/Sun
                cb.config(state='disabled')

    # ------------------------------------------------------------------
    # Generate
    # ------------------------------------------------------------------

    def _on_generate(self):
        # --- validate ---
        tpl = self.var_template_path.get().strip()
        if not tpl or not os.path.isfile(tpl):
            messagebox.showerror('Error', 'Template file not found.')
            return
        if not self.assignments:
            messagebox.showerror('Error', 'No data loaded. Click "Load Data" first.')
            return
        out_dir = self.var_output_dir.get().strip()
        if not out_dir:
            messagebox.showerror('Error', 'Select an output folder.')
            return

        contacts = {proj: var.get().strip()
                    for proj, var in self.contact_entries.items()}
        empty = [p for p, v in contacts.items() if not v]
        if empty:
            if not messagebox.askokcancel(
                    'Warning',
                    f'{len(empty)} project(s) have no contact person.\n'
                    'Continue anyway?'):
                return

        year = self.var_year.get()
        month = self.var_month.get()
        holidays = [d for d, var in self.holiday_vars.items() if var.get()]
        weekdays = compute_weekdays(year, month, holidays)

        params = dict(
            template_path=tpl,
            assignments=self.assignments,
            year=year,
            month=month,
            weekdays=weekdays,
            contacts=contacts,
            output_dir=out_dir,
        )

        self.btn_generate.config(state='disabled')
        self.progress_bar['value'] = 0
        thread = threading.Thread(
            target=self._generate_worker, args=(params,), daemon=True)
        thread.start()
        self.root.after(100, self._poll_progress)

    def _generate_worker(self, p):
        try:
            os.makedirs(p['output_dir'], exist_ok=True)
            with open(p['template_path'], 'rb') as f:
                template_bytes = f.read()

            total = len(p['assignments'])
            seen = {}
            for i, a in enumerate(p['assignments'], 1):
                safe_proj = a['project_name'].replace('/', '_').replace('\\', '_')
                proj_dir = os.path.join(p['output_dir'], safe_proj)
                os.makedirs(proj_dir, exist_ok=True)
                safe = a['resource'].replace('/', '_').replace('\\', '_')
                base = f"{safe}_{a['purchase_order']}_{p['year']}_{p['month']:02d}"
                seen[base] = seen.get(base, 0) + 1
                fn = (f'{base}.xlsx' if seen[base] == 1
                      else f'{base}_{seen[base]}.xlsx')
                contact = p['contacts'].get(a['project_name'], '')
                generate_file(
                    template_bytes, a, p['year'], p['month'],
                    p['weekdays'], contact,
                    os.path.join(proj_dir, fn))
                self.progress_queue.put(('progress', i, total, f'{safe_proj}/{fn}'))

            self.progress_queue.put(('done', total, p['output_dir']))
        except Exception as exc:
            self.progress_queue.put(('error', str(exc)))

    def _poll_progress(self):
        while not self.progress_queue.empty():
            msg = self.progress_queue.get_nowait()
            if msg[0] == 'progress':
                _, i, total, fn = msg
                self.progress_bar['value'] = (i / total) * 100
                self.lbl_status.config(text=f'Generating... {i}/{total}: {fn}')
            elif msg[0] == 'done':
                _, total, out_dir = msg
                self.lbl_status.config(text=f'Done! {total} files in {out_dir}')
                self.btn_generate.config(state='normal')
                messagebox.showinfo('Done', f'Generated {total} files.')
                return
            elif msg[0] == 'error':
                self.lbl_status.config(text=f'Error: {msg[1]}')
                self.btn_generate.config(state='normal')
                messagebox.showerror('Error', msg[1])
                return
        self.root.after(100, self._poll_progress)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
