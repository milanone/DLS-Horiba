"""
NSZ Viewer — Horiba SZ-100 DLS file browser & plotter
Supporta drag & drop, confronto multiplo, esportazione Excel.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, struct, re, csv, io
try:
    import olefile
    OLEFILE_AVAILABLE = True
except ImportError:
    OLEFILE_AVAILABLE = False
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import numpy as np

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ── Palette colori per campioni multipli ──────────────────────────────────────
COLORS = ["#2563EB","#DC2626","#16A34A","#D97706","#7C3AED",
          "#0891B2","#BE185D","#65A30D","#EA580C","#4338CA","#0D9488"]

# ── Parser NSZ ────────────────────────────────────────────────────────────────

def f32(data):
    n = len(data)//4
    return list(struct.unpack(f'<{n}f', data[:n*4]))

def parse_props(data):
    props = {}
    i = 0
    while i < len(data) - 10:
        j = data.find(b'\x00', i)
        if j == -1 or j - i < 2:
            i += 1; continue
        key_b = data[i:j]
        if all(32 <= b < 127 for b in key_b) and len(key_b) >= 3:
            if j + 5 < len(data) and data[j+5] == 0x00:
                try:
                    v = struct.unpack('<f', data[j+1:j+5])[0]
                    if not (v != v):
                        props[key_b.decode('ascii')] = v
                except (struct.error, UnicodeDecodeError): pass
            i = j + 1
        else:
            i += 1
    return props

def _open_ole(path):
    if not OLEFILE_AVAILABLE:
        raise ImportError("Libreria olefile non disponibile.\nInstallala con: pip install olefile")
    return olefile.OleFileIO(path)

def dump_nsz(path):
    """Stampa struttura interna del file .nsz per debug (time base, parametri correlatore)."""
    ole = _open_ole(path)
    try:
        def s(name):
            return ole.openstream(name).read() if ole.exists(name) else b''

        print(f"\n=== {os.path.basename(path)} ===")
        print("Object3 header (primi 16 float32):")
        hdr = f32(s('Object3'))[:16]
        for i, v in enumerate(hdr):
            print(f"  [{i:2d}] {v:.6g}")

        print("\nBase34 tutti i parametri float:")
        p34 = parse_props(s('Base34'))
        for k, v in sorted(p34.items()):
            print(f"  {k} = {v:.6g}")
    finally:
        ole.close()

def load_nsz(path):
    """Carica un file .nsz e restituisce un dict con tutti i dati."""
    ole = _open_ole(path)
    try:
        return _read_nsz(path, ole)
    finally:
        ole.close()

def _read_nsz(path, ole):
    def s(name):
        return ole.openstream(name).read() if ole.exists(name) else b''

    # ACF — tronca i canali baseline ripetuti alla fine (multi-tau correlator)
    tau_full = f32(s('Object3'))[16:]
    acf_full = f32(s('Object2'))[3:]
    cut = next((i for i in range(1, len(tau_full)) if tau_full[i] <= tau_full[i-1]),
               len(tau_full))
    tau = tau_full[:cut]
    acf = acf_full[:cut]

    # Fit e residui Horiba.
    # Object10 header: [int=1][int=n_fit][int=fit_offset], poi i valori del fit.
    # fit_offset = numero di canali iniziali (afterpulsing) esclusi dall'analisi.
    obj10_raw = f32(s('Object10'))
    FIT_OFFSET = struct.unpack('<I', struct.pack('<f', obj10_raw[2]))[0] if len(obj10_raw) > 2 else 6
    fit   = obj10_raw[3:]
    resid = f32(s('Object7'))[3:]

    # Size distribution
    sz    = f32(s('Object5'))[16:100]   # 84 nm bins
    intens= f32(s('Object4'))[2:2+84]   # header=2 float32 (validato vs CSV software)
    cumul = f32(s('Object31'))[2:2+84]
    nb    = min(len(sz), len(intens), len(cumul))
    sz, intens, cumul = sz[:nb], intens[:nb], cumul[:nb]

    # Parametri
    p1  = parse_props(s('Base1'))
    p34 = parse_props(s('Base34'))
    d34 = s('Base34')
    d38 = s('Base38')

    def get_str(data, key):
        m = re.search(key + b'\x00([\\x20-\\x7e]+)', data)
        return m.group(1).decode().strip() if m else ''

    m = re.search(rb'MeasID\x00{0,4}(\d{15})', d34)
    raw_id = m.group(1).decode() if m else '000000000000000'

    def pv(d, k, scale=1):
        v = d.get(k)
        if v is not None and abs(v) < 1e6:
            return round(v * scale, 4)
        return None

    params = {
        'Sample_Name': get_str(d38, b'Sample_Name') or os.path.splitext(os.path.basename(path))[0],
        'Date': f'{raw_id[0:4]}-{raw_id[4:6]}-{raw_id[6:8]}',
        'Time': f'{raw_id[8:10]}:{raw_id[10:12]}:{raw_id[12:14]}',
        'Solvent': get_str(d34, b'MeasSolventName'),
        'Software': get_str(d34, b'AppVersion'),
        'Angle_deg': pv(p34, 'MeasAngle'),
        'Lambda_nm': pv(p34, 'MeasWaveLength'),
        'Temp_C':   pv(p34, 'MeasHolderTemp'),
        'Visc_mPas':pv(p34, 'MeasSolventVisco'),
        'n_solvent': pv(p34, 'MeasSolventRef'),
        # I campi CalcMean/SD/Mode/PeakPos/D10/D50/D90 sono in RAGGIO → ×2
        # PDI: CalcTotalPI è errato, usare Cum_fPi
        'Peak1_nm': pv(p1, 'CalcPeakPos[0]', 2.0),
        'Mean1_nm': pv(p1, 'CalcMean[0]',    2.0),
        'SD1_nm':   pv(p1, 'CalcSD[0]',      2.0),
        'Area1_pct':pv(p1, 'CalcArea[0]'),
        'Peak2_nm': pv(p1, 'CalcPeakPos[1]', 2.0),
        'Mean2_nm': pv(p1, 'CalcMean[1]',    2.0),
        'SD2_nm':   pv(p1, 'CalcSD[1]',      2.0),
        'D10_nm':   pv(p1, 'CalcPerDiameter[0]', 2.0),
        'D50_nm':   pv(p1, 'CalcPerDiameter[4]', 2.0),
        'D90_nm':   pv(p1, 'CalcPerDiameter[8]', 2.0),
        'Mode_nm':  pv(p1, 'CalcMode',   2.0),
        'Median_nm':pv(p1, 'CalcMedian', 2.0),
        'ZAvg_nm':  pv(p1, 'Cum_fMean',  2.0),
        'PDI':      pv(p1, 'Cum_fPi'),          # Cum_fPi = PDI corretto
        'Span':     pv(p1, 'CalcSpan'),
        'GeoMean_nm': pv(p1, 'CalcGeoMean', 2.0),
        'AriMean_nm': pv(p1, 'CalcAriMean', 2.0),
    }

    return {
        'path': path,
        'name': os.path.splitext(os.path.basename(path))[0],
        'tau':        np.array(tau,   dtype=float),
        'acf':        np.array(acf,   dtype=float),
        'fit':        np.array(fit,   dtype=float),
        'resid':      np.array(resid, dtype=float),
        'fit_offset': FIT_OFFSET,
        'sz':         np.array(sz,    dtype=float),
        'intens':     np.array(intens,dtype=float),
        'cumul':      np.array(cumul, dtype=float),
        'params':     params,
    }

# ── App principale ────────────────────────────────────────────────────────────

class NSZViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("NSZ Viewer — Horiba SZ-100")
        self.root.geometry("1380x840")
        self.root.configure(bg="#F1F5F9")

        self.samples = {}   # path → data dict
        self.selected = []  # list of paths currently shown

        self._build_ui()
        self._setup_dnd()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Sidebar sinistra ──────────────────────────────────────────────────
        sidebar = tk.Frame(self.root, bg="#1E293B", width=260)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=0, pady=0)
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="NSZ Viewer", font=("Helvetica", 15, "bold"),
                 fg="white", bg="#1E293B").pack(pady=(18,4))
        tk.Label(sidebar, text="Horiba SZ-100 DLS", font=("Helvetica", 9),
                 fg="#94A3B8", bg="#1E293B").pack(pady=(0,14))

        # Pulsanti
        btn_frame = tk.Frame(sidebar, bg="#1E293B")
        btn_frame.pack(fill=tk.X, padx=10, pady=(0,10))

        for text, cmd in [
            ("＋  Apri file .nsz", self._open_files),
            ("📁  Apri cartella",  self._open_folder),
            ("🗑  Rimuovi sel.",   self._remove_selected),
            ("💾  Esporta Excel",  self._export_xlsx),
        ]:
            b = tk.Button(btn_frame, text=text, command=cmd,
                          bg="#334155", fg="white", relief=tk.FLAT,
                          font=("Helvetica", 10), anchor="w",
                          padx=10, pady=6, cursor="hand2",
                          activebackground="#475569", activeforeground="white")
            b.pack(fill=tk.X, pady=2)

        # Zona drag & drop
        drop_lbl = tk.Label(sidebar,
            text="⬇ Trascina qui i file .nsz",
            font=("Helvetica", 9), fg="#64748B", bg="#1E293B",
            pady=10)
        drop_lbl.pack(fill=tk.X, padx=10)

        ttk.Separator(sidebar, orient="horizontal").pack(fill=tk.X, padx=10, pady=8)

        # Lista campioni
        tk.Label(sidebar, text="CAMPIONI CARICATI", font=("Helvetica", 8, "bold"),
                 fg="#64748B", bg="#1E293B").pack(anchor="w", padx=12)

        list_frame = tk.Frame(sidebar, bg="#1E293B")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED,
                                  bg="#0F172A", fg="white",
                                  selectbackground="#2563EB",
                                  font=("Helvetica", 9),
                                  borderwidth=0, highlightthickness=0,
                                  activestyle="none")
        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                               command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scroll.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # ── Area principale ───────────────────────────────────────────────────
        main = tk.Frame(self.root, bg="#F1F5F9")
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Tabs
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background="#F1F5F9", borderwidth=0)
        style.configure("TNotebook.Tab", padding=[14,7], font=("Helvetica", 10))
        style.map("TNotebook.Tab",
                  background=[("selected","white"),("!selected","#E2E8F0")],
                  foreground=[("selected","#1E293B"),("!selected","#64748B")])

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: ACF
        self.tab_acf = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_acf, text="  Autocorrelogramma  ")

        # Tab 2: Distribuzione dimensionale
        self.tab_size = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_size, text="  Distribuzione  ")

        # Tab 3: Parametri
        self.tab_params = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_params, text="  Parametri  ")

        self._build_acf_tab()
        self._build_size_tab()
        self._build_params_tab()

    def _build_acf_tab(self):
        ctrl = tk.Frame(self.tab_acf, bg="white", pady=4)
        ctrl.pack(fill=tk.X, padx=14)

        tk.Label(ctrl, text="Scala X:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        self.acf_xscale = tk.StringVar(value="log")
        for val, lbl in [("log","Log"),("linear","Lineare")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.acf_xscale, value=val,
                           bg="white", font=("Helvetica",9),
                           command=self._plot_acf).pack(side=tk.LEFT, padx=4)

        self.show_fit = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text="Mostra fit ideale", variable=self.show_fit,
                       bg="white", font=("Helvetica",9),
                       command=self._plot_acf).pack(side=tk.LEFT, padx=12)

        self.show_resid = tk.BooleanVar(value=False)
        tk.Checkbutton(ctrl, text="Mostra residui", variable=self.show_resid,
                       bg="white", font=("Helvetica",9),
                       command=self._plot_acf).pack(side=tk.LEFT, padx=4)

        self.fig_acf = Figure(figsize=(9,5.5), dpi=100, facecolor="white")
        self.canvas_acf = FigureCanvasTkAgg(self.fig_acf, self.tab_acf)
        NavigationToolbar2Tk(self.canvas_acf, self.tab_acf).pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_acf.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

    def _build_size_tab(self):
        ctrl = tk.Frame(self.tab_size, bg="white", pady=4)
        ctrl.pack(fill=tk.X, padx=14)

        tk.Label(ctrl, text="Scala X:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        self.size_xscale = tk.StringVar(value="log")
        for val, lbl in [("log","Log"),("linear","Lineare")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.size_xscale, value=val,
                           bg="white", font=("Helvetica",9),
                           command=self._plot_size).pack(side=tk.LEFT, padx=4)

        ttk.Separator(ctrl, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(ctrl, text="Rappresentazione:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        self.size_repr = tk.StringVar(value="intensity")
        for val, lbl in [("intensity","Intensità"),("number","Numero")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.size_repr, value=val,
                           bg="white", font=("Helvetica",9),
                           command=self._plot_size).pack(side=tk.LEFT, padx=4)

        ttk.Separator(ctrl, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        self.show_cumul = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text="Mostra cumulativa", variable=self.show_cumul,
                       bg="white", font=("Helvetica",9),
                       command=self._plot_size).pack(side=tk.LEFT, padx=4)

        self.fig_size = Figure(figsize=(9,5.5), dpi=100, facecolor="white")
        self.canvas_size = FigureCanvasTkAgg(self.fig_size, self.tab_size)
        NavigationToolbar2Tk(self.canvas_size, self.tab_size).pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_size.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

    def _build_params_tab(self):
        frame = tk.Frame(self.tab_params, bg="white")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        cols = ("Parametro",
                "Campione 1","Campione 2","Campione 3",
                "Campione 4","Campione 5")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.heading("Parametro", text="Parametro")
        self.tree.column("Parametro", width=160, anchor="w")
        for c in cols[1:]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=180, anchor="center")

        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica",9), rowheight=22)
        style.configure("Treeview.Heading", font=("Helvetica",9,"bold"))
        style.map("Treeview",
                  background=[("selected","#DBEAFE")],
                  foreground=[("selected","#1E3A5F")])

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Alternanza righe
        self.tree.tag_configure("odd",  background="#F8FAFC")
        self.tree.tag_configure("even", background="white")

    # ── Drag & Drop ───────────────────────────────────────────────────────────

    def _setup_dnd(self):
        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._on_drop)
        # fallback: click sulla label
        self.listbox.bind("<Button-3>", lambda e: self._open_files())

    def _on_drop(self, event):
        raw = event.data
        # tkinterdnd2 restituisce percorsi separati da spazi, con {} per quelli con spazi
        paths = self.root.tk.splitlist(raw)
        nsz_files = []
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith('.nsz'):
                nsz_files.append(p)
            elif os.path.isdir(p):
                for f in os.listdir(p):
                    if f.lower().endswith('.nsz'):
                        nsz_files.append(os.path.join(p, f))
        self._load_files(nsz_files)

    # ── File loading ──────────────────────────────────────────────────────────

    def _open_files(self):
        files = filedialog.askopenfilenames(
            title="Apri file NSZ",
            filetypes=[("Horiba NSZ", "*.nsz"), ("Tutti i file", "*.*")])
        self._load_files(files)

    def _open_folder(self):
        folder = filedialog.askdirectory(title="Apri cartella con file NSZ")
        if folder:
            files = [os.path.join(folder, f)
                     for f in sorted(os.listdir(folder))
                     if f.lower().endswith('.nsz')]
            self._load_files(files)

    def _load_files(self, paths):
        new = 0
        errors = []
        for path in paths:
            path = os.path.normpath(path)
            if path in self.samples:
                continue
            try:
                data = load_nsz(path)
                self.samples[path] = data
                short = os.path.basename(path)
                self.listbox.insert(tk.END, short)
                new += 1
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")

        if errors:
            messagebox.showerror("Errore caricamento",
                                 "\n".join(errors[:5]))
        if new:
            # Seleziona automaticamente i nuovi
            n = self.listbox.size()
            self.listbox.selection_clear(0, tk.END)
            for i in range(n - new, n):
                self.listbox.selection_set(i)
            self._on_select()

    def _remove_selected(self):
        sel = list(self.listbox.curselection())
        paths = list(self.samples.keys())
        for i in reversed(sel):
            path = paths[i]
            del self.samples[path]
            self.listbox.delete(i)
        self.selected = []
        self._refresh_all()

    # ── Selezione ─────────────────────────────────────────────────────────────

    def _on_select(self, event=None):
        indices = self.listbox.curselection()
        paths = list(self.samples.keys())
        self.selected = [paths[i] for i in indices if i < len(paths)]
        self._refresh_all()

    def _refresh_all(self):
        self._plot_acf()
        self._plot_size()
        self._update_params_table()

    # ── Plot ACF ──────────────────────────────────────────────────────────────

    def _plot_acf(self):
        self.fig_acf.clear()

        if self.show_resid.get() and self.selected:
            gs = self.fig_acf.add_gridspec(2, 1, height_ratios=[3,1], hspace=0.08)
            ax  = self.fig_acf.add_subplot(gs[0])
            axr = self.fig_acf.add_subplot(gs[1], sharex=ax)
        else:
            ax  = self.fig_acf.add_subplot(111)
            axr = None

        xscale = self.acf_xscale.get()

        for idx, path in enumerate(self.selected[:len(COLORS)]):
            d = self.samples[path]
            c = COLORS[idx % len(COLORS)]
            lbl = self._short_label(d['name'])

            tau = d['tau']
            acf = d['acf']

            # Filtra valori sensati
            mask = (tau > 0) & (acf >= -0.05) & (acf <= 1.05)
            if mask.sum() < 2:
                continue

            ax.plot(tau[mask], acf[mask], 'o', color=c, markersize=3.5,
                    linewidth=0, label=lbl)

            # Fit ideale — allineato a tau[fit_offset:]
            off = d.get('fit_offset', 6)
            if self.show_fit.get() and len(d['fit']) > 2:
                tau_fit = tau[off:off + len(d['fit'])]
                fit_vals = d['fit'][:len(tau_fit)]
                ax.plot(tau_fit, fit_vals, '-', color=c, linewidth=1.5,
                        alpha=0.85)

            # Residui — stesso offset del fit
            if axr is not None and len(d['resid']) > 2:
                tau_res = tau[off:off + len(d['resid'])]
                res_vals = d['resid'][:len(tau_res)]
                axr.plot(tau_res, res_vals, 'o', color=c, markersize=2.5,
                         linewidth=1, alpha=0.8)

        ax.set_xscale(xscale)
        ax.set_xlabel("Tempo di ritardo τ (µs)", fontsize=10)
        ax.set_ylabel("g₁(τ) normalizzata", fontsize=10)
        ax.set_ylim(-0.05, 1.1)
        ax.set_title("Funzione di autocorrelazione (ACF)", fontsize=11, fontweight='bold')
        ax.grid(True, which='both', alpha=0.3, linestyle='--')
        ax.axhline(0, color='gray', linewidth=0.7, linestyle='--')
        if self.selected:
            ax.legend(fontsize=8, loc='upper right', framealpha=0.9)

        if axr is not None:
            axr.set_xscale(xscale)
            axr.set_xlabel("Tempo di ritardo τ (µs)", fontsize=10)
            axr.set_ylabel("Residui", fontsize=9)
            axr.axhline(0, color='gray', linewidth=0.8)
            axr.grid(True, which='both', alpha=0.3, linestyle='--')
            plt.setp(ax.get_xticklabels(), visible=False)
            ax.set_xlabel("")

        if not self.selected:
            ax.text(0.5, 0.5, "Nessun campione selezionato\n\nCarica file .nsz e selezionali dalla lista",
                    ha='center', va='center', transform=ax.transAxes,
                    fontsize=12, color='#94A3B8')

        self.fig_acf.tight_layout()
        self.canvas_acf.draw()

    # ── Plot distribuzione ────────────────────────────────────────────────────

    @staticmethod
    def _to_number_dist(sz, intens):
        """Converti distribuzione in intensità → numero (approssimazione Rayleigh: N ∝ I/d⁶)."""
        d6 = np.where(sz > 0, sz ** 6, np.nan)
        num = np.where(sz > 0, intens / d6, 0.0)
        total = num.sum()
        if total > 0:
            num = num / total * 100.0
        cumul = np.cumsum(num)
        if cumul[-1] > 0:
            cumul = cumul / cumul[-1] * 100.0
        return num, cumul

    def _refresh_size(self):
        self._plot_size()
        self._update_params_table()

    def _plot_size(self):
        self.fig_size.clear()

        by_number = self.size_repr.get() == "number"
        show_c = self.show_cumul.get()
        if show_c and self.selected:
            gs = self.fig_size.add_gridspec(1, 2, wspace=0.3)
            ax  = self.fig_size.add_subplot(gs[0])
            axc = self.fig_size.add_subplot(gs[1])
        else:
            ax  = self.fig_size.add_subplot(111)
            axc = None

        xscale = self.size_xscale.get()

        for idx, path in enumerate(self.selected[:len(COLORS)]):
            d = self.samples[path]
            c = COLORS[idx % len(COLORS)]
            lbl = self._short_label(d['name'])

            sz    = d['sz']
            mask  = sz > 0
            if mask.sum() < 2:
                continue

            if by_number:
                freq, cumul = self._to_number_dist(sz, d['intens'])
            else:
                freq  = d['intens']
                cumul = d['cumul']

            ax.plot(sz[mask], freq[mask], 'o-', color=c, markersize=3,
                    linewidth=1.5, label=lbl)
            ax.fill_between(sz[mask], freq[mask], alpha=0.08, color=c)

            p = d['params']
            if not by_number and p.get('D50_nm'):
                ax.axvline(p['D50_nm'], color=c, linewidth=0.8,
                           linestyle=':', alpha=0.7)

            if axc is not None:
                axc.plot(sz[mask], cumul[mask], 'o-', color=c, markersize=3,
                         linewidth=1.5, label=lbl)
                if not by_number:
                    for pct in [10, 50, 90]:
                        val = p.get(f'D{pct}_nm')
                        if val:
                            axc.axvline(val, color=c, linewidth=0.7,
                                        linestyle='--', alpha=0.6)

        repr_lbl = "numero" if by_number else "intensità"
        ax.set_xscale(xscale)
        ax.set_xlabel("Diametro (nm)", fontsize=10)
        ax.set_ylabel(f"{'Numero' if by_number else 'Intensità'} (%)", fontsize=10)
        ax.set_title(f"Distribuzione dimensionale in {repr_lbl}", fontsize=11, fontweight='bold')
        ax.grid(True, which='both', alpha=0.3, linestyle='--')
        if self.selected:
            ax.legend(fontsize=8, loc='upper right', framealpha=0.9)

        if axc is not None:
            axc.set_xscale(xscale)
            axc.set_xlabel("Diametro (nm)", fontsize=10)
            axc.set_ylabel("Cumulativa (%)", fontsize=10)
            axc.set_title(f"Distribuzione cumulativa in {repr_lbl}", fontsize=11, fontweight='bold')
            axc.set_ylim(-2, 104)
            axc.axhline(10, color='gray', linewidth=0.5, linestyle=':')
            axc.axhline(50, color='gray', linewidth=0.5, linestyle=':')
            axc.axhline(90, color='gray', linewidth=0.5, linestyle=':')
            axc.grid(True, which='both', alpha=0.3, linestyle='--')
            if self.selected:
                axc.legend(fontsize=8, loc='upper left', framealpha=0.9)

        if not self.selected:
            ax.text(0.5, 0.5, "Nessun campione selezionato\n\nCarica file .nsz e selezionali dalla lista",
                    ha='center', va='center', transform=ax.transAxes,
                    fontsize=12, color='#94A3B8')

        self.fig_size.tight_layout()
        self.canvas_size.draw()

    # ── Tabella parametri ─────────────────────────────────────────────────────

    def _update_params_table(self):
        self.tree.delete(*self.tree.get_children())
        sel = self.selected[:5]
        if not sel:
            return

        # Aggiorna intestazioni colonne
        for i, path in enumerate(sel):
            n = self.samples[path]['params'].get('Sample_Name','')
            lbl = n
            self.tree.heading(f"Campione {i+1}", text=lbl)

        ROWS = [
            ("── Misurazione ──", None),
            ("Data / Ora", lambda p: f"{p.get('Date','')} {p.get('Time','')}"),
            ("Angolo (°)", lambda p: p.get('Angle_deg')),
            ("Lunghezza d'onda (nm)", lambda p: p.get('Lambda_nm')),
            ("Temperatura (°C)", lambda p: p.get('Temp_C')),
            ("Viscosità (mPas)", lambda p: p.get('Visc_mPas')),
            ("n solvente", lambda p: p.get('n_solvent')),
            ("Solvente", lambda p: p.get('Solvent')),
            ("── Risultati ──", None),
            ("Peak 1 (nm)", lambda p: p.get('Peak1_nm')),
            ("Mean 1 (nm)", lambda p: p.get('Mean1_nm')),
            ("SD 1 (nm)", lambda p: p.get('SD1_nm')),
            ("Area 1 (%)", lambda p: p.get('Area1_pct')),
            ("Peak 2 (nm)", lambda p: p.get('Peak2_nm')),
            ("Mean 2 (nm)", lambda p: p.get('Mean2_nm')),
            ("D10 (nm)", lambda p: p.get('D10_nm')),
            ("D50 (nm)", lambda p: p.get('D50_nm')),
            ("D90 (nm)", lambda p: p.get('D90_nm')),
            ("Mode (nm)", lambda p: p.get('Mode_nm')),
            ("Median (nm)", lambda p: p.get('Median_nm')),
            ("Z-Average (nm)", lambda p: p.get('ZAvg_nm')),
            ("PDI", lambda p: p.get('PDI')),
            ("Span", lambda p: p.get('Span')),
            ("Mean aritmetico (nm)", lambda p: p.get('AriMean_nm')),
            ("Mean geometrico (nm)", lambda p: p.get('GeoMean_nm')),
        ]

        for row_i, (label, fn) in enumerate(ROWS):
            tag = "odd" if row_i % 2 else "even"
            if fn is None:
                vals = [label] + [""] * 5
                self.tree.insert("", tk.END, values=vals, tags=("section",))
                self.tree.tag_configure("section",
                                        background="#EFF6FF",
                                        font=("Helvetica", 9, "bold"))
            else:
                cells = [label]
                for path in sel:
                    v = fn(self.samples[path]['params'])
                    cells.append(str(v) if v is not None else "—")
                while len(cells) < 6:
                    cells.append("")
                self.tree.insert("", tk.END, values=cells, tags=(tag,))


    # ── Esportazione ─────────────────────────────────────────────────────────

    @staticmethod
    def _write_xlsx(dest_path, d):
        wb = openpyxl.Workbook()

        hdr_font  = Font(name='Arial', bold=True, size=10)
        hdr_fill  = PatternFill('solid', start_color='1E3A5F')
        hdr_font_w = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        sec_fill  = PatternFill('solid', start_color='DBEAFE')
        sec_font  = Font(name='Arial', bold=True, size=10)
        cell_font = Font(name='Arial', size=10)
        thin = Side(style='thin', color='CBD5E1')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal='center')

        def style_header(cell, text):
            cell.value = text
            cell.font  = hdr_font_w
            cell.fill  = hdr_fill
            cell.alignment = center
            cell.border = border

        # ── Foglio ACF ────────────────────────────────────────────────────────
        ws_acf = wb.active
        ws_acf.title = 'ACF'
        headers = ['tau (µs)', 'g₁ exp', 'g₁ fit', 'Residuo']
        for ci, h in enumerate(headers, 1):
            style_header(ws_acf.cell(1, ci), h)
        ws_acf.column_dimensions['A'].width = 14
        for col in 'BCD':
            ws_acf.column_dimensions[col].width = 13

        tau   = d['tau']
        acf   = d['acf']
        fit   = d['fit']
        resid = d['resid']
        off   = d.get('fit_offset', 6)
        n_fit, n_res = len(fit), len(resid)

        for ri, (t, g) in enumerate(zip(tau, acf), 2):
            if t <= 0:
                continue
            fi = ri - 2 - off  # indice nel vettore fit/resid
            ws_acf.cell(ri, 1, round(float(t), 4)).font = cell_font
            ws_acf.cell(ri, 2, round(float(g), 6)).font = cell_font
            if 0 <= fi < n_fit:
                ws_acf.cell(ri, 3, round(float(fit[fi]), 6)).font = cell_font
            if 0 <= fi < n_res:
                ws_acf.cell(ri, 4, round(float(resid[fi]), 6)).font = cell_font

        # ── Foglio Distribuzione in intensità ────────────────────────────────
        ws_sz = wb.create_sheet('Distribuzione intensità')
        for ci, h in enumerate(['Diametro (nm)', 'Intensità (%)', 'Cumulativa (%)'], 1):
            style_header(ws_sz.cell(1, ci), h)
        for col in ('A','B','C'):
            ws_sz.column_dimensions[col].width = 16

        for ri, (s, iv, cv) in enumerate(zip(d['sz'], d['intens'], d['cumul']), 2):
            ws_sz.cell(ri, 1, round(float(s),  4)).font = cell_font
            ws_sz.cell(ri, 2, round(float(iv), 4)).font = cell_font
            ws_sz.cell(ri, 3, round(float(cv), 4)).font = cell_font

        # ── Foglio Distribuzione per numero (Rayleigh: N ∝ I/d⁶) ─────────────
        num_freq, num_cumul = NSZViewer._to_number_dist(d['sz'], d['intens'])
        ws_nb = wb.create_sheet('Distribuzione numero')
        for ci, h in enumerate(['Diametro (nm)', 'Numero (%)', 'Cumulativa (%)'], 1):
            style_header(ws_nb.cell(1, ci), h)
        for col in ('A','B','C'):
            ws_nb.column_dimensions[col].width = 16

        for ri, (s, nv, cv) in enumerate(zip(d['sz'], num_freq, num_cumul), 2):
            ws_nb.cell(ri, 1, round(float(s),  4)).font = cell_font
            ws_nb.cell(ri, 2, round(float(nv), 4)).font = cell_font
            ws_nb.cell(ri, 3, round(float(cv), 4)).font = cell_font

        # ── Foglio Parametri ──────────────────────────────────────────────────
        ws_p = wb.create_sheet('Parametri')
        style_header(ws_p.cell(1, 1), 'Parametro')
        style_header(ws_p.cell(1, 2), 'Valore')
        ws_p.column_dimensions['A'].width = 26
        ws_p.column_dimensions['B'].width = 22

        SECTIONS = {
            'Sample_Name': ('── Campione ──', True),
            'Angle_deg':   ('── Acquisizione ──', True),
            'Peak1_nm':    ('── Risultati ──', True),
        }
        LABELS = {
            'Sample_Name': 'Nome campione',
            'Date': 'Data',  'Time': 'Ora',
            'Solvent': 'Solvente', 'Software': 'Software',
            'Angle_deg': 'Angolo (°)', 'Lambda_nm': 'Lunghezza d\'onda (nm)',
            'Temp_C': 'Temperatura (°C)', 'Visc_mPas': 'Viscosità (mPa·s)',
            'n_solvent': 'n solvente',
            'Peak1_nm': 'Peak 1 (nm)', 'Mean1_nm': 'Mean 1 (nm)',
            'SD1_nm': 'SD 1 (nm)', 'Area1_pct': 'Area 1 (%)',
            'Peak2_nm': 'Peak 2 (nm)', 'Mean2_nm': 'Mean 2 (nm)',
            'SD2_nm': 'SD 2 (nm)',
            'D10_nm': 'D10 (nm)', 'D50_nm': 'D50 (nm)', 'D90_nm': 'D90 (nm)',
            'Mode_nm': 'Mode (nm)', 'Median_nm': 'Median (nm)',
            'ZAvg_nm': 'Z-Average (nm)', 'PDI': 'PDI',
            'Span': 'Span', 'GeoMean_nm': 'Mean geometrico (nm)',
            'AriMean_nm': 'Mean aritmetico (nm)',
        }

        ri = 2
        for key, val in d['params'].items():
            if key in SECTIONS:
                sec_label, _ = SECTIONS[key]
                c1, c2 = ws_p.cell(ri, 1), ws_p.cell(ri, 2)
                c1.value = sec_label
                c1.font = sec_font; c1.fill = sec_fill; c1.border = border
                c2.fill = sec_fill; c2.border = border
                ri += 1
            label = LABELS.get(key, key)
            c1 = ws_p.cell(ri, 1, label)
            c2 = ws_p.cell(ri, 2, val if val is not None else '—')
            c1.font = cell_font; c1.border = border
            c2.font = cell_font; c2.border = border
            ri += 1

        wb.save(dest_path)

    def _export_xlsx(self):
        if not self.selected:
            messagebox.showinfo("Esporta", "Nessun campione selezionato.")
            return
        if not XLSX_AVAILABLE:
            messagebox.showerror("Errore", "Libreria openpyxl non disponibile.\nInstallala con: pip install openpyxl")
            return

        if len(self.selected) == 1:
            d = self.samples[self.selected[0]]
            dest = filedialog.asksaveasfilename(
                title="Salva come...",
                initialfile=f'{d["name"]}.xlsx',
                defaultextension='.xlsx',
                filetypes=[("Excel", "*.xlsx")])
            if dest:
                self._write_xlsx(dest, d)
                messagebox.showinfo("Esportazione completata", f"Salvato:\n{dest}")
        else:
            folder = filedialog.askdirectory(title="Cartella di destinazione")
            if not folder:
                return
            saved = []
            for path in self.selected:
                d = self.samples[path]
                out = os.path.join(folder, f'{d["name"]}.xlsx')
                self._write_xlsx(out, d)
                saved.append(os.path.basename(out))
            messagebox.showinfo("Esportazione completata",
                                f"Salvati {len(saved)} file in:\n{folder}\n\n" +
                                "\n".join(saved))

    # ── Utility ───────────────────────────────────────────────────────────────

    def _short_label(self, name):
        if len(name) > 35:
            return name[:32] + '…'
        return name


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    import sys

    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("Nota: tkinterdnd2 non disponibile, drag & drop disabilitato.")

    app = NSZViewer(root)

    if not OLEFILE_AVAILABLE:
        messagebox.showwarning("Dipendenza mancante",
                               "Libreria olefile non disponibile: non sarà possibile aprire i file .nsz.\n"
                               "Installala con: pip install olefile")

    # Se passati argomenti dalla riga di comando, carica quei file
    if len(sys.argv) > 1:
        files = [a for a in sys.argv[1:] if a.lower().endswith('.nsz')]
        if files:
            app._load_files(files)

    root.mainloop()


if __name__ == '__main__':
    main()
