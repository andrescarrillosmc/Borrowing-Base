from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

if __package__ in (None, ""):
    sys.path.append(str(Path(__file__).resolve().parents[1]))

from borrowing_base_workbench.analysis import DEFAULT_WORKBOOK, analyze_workbook, diagnosis_to_json, diagnosis_to_markdown
from borrowing_base_workbench.excel_runner import probe_excel_workbook, run_pro_forma_workbook
from borrowing_base_workbench.validation import build_commentary, validate_scenario

APP_BG = "#eef3f9"
HERO_BLUE = "#0b64a0"
PANEL_BG = "#ffffff"
PANEL_BORDER = "#ccd9e8"
SURFACE_BG = "#f5f7fb"
SLATE = "#17324d"

FIELD_FORMATS = {
    "ltm_revenue": "currency",
    "ltm_adj_ebitda": "currency",
    "drawn_revolver": "currency",
    "first_out_balance": "currency",
    "pari_passu": "currency",
    "bdc_balance": "currency",
    "total_sm_balance": "currency",
    "cash_balance": "currency",
    "interest_coverage": "multiple",
    "purchase_price": "percent",
    "pik_pct": "percent",
    "spread": "percent",
    "sofr_floor": "percent",
    "attach_point": "multiple",
}


def _fmt_currency(value) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"${float(value):,.0f}"
    except (TypeError, ValueError):
        return str(value)


def _fmt_percent(value) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"{float(value) * 100:,.1f}%"
    except (TypeError, ValueError):
        return str(value)


def _fmt_multiple(value) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"{float(value):.2f}x"
    except (TypeError, ValueError):
        return str(value)


def _parse_numeric_text(value, kind: str | None = None) -> float | None:
    if value in (None, ""):
        return None
    text = str(value).strip()
    if not text:
        return None
    had_percent = "%" in text
    cleaned = text.replace("$", "").replace(",", "").replace("x", "").replace("X", "").replace("%", "").strip()
    if not cleaned:
        return None
    number = float(cleaned)
    if kind == "percent":
        if had_percent or number > 1:
            return number / 100.0
    return number


def _format_input_value(value, kind: str | None) -> str:
    number = _parse_numeric_text(value, kind)
    if number is None:
        return ""
    if kind == "currency":
        return f"${number:,.0f}"
    if kind == "percent":
        return f"{number:.2%}"
    if kind == "multiple":
        return f"{number:.2f}x"
    return str(value)


class BorrowingBaseWorkbench(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Borrowing Base Workbench")
        self.geometry("1380x920")
        self.minsize(1220, 820)
        self.configure(bg=APP_BG)

        self.workbook_path = tk.StringVar(value=str(DEFAULT_WORKBOOK))
        self.status_text = tk.StringVar(value="Select the governed workbook, enter a scenario, and run the pro forma.")
        self.diagnosis = None
        self.form_vars: dict[str, tk.StringVar] = {}
        self.last_probe_result: dict | None = None
        self._loading_overlay: tk.Toplevel | None = None
        self._tab_transition_job: str | None = None

        self._configure_theme()
        self._build_shell()
        self._try_load_default()

    def _configure_theme(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background=APP_BG)
        style.configure("Card.TFrame", background="white")
        style.configure("App.TLabel", background=APP_BG, foreground=SLATE, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=APP_BG, foreground="#56708f", font=("Segoe UI", 10))
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            background=HERO_BLUE,
            foreground="white",
            borderwidth=0,
        )
        style.map("Primary.TButton", background=[("active", "#0a588e")])
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            background="#dbe7f5",
            foreground=SLATE,
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", "#c8d8ee")])
        style.configure(
            "App.TEntry",
            fieldbackground="white",
            bordercolor="#c3d4e8",
            lightcolor="#c3d4e8",
            darkcolor="#c3d4e8",
            padding=6,
        )
        style.configure("App.TNotebook", background=APP_BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure(
            "App.TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 7),
            background="#e8eff7",
            foreground=SLATE,
        )
        style.map("App.TNotebook.Tab", background=[("selected", HERO_BLUE)], foreground=[("selected", "white")])
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10), fieldbackground="white", background="white")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#dbe7f5", foreground=SLATE)
        style.configure("TLabelframe", background=SURFACE_BG, borderwidth=1)
        style.configure("TLabelframe.Label", background=SURFACE_BG, foreground=SLATE, font=("Segoe UI", 11, "bold"))
        style.configure(
            "Minimal.Vertical.TScrollbar",
            gripcount=0,
            background="#8ea0b4",
            troughcolor="#e7edf4",
            bordercolor="#e7edf4",
            arrowcolor="#8ea0b4",
            relief="flat",
            arrowsize=10,
        )
        style.map("Minimal.Vertical.TScrollbar", background=[("active", "#73869a")])
        style.configure(
            "Minimal.Horizontal.TScrollbar",
            gripcount=0,
            background="#8ea0b4",
            troughcolor="#e7edf4",
            bordercolor="#e7edf4",
            arrowcolor="#8ea0b4",
            relief="flat",
            arrowsize=10,
        )
        style.map("Minimal.Horizontal.TScrollbar", background=[("active", "#73869a")])

    def _build_shell(self) -> None:
        shell = tk.Frame(self, bg=APP_BG)
        shell.pack(fill="both", expand=True, padx=14, pady=14)

        hero = self._create_rounded_panel(shell, bg=HERO_BLUE, border="#0a588e", radius=22, padding=(22, 18))
        tk.Label(
            hero,
            text="Borrowing Base Pro Forma Workbench",
            bg=HERO_BLUE,
            fg="white",
            font=("Segoe UI", 22, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            hero,
            text="Governed workbook inputs, live baseline reads, and staged pro forma results in one place.",
            bg=HERO_BLUE,
            fg="#dbe7f5",
            font=("Segoe UI", 11),
            anchor="w",
        ).pack(fill="x", pady=(5, 0))

        top = self._create_rounded_panel(shell, bg=PANEL_BG, border=PANEL_BORDER, radius=18, padding=(16, 14))

        tk.Label(top, text="Governed Workbook", bg=PANEL_BG, fg=SLATE, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(top, textvariable=self.workbook_path, width=92, style="App.TEntry").pack(side="left", padx=(12, 10))
        ttk.Button(top, text="Browse", command=self._browse, style="Secondary.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Read Current Model", command=self._probe_excel, style="Primary.TButton").pack(side="left")

        notebook_shell = self._create_rounded_panel(shell, bg=PANEL_BG, border=PANEL_BORDER, radius=20, padding=(14, 14), fill="both", expand=True)
        notebook = ttk.Notebook(notebook_shell, style="App.TNotebook")
        notebook.pack(fill="both", expand=True)
        self.notebook = notebook
        notebook.bind("<<NotebookTabChanged>>", self._animate_tab_change)

        self.overview_tab = ttk.Frame(notebook, padding=12, style="App.TFrame")
        self.scenario_tab = ttk.Frame(notebook, padding=12, style="App.TFrame")
        self.results_tab = ttk.Frame(notebook, padding=12, style="App.TFrame")
        self.admin_tab = ttk.Frame(notebook, padding=12, style="App.TFrame")

        notebook.add(self.overview_tab, text="Overview")
        notebook.add(self.scenario_tab, text="Scenario")
        notebook.add(self.results_tab, text="Results")
        notebook.add(self.admin_tab, text="Admin")

        self._build_overview_tab()
        self._build_scenario_tab()
        self._build_results_tab()
        self._build_admin_tab()
        self._populate_overview()
        self._populate_admin()
        self._clear_results_panel()

        status = tk.Label(
            self._create_rounded_panel(shell, bg="#dbe7f5", border="#c3d4e8", radius=16, padding=(14, 12)),
            textvariable=self.status_text,
            bg="#dbe7f5",
            fg=SLATE,
            anchor="w",
            font=("Segoe UI", 10),
        )
        status.pack(fill="x")

    def _build_overview_tab(self) -> None:
        frame = tk.Frame(self.overview_tab, bg="#f5f7fb")
        frame.pack(fill="both", expand=True)

        hero = tk.Frame(frame, bg="#0b64a0", padx=20, pady=18)
        hero.pack(fill="x", pady=(0, 12))

        tk.Label(
            hero,
            text="Borrowing Base Pro Forma Workbench",
            bg="#0b64a0",
            fg="white",
            font=("Segoe UI", 24, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            hero,
            text="Run governed scenarios against the current borrowing base workbook without touching the master file.",
            bg="#0b64a0",
            fg="white",
            font=("Segoe UI", 12),
            anchor="w",
        ).pack(fill="x", pady=(6, 0))

        cards = tk.Frame(frame, bg="#f5f7fb")
        cards.pack(fill="x", pady=(0, 12))
        for idx in range(3):
            cards.grid_columnconfigure(idx, weight=1)

        self.overview_workbook_card = self._make_overview_card(cards, 0, "Governed Workbook", "No workbook loaded yet.")
        self.overview_snapshot_card = self._make_overview_card(cards, 1, "Current Model Snapshot", "Read Current Model to capture the live baseline.")
        self.overview_mode_card = self._make_overview_card(cards, 2, "Admin Workflow", "Replace the governed workbook monthly when a new template arrives.")

        body = tk.Frame(frame, bg="#f5f7fb")
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)

        how_to = tk.LabelFrame(body, text="How To Run A Scenario", padx=16, pady=12)
        how_to.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.overview_how_to = tk.Text(how_to, wrap="word", height=14, relief="flat", bg="white", font=("Segoe UI", 11))
        self.overview_how_to.pack(fill="both", expand=True)

        operator = tk.LabelFrame(body, text="What The App Does For You", padx=16, pady=12)
        operator.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.overview_operator = tk.Text(operator, wrap="word", height=14, relief="flat", bg="white", font=("Segoe UI", 11))
        self.overview_operator.pack(fill="both", expand=True)

    def _make_overview_card(self, parent: tk.Widget, column: int, title: str, value: str) -> tk.StringVar:
        card = tk.Frame(parent, bg="white", highlightbackground="#d0d7e2", highlightthickness=1, padx=16, pady=14)
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0))
        tk.Label(card, text=title, bg="white", fg="#4a5a70", font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x")
        var = tk.StringVar(value=value)
        tk.Label(card, textvariable=var, bg="white", fg="#0f1724", font=("Segoe UI", 12), anchor="w", justify="left", wraplength=360).pack(fill="x", pady=(8, 0))
        return var

    def _bind_mousewheel_to_canvas(self, canvas: tk.Canvas, widget: tk.Widget) -> None:
        def _on_mousewheel(event):
            delta = 0
            if getattr(event, "delta", 0):
                delta = int(-1 * (event.delta / 120))
            elif getattr(event, "num", None) == 5:
                delta = 1
            elif getattr(event, "num", None) == 4:
                delta = -1
            if delta:
                canvas.yview_scroll(delta, "units")

        def _bind(_event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind(_event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        widget.bind("<Enter>", _bind)
        widget.bind("<Leave>", _unbind)

    def _create_rounded_panel(
        self,
        parent: tk.Widget,
        *,
        bg: str = PANEL_BG,
        border: str = PANEL_BORDER,
        radius: int = 18,
        padding: tuple[int, int] = (16, 14),
        fill: str = "x",
        expand: bool = False,
    ) -> tk.Frame:
        outer = tk.Frame(parent, bg=APP_BG, highlightthickness=0, bd=0)
        outer.pack(fill=fill, expand=expand, pady=(0, 12))
        inner = tk.Frame(
            outer,
            bg=bg,
            padx=padding[0],
            pady=padding[1],
            highlightbackground=border,
            highlightthickness=1,
            bd=0,
        )
        inner.pack(fill=fill, expand=expand)
        return inner

    def _animate_tab_change(self, _event=None) -> None:
        if self._tab_transition_job:
            self.after_cancel(self._tab_transition_job)
            self._tab_transition_job = None
        def _settle() -> None:
            try:
                self.update_idletasks()
            finally:
                self._tab_transition_job = None

        self._tab_transition_job = self.after(16, _settle)

    def _show_loading_overlay(self, title: str, message: str) -> None:
        if not self.winfo_ismapped():
            return
        self._hide_loading_overlay()
        overlay = tk.Toplevel(self)
        overlay.transient(self)
        overlay.overrideredirect(True)
        overlay.configure(bg="#10243a")
        overlay.attributes("-topmost", True)
        overlay.grab_set()

        width = 420
        height = 180
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (width // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (height // 2)
        overlay.geometry(f"{width}x{height}+{x}+{y}")

        card = tk.Frame(overlay, bg="#10243a", padx=28, pady=24)
        card.pack(fill="both", expand=True)
        tk.Label(card, text=title, bg="#10243a", fg="white", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(card, text=message, bg="#10243a", fg="#dbe7f5", font=("Segoe UI", 11), wraplength=360, justify="left").pack(anchor="w", pady=(10, 16))
        progress = ttk.Progressbar(card, mode="indeterminate", length=360)
        progress.pack(anchor="w")
        progress.start(12)

        self._loading_overlay = overlay
        self.update_idletasks()

    def _hide_loading_overlay(self) -> None:
        if not self._loading_overlay:
            return
        try:
            self._loading_overlay.grab_release()
        except tk.TclError:
            pass
        try:
            self._loading_overlay.destroy()
        except tk.TclError:
            pass
        self._loading_overlay = None

    def _build_scenario_tab(self) -> None:
        outer = tk.Frame(self.scenario_tab, bg="#eef3f9")
        outer.pack(fill="both", expand=True)

        controls = tk.Frame(outer, bg="#eef3f9")
        controls.pack(fill="x", pady=(0, 10))
        ttk.Button(controls, text="Validate Scenario", command=self._validate_scenario, style="Secondary.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Run Pro Forma", command=self._run_pro_forma, style="Primary.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Read Current Model", command=self._probe_excel, style="Secondary.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Clear Scenario", command=self._clear_scenario, style="Secondary.TButton").pack(side="left")

        canvas = tk.Canvas(outer, highlightthickness=0, bg="#eef3f9")
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview, style="Minimal.Vertical.TScrollbar")
        frame = tk.Frame(canvas, bg="#f4f4f4")
        frame.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._bind_mousewheel_to_canvas(canvas, outer)
        self._bind_mousewheel_to_canvas(canvas, frame)
        self.scenario_canvas = canvas

        defaults = {
            "security_type": "First Lien",
            "loan_denomination": "USD",
            "country": "United States",
            "rate_type": "Floating",
            "payment_frequency": "M",
            "purchase_price": "100.00%",
            "pik_pct": "0.00%",
        }
        self.readiness_required_keys = [
            "company_name",
            "security_type",
            "ltm_revenue",
            "ltm_adj_ebitda",
            "bdc_balance",
            "total_sm_balance",
            "purchase_price",
            "industry_classification",
            "attach_point",
        ]

        for key in [
            "company_name",
            "security_type",
            "ltm_revenue",
            "ltm_adj_ebitda",
            "drawn_revolver",
            "first_out_balance",
            "pari_passu",
            "bdc_balance",
            "total_sm_balance",
            "cash_balance",
            "interest_coverage",
            "loan_denomination",
            "purchase_price",
            "country",
            "investment_date",
            "maturity_date",
            "rate_type",
            "industry_classification",
            "payment_frequency",
            "pik_pct",
            "spread",
            "sofr_floor",
            "attach_point",
        ]:
            self.form_vars[key] = tk.StringVar(value=defaults.get(key, ""))

        header = tk.Label(
            frame,
            text="BORROWING BASE PRO FORMA - DEAL TEAM INPUT FORM",
            bg="#0b64a0",
            fg="white",
            font=("Segoe UI", 22, "bold"),
            anchor="w",
            padx=12,
            pady=8,
        )
        header.grid(row=0, column=0, columnspan=3, sticky="ew")

        tk.Label(frame, text="Status:", bg="#f4f4f4", fg="black", font=("Segoe UI", 16, "bold"), anchor="w").grid(
            row=1, column=0, sticky="w", padx=(4, 8), pady=(8, 10)
        )
        self.readiness_label = tk.Label(
            frame,
            text="Fill Required Fields",
            bg="#d9deef",
            fg="#c00000",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
            padx=12,
            pady=6,
        )
        self.readiness_label.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(8, 10))

        row = 2
        row = self._build_section_header(frame, row, "COMPANY INFORMATION")
        row = self._add_input_row(frame, row, "Company Name:", "company_name", "entry", "", wide=True)
        row = self._add_input_row(frame, row, "Security Type:", "security_type", "combo", "(First Lien or FILO)", values=["First Lien", "FILO"], wide=True)

        row += 1
        row = self._build_section_header(frame, row, "PHASE 1: FINANCIAL METRICS (SM Support Tab)")
        row = self._build_column_header(frame, row)
        phase_one_rows = [
            ("1. LTM Revenue", "ltm_revenue", "entry", "In dollars (e.g., $95,000,000)", []),
            ("2. LTM Adj. EBITDA", "ltm_adj_ebitda", "entry", "In dollars (e.g., $12,500,000)", []),
            ("3. Drawn Revolver Balance", "drawn_revolver", "entry", "In dollars (enter $0 if none)", []),
            ("4. Drawn First Out Balance", "first_out_balance", "entry", "In dollars (enter $0 if none)", []),
            ("5. Pari Passu", "pari_passu", "entry", "In dollars (e.g., $7,500,000)", []),
            ("6. BDC Balance (Debt Balance)", "bdc_balance", "entry", "KEY: BDC's specific investment amount", []),
            ("7. Total SM Balance", "total_sm_balance", "entry", "Total Star Mountain Exposure", []),
            ("8. Cash Balance", "cash_balance", "entry", "Company's cash at close (enter $0 if none)", []),
            ("9. Current Interest Coverage", "interest_coverage", "entry", "Ratio (e.g., 1.05x)", []),
        ]
        for label, key, widget_type, note, values in phase_one_rows:
            row = self._add_input_row(frame, row, label, key, widget_type, note, values=values)

        row += 1
        row = self._build_section_header(frame, row, "PHASE 2: LOAN TERMS (Loan Tape Tab)")
        row = self._build_column_header(frame, row)
        phase_two_rows = [
            ("10. Loan Denomination", "loan_denomination", "combo", "USD or CAD", ["USD", "CAD"]),
            ("11. Purchase Price", "purchase_price", "entry", "% of par (e.g., 95.00% = 0.95)", []),
            ("12. Country", "country", "combo", "United States", ["United States", "Canada", "Other"]),
            ("13. Investment Date", "investment_date", "entry", "Date (e.g., 2/28/2026)", []),
            ("14. Maturity Date", "maturity_date", "entry", "Date (e.g., 2/28/2030)", []),
            ("15. Floating Rate?", "rate_type", "combo", "Floating or Fixed", ["Floating", "Fixed"]),
            ("16. S&P Industry Classification", "industry_classification", "combo", "MUST match approved industry list exactly", []),
            ("17. Payment Frequency", "payment_frequency", "combo", "M (Monthly) or Q (Quarterly)", ["M", "Q"]),
            ("18. PIK %", "pik_pct", "entry", "% PIK (e.g., 50.00% or 0.00%)", []),
            ("19. Spread", "spread", "entry", "% (e.g., 9.50% = 0.095)", []),
            ("20. SOFR Floor", "sofr_floor", "entry", "% (e.g., 2.00% = 0.02)", []),
            ("21. SM Debt Attach Point", "attach_point", "entry", "Leverage ratio (e.g., 4.20x)", []),
        ]
        for label, key, widget_type, note, values in phase_two_rows:
            row = self._add_input_row(frame, row, label, key, widget_type, note, values=values)

        frame.grid_columnconfigure(0, weight=0, minsize=530)
        frame.grid_columnconfigure(1, weight=0, minsize=340)
        frame.grid_columnconfigure(2, weight=1, minsize=340)

        for var in self.form_vars.values():
            var.trace_add("write", self._update_readiness_status)
        for key, var in self.form_vars.items():
            if key in FIELD_FORMATS and var.get():
                var.set(_format_input_value(var.get(), FIELD_FORMATS[key]))
        self._update_readiness_status()

    def _build_section_header(self, parent: tk.Widget, row: int, text: str) -> int:
        label = tk.Label(
            parent,
            text=text,
            bg="#4472c4",
            fg="white",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
            padx=8,
            pady=4,
        )
        label.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        return row + 1

    def _build_column_header(self, parent: tk.Widget, row: int) -> int:
        headers = ["Input Field", "Your Input", "Format / Notes"]
        widths = [(0, "w"), (1, "center"), (2, "w")]
        for column, anchor in widths:
            label = tk.Label(
                parent,
                text=headers[column],
                bg="#d9deef",
                fg="black",
                font=("Segoe UI", 12, "bold"),
                anchor=anchor,
                padx=6,
                pady=4,
            )
            label.grid(row=row, column=column, sticky="ew")
        return row + 1

    def _add_input_row(
        self,
        parent: tk.Widget,
        row: int,
        label_text: str,
        key: str,
        widget_type: str,
        note_text: str,
        values: list[str] | None = None,
        wide: bool = False,
    ) -> int:
        is_key_row = "BDC Balance" in label_text
        fg = "#ff0000" if is_key_row else "black"
        tk.Label(
            parent,
            text=label_text,
            bg="#f4f4f4",
            fg=fg,
            font=("Segoe UI", 14, "bold"),
            anchor="w",
            padx=4,
            pady=6,
        ).grid(row=row, column=0, sticky="w")

        input_bg = "#e6edf8"
        if wide:
            input_parent = tk.Frame(parent, bg=input_bg, height=46)
            input_parent.grid(row=row, column=1, sticky="ew", pady=3)
        else:
            input_parent = tk.Frame(parent, bg=input_bg, height=42)
            input_parent.grid(row=row, column=1, sticky="ew", pady=2)
        input_parent.grid_propagate(False)
        input_parent.columnconfigure(0, weight=1)
        input_parent.rowconfigure(0, weight=1)

        if widget_type == "combo":
            widget = ttk.Combobox(
                input_parent,
                textvariable=self.form_vars[key],
                values=values or [],
                state="readonly" if values else "normal",
                width=30,
            )
        else:
            widget = tk.Entry(
                input_parent,
                textvariable=self.form_vars[key],
                relief="flat",
                bg=input_bg,
                font=("Segoe UI", 14),
                insertbackground="black",
            )
            if key in FIELD_FORMATS:
                widget.bind("<FocusOut>", lambda _event, field_key=key: self._format_field_value(field_key))
        widget.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)

        tk.Label(
            parent,
            text=note_text,
            bg="#f4f4f4",
            fg="black",
            font=("Segoe UI", 13),
            anchor="w",
            justify="left",
            padx=6,
            pady=6,
            wraplength=360,
        ).grid(row=row, column=2, sticky="w")
        return row + 1

    def _update_readiness_status(self, *_args) -> None:
        ready = all(self.form_vars[key].get().strip() for key in self.readiness_required_keys)
        if ready:
            self.readiness_label.configure(text="READY TO RUN", bg="#dcead3", fg="#008000")
        else:
            self.readiness_label.configure(text="Fill Required Fields", bg="#d9deef", fg="#c00000")

    def _format_field_value(self, key: str) -> None:
        kind = FIELD_FORMATS.get(key)
        if not kind:
            return
        current = self.form_vars[key].get()
        if not current.strip():
            return
        try:
            self.form_vars[key].set(_format_input_value(current, kind))
        except ValueError:
            return

    def _collect_form_values(self) -> dict[str, str]:
        values: dict[str, str] = {}
        for key, var in self.form_vars.items():
            raw = var.get().strip()
            kind = FIELD_FORMATS.get(key)
            if not kind:
                values[key] = raw
                continue
            try:
                number = _parse_numeric_text(raw, kind)
            except ValueError:
                values[key] = raw
                continue
            values[key] = "" if number is None else str(number)
        return values

    def _clear_tree(self, tree: ttk.Treeview) -> None:
        for item in tree.get_children():
            tree.delete(item)

    def _set_text_content(self, widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def _build_labeled_tree(
        self,
        parent: tk.Widget,
        title: str,
        columns: list[tuple[str, int]],
        *,
        height: int,
        stretch_last: bool = True,
    ) -> ttk.Treeview:
        section = tk.LabelFrame(parent, text=title, padx=10, pady=8, bg="#f5f7fb")
        section.pack(fill="x", pady=(0, 10))
        frame = ttk.Frame(section)
        frame.pack(fill="x", expand=True)

        tree = ttk.Treeview(frame, columns=[name for name, _ in columns], show="headings", height=height)
        for idx, (name, width) in enumerate(columns):
            tree.heading(name, text=name)
            tree.column(name, width=width, anchor="w", stretch=stretch_last if idx == len(columns) - 1 else False)

        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview, style="Minimal.Horizontal.TScrollbar")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview, style="Minimal.Vertical.TScrollbar")
        tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        return tree

    def _build_results_tab(self) -> None:
        outer = tk.Frame(self.results_tab, bg="#f5f7fb")
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg="#0b64a0", padx=16, pady=12)
        header.pack(fill="x", pady=(0, 10))
        title_row = tk.Frame(header, bg="#0b64a0")
        title_row.pack(fill="x")
        tk.Label(
            title_row,
            text="PRO FORMA RESULTS",
            bg="#0b64a0",
            fg="white",
            font=("Segoe UI", 20, "bold"),
            anchor="w",
        ).pack(side="left")
        ttk.Button(title_row, text="Export Results", command=self._export_results, style="Secondary.TButton").pack(side="right")
        self.results_banner = tk.StringVar(value="Run validation, current model read, or pro forma to populate this report.")
        tk.Label(
            header,
            textvariable=self.results_banner,
            bg="#0b64a0",
            fg="white",
            font=("Segoe UI", 11),
            anchor="w",
        ).pack(fill="x", pady=(6, 0))

        issue_section = tk.LabelFrame(outer, text="Scenario Checks", padx=10, pady=8, bg="#f5f7fb")
        issue_section.pack(fill="x", pady=(0, 10))
        self.issue_tree = ttk.Treeview(issue_section, columns=("severity", "field", "message"), show="headings", height=5)
        for column, width in [("severity", 120), ("field", 180), ("message", 900)]:
            self.issue_tree.heading(column, text=column.title())
            self.issue_tree.column(column, width=width, anchor="w", stretch=column == "message")
        self.issue_tree.pack(fill="x")

        canvas = tk.Canvas(outer, highlightthickness=0, bg="#f5f7fb")
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview, style="Minimal.Vertical.TScrollbar")
        report = tk.Frame(canvas, bg="#f5f7fb")
        report.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=report, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._bind_mousewheel_to_canvas(canvas, outer)
        self._bind_mousewheel_to_canvas(canvas, report)
        self.results_canvas = canvas

        self.availability_tree = self._build_labeled_tree(
            report,
            "Availability Impact",
            [("Metric", 260), ("Before", 150), ("After", 150), ("Change", 140), ("Notes", 420)],
            height=6,
        )
        self.portfolio_tree = self._build_labeled_tree(
            report,
            "Portfolio Composition",
            [("Metric", 260), ("Value", 160), ("Calculated From", 560)],
            height=5,
        )
        self.concentration_tree = self._build_labeled_tree(
            report,
            "Concentration Limits",
            [
                ("Limit Type", 340),
                ("Limit %", 90),
                ("Applicable Limit $", 130),
                ("Prior Actual $", 120),
                ("Prior Excess $", 120),
                ("Pro Forma Actual $", 130),
                ("Pro Forma Excess $", 130),
                ("Delta Actual $", 120),
                ("Delta Excess $", 120),
            ],
            height=10,
            stretch_last=False,
        )
        commentary_section = tk.LabelFrame(report, text="Commentary", padx=10, pady=8, bg="#f5f7fb")
        commentary_section.pack(fill="both", expand=True, pady=(0, 10))
        self.commentary_text = tk.Text(commentary_section, wrap="word", height=10, relief="flat", bg="white", font=("Segoe UI", 11))
        self.commentary_text.pack(fill="both", expand=True)
        self.commentary_text.configure(state="disabled")

        self.model_metrics_tree = self._build_labeled_tree(
            report,
            "Workbook Metrics",
            [("Metric", 260), ("Current", 150), ("Pro Forma", 150), ("Notes", 430)],
            height=6,
        )

    def _build_admin_tab(self) -> None:
        outer = tk.Frame(self.admin_tab, bg="#f5f7fb")
        outer.pack(fill="both", expand=True)

        hero = tk.Frame(outer, bg="#17324d", padx=16, pady=14)
        hero.pack(fill="x", pady=(0, 10))
        tk.Label(
            hero,
            text="Admin Control Center",
            bg="#17324d",
            fg="white",
            font=("Segoe UI", 20, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            hero,
            text="The safe operating model is monthly workbook replacement: update the governed borrowing base file, then let the app inherit the revised logic.",
            bg="#17324d",
            fg="white",
            font=("Segoe UI", 11),
            anchor="w",
            justify="left",
            wraplength=1100,
        ).pack(fill="x", pady=(4, 0))

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(0, 10))
        ttk.Button(controls, text="Reload Current Workbook", command=self._load_workbook).pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Browse Monthly Workbook", command=self._browse).pack(side="left", padx=(0, 8))
        ttk.Button(controls, text="Export Diagnosis", command=self._export_diagnosis).pack(side="left")

        self.admin_text = tk.Text(outer, wrap="word", height=10, relief="flat", bg="white", font=("Segoe UI", 11))
        self.admin_text.pack(fill="x", pady=(0, 10))
        self.admin_text.configure(state="disabled")

        self.admin_surface_tree = self._build_labeled_tree(
            outer,
            "Monthly Workbook Surfaces",
            [("Surface", 220), ("Owner Action", 320), ("How The App Uses It", 500)],
            height=8,
        )
        self.capacity_tree = self._build_labeled_tree(
            outer,
            "Capacity And Range Health",
            [("Area", 260), ("Current Row", 90), ("Ceiling", 90), ("Headroom", 90), ("Operator Note", 520)],
            height=5,
        )
        self.policy_tree = self._build_labeled_tree(
            outer,
            "Controlled Write Surfaces",
            [("Sheet", 200), ("Permission", 120), ("Allowed Range", 260), ("Purpose", 320), ("Risk", 80)],
            height=6,
        )

    def _browse(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel Workbooks", "*.xlsm *.xlsx"), ("All Files", "*.*")])
        if path:
            self.workbook_path.set(path)
            self._load_workbook()

    def _try_load_default(self) -> None:
        if Path(self.workbook_path.get()).exists():
            self._load_workbook()

    def _populate_form_dropdowns(self) -> None:
        if not self.diagnosis:
            return
        stack = [self.scenario_tab]
        while stack:
            node = stack.pop()
            for child in node.winfo_children():
                stack.append(child)
                if isinstance(child, ttk.Combobox):
                    if str(child.cget("textvariable")) == str(self.form_vars["industry_classification"]):
                        child.configure(values=self.diagnosis.industries, state="readonly")
                        return

    def _populate_overview(self) -> None:
        if not self.diagnosis:
            self.overview_workbook_card.set("No workbook loaded yet.")
            self.overview_snapshot_card.set("Read Current Model to capture the live baseline.")
            self.overview_mode_card.set("Replace the governed workbook monthly when a new template arrives.")
            self._set_text_content(
                self.overview_how_to,
                "1. Load the governed workbook.\n"
                "2. Go to Scenario and complete the deal team inputs.\n"
                "3. Click Read Current Model to capture the baseline.\n"
                "4. Click Run Pro Forma to compare before and after results.\n"
                "5. Review concentration limits, commentary, and workbook metrics on Results.",
            )
            self._set_text_content(
                self.overview_operator,
                "This app keeps the monthly borrowing base workbook as the source of truth.\n\n"
                "Deal teams use the Scenario tab to test a deal. Operators use the Admin tab to swap in the new monthly workbook and confirm the governed tabs, ranges, and logic remain healthy.",
            )
            return

        measurement = self.diagnosis.measurement_date or "Unknown measurement date"
        version = self.diagnosis.workbook_version or "Version not labeled"
        workbook_name = Path(self.diagnosis.workbook_path).name
        self.overview_workbook_card.set(f"{workbook_name}\n{version}\n{measurement}")

        if self.last_probe_result and self.last_probe_result.get("status") == "ok":
            metrics = self.last_probe_result.get("metrics", {})
            self.overview_snapshot_card.set(
                f"Availability {_fmt_currency(metrics.get('availability'))}\n"
                f"Adjusted BV {_fmt_currency(metrics.get('aggregate_adjusted_bv'))}\n"
                f"Advance rate {_fmt_percent(metrics.get('weighted_avg_advance_rate'))}"
            )
        else:
            self.overview_snapshot_card.set("No live snapshot loaded yet.\nClick Read Current Model to pull the baseline.")

        self.overview_mode_card.set(
            "Monthly workbook replacement is the recommended admin pattern.\n"
            "Update the governed workbook, reload it here, and the app will inherit the refreshed logic."
        )

        self._set_text_content(
            self.overview_how_to,
            "1. Confirm the workbook path at the top of the app is the current governed borrowing base file.\n"
            "2. Move to Scenario and enter the new investment exactly as the deal team would on the Deal Team Input tab.\n"
            "3. Click Read Current Model before running a pro forma so the app can capture the before-state.\n"
            "4. Click Run Pro Forma to execute the scenario on a staged workbook copy.\n"
            "5. Use Results to compare availability, collateral value, concentration tests, and commentary side by side.\n"
            "6. Click Clear Scenario before starting the next deal.",
        )
        self._set_text_content(
            self.overview_operator,
            "What the app does:\n"
            "- Reads the governed workbook and loads the approved dropdowns and rule surfaces.\n"
            "- Runs scenarios on staged workbook copies so the source workbook is not changed.\n"
            "- Captures current and pro forma borrowing base metrics side by side.\n\n"
            "What the app does not do in deal mode:\n"
            "- It does not rewrite the master workbook.\n"
            "- It does not let deal teams alter monthly concentration logic or loan tape governance.\n"
            "- Those changes belong in the monthly workbook refresh process on the Admin tab.",
        )

    def _populate_admin(self) -> None:
        if not self.diagnosis:
            self._set_text_content(
                self.admin_text,
                "Admin mode is designed for the workbook owner.\n\n"
                "Recommended operating model:\n"
                "1. Receive the new monthly borrowing base workbook.\n"
                "2. Replace the governed workbook path in this app.\n"
                "3. Reload the workbook and confirm the update surfaces below.\n"
                "4. Deal teams then run scenarios against that governed monthly version.",
            )
            self._clear_tree(self.admin_surface_tree)
            self._clear_tree(self.capacity_tree)
            self._clear_tree(self.policy_tree)
            return

        workbook_name = Path(self.diagnosis.workbook_path).name
        self._set_text_content(
            self.admin_text,
            f"Current governed workbook: {workbook_name}\n\n"
            "Admin intent:\n"
            "The app should inherit monthly logic from the workbook, not duplicate or independently override it. "
            "That means the spreadsheet owner updates the monthly workbook first, then reloads it here. "
            "This is the safest way to handle loan tape changes, concentration limit updates, and any formula logic revisions.\n\n"
            "Near-term admin workflow:\n"
            "1. Browse to the new monthly workbook.\n"
            "2. Reload the workbook in this app.\n"
            "3. Confirm the monthly update surfaces and capacity diagnostics below.\n"
            "4. Publish that governed workbook path to the deal team.",
        )

        self._clear_tree(self.admin_surface_tree)
        monthly_actions = {
            "Loan Tape - Settled": ("Refresh actual loan tape and new rows", "Scenarios append against the staged workbook copy using this structure."),
            "SM Support": ("Refresh support data and exposure rows", "Scenario inputs mirror the support-sheet fields."),
            "Availability": ("Confirm facility outputs and formula integrity", "Results read live borrowing base outputs from this tab."),
            "Deal Team Input": ("Confirm UX formulas and linked driver cells", "The scenario UI mirrors this input surface."),
            "Memory": ("Review cached logic notes and assumptions", "Supports governance and operator review."),
            "MAPPING_LAYER": ("Update governed field mappings or validation contract", "Controls app field behavior and write/read expectations."),
            "Portfolio": ("Check insertion markers and eligibility formula propagation", "Pro forma runs rely on this sheet for eligibility and concentration effects."),
            "Concentration Limits": ("Review concentration tests and thresholds", "The app displays these tests side by side in Results."),
        }
        for surface in self.diagnosis.monthly_update_tabs:
            owner_action, app_use = monthly_actions.get(surface, ("Review workbook section", "Used by the governed workbook logic."))
            self.admin_surface_tree.insert("", "end", values=(surface, owner_action, app_use))
        if "Concentration Limits" not in self.diagnosis.monthly_update_tabs:
            owner_action, app_use = monthly_actions["Concentration Limits"]
            self.admin_surface_tree.insert("", "end", values=("Concentration Limits", owner_action, app_use))

        self._clear_tree(self.capacity_tree)
        for item in self.diagnosis.capacities:
            self.capacity_tree.insert(
                "",
                "end",
                values=(
                    item.area,
                    item.current_last_row,
                    item.modeled_ceiling if item.modeled_ceiling is not None else "",
                    item.remaining_headroom if item.remaining_headroom is not None else "",
                    item.note,
                ),
            )

        self._clear_tree(self.policy_tree)
        for policy in self.diagnosis.write_policies:
            self.policy_tree.insert(
                "",
                "end",
                values=(policy.sheet_name, policy.permission, policy.allowed_range, policy.purpose, policy.risk_level),
            )

    def _load_workbook(self) -> None:
        try:
            self._show_loading_overlay("Loading workbook", "Reading the governed workbook structure and refreshing the app surfaces.")
            self.diagnosis = analyze_workbook(self.workbook_path.get())
        except Exception as exc:
            self._hide_loading_overlay()
            messagebox.showerror("Load workbook", str(exc))
            self.status_text.set(f"Workbook load failed: {exc}")
            return

        self.last_probe_result = None
        self._clear_results_panel()
        self._populate_form_dropdowns()
        self._populate_overview()
        self._populate_admin()
        self.notebook.select(self.overview_tab)
        staging_note = self.diagnosis.debug.get("staging_note")
        if staging_note:
            self.status_text.set(f"Workbook loaded from staged copy: {Path(self.workbook_path.get()).name}")
        else:
            self.status_text.set(f"Workbook loaded: {Path(self.workbook_path.get()).name}")
        self._hide_loading_overlay()

    def _clear_scenario(self) -> None:
        defaults = {
            "security_type": "First Lien",
            "loan_denomination": "USD",
            "country": "United States",
            "rate_type": "Floating",
            "payment_frequency": "M",
            "purchase_price": "100.00%",
            "pik_pct": "0.00%",
        }
        for key, var in self.form_vars.items():
            var.set(defaults.get(key, ""))
        self.last_probe_result = None
        self._clear_results_panel()
        self.status_text.set("Scenario cleared. No workbook writes were made.")

    def _current_snapshot(self) -> dict | None:
        if self.last_probe_result and self.last_probe_result.get("status") == "ok":
            snapshot = dict(self.last_probe_result.get("metrics", {}))
            snapshot["concentration_limits"] = self.last_probe_result.get("concentration_limits", [])
            return snapshot
        return None

    def _clear_results_panel(self) -> None:
        for tree in [
            self.issue_tree,
            self.availability_tree,
            self.portfolio_tree,
            self.concentration_tree,
            self.model_metrics_tree,
        ]:
            self._clear_tree(tree)
        self.results_banner.set("Run validation, current model read, or pro forma to populate this report.")
        self._set_text_content(self.commentary_text, "")

    def _format_delta(self, before, after, formatter) -> str:
        if before in (None, "") or after in (None, ""):
            return ""
        try:
            return formatter((after or 0) - (before or 0))
        except (TypeError, ValueError):
            return ""

    def _populate_results_view(
        self,
        values: dict[str, str],
        commentary: str,
        snapshot: dict | None,
        *,
        after_snapshot: dict | None = None,
        eligibility: dict | None = None,
        banner: str | None = None,
    ) -> None:
        try:
            ebitda = float(values.get("ltm_adj_ebitda", "0") or 0)
            revolver = float(values.get("drawn_revolver", "0") or 0)
            first_out = float(values.get("first_out_balance", "0") or 0)
            pari = float(values.get("pari_passu", "0") or 0)
            total_sm_raw = float(values.get("total_sm_balance", "0") or 0)
            cash = float(values.get("cash_balance", "0") or 0)
            total_net = _fmt_multiple((revolver + first_out + pari + total_sm_raw - cash) / ebitda) if ebitda > 0 else ""
            net_attachment = _fmt_multiple((revolver + first_out - cash) / ebitda) if ebitda > 0 else ""
        except ValueError:
            total_sm_raw = 0
            total_net = ""
            net_attachment = ""

        try:
            share_pct = _fmt_percent(float(values.get("bdc_balance", "0") or 0) / total_sm_raw) if total_sm_raw > 0 else ""
        except ValueError:
            share_pct = ""

        self._clear_tree(self.availability_tree)
        self._clear_tree(self.portfolio_tree)
        self._clear_tree(self.concentration_tree)
        self._clear_tree(self.model_metrics_tree)

        self.results_banner.set(
            banner
            or (
                f"Scenario for {values.get('company_name', '').strip() or 'unnamed company'}"
                if values.get("company_name")
                else "Scenario results"
            )
        )

        before = snapshot or {}
        after = after_snapshot or {}

        availability_rows = [
            ("BDC Investment Amount", "", _fmt_currency(values.get("bdc_balance")), "", "Scenario input"),
            (
                "Availability",
                _fmt_currency(before.get("availability")),
                _fmt_currency(after.get("availability")) if after_snapshot else "",
                self._format_delta(before.get("availability"), after.get("availability"), _fmt_currency) if after_snapshot else "",
                "Borrowing base availability",
            ),
            (
                "Aggregate Adjusted Borrowing Value",
                _fmt_currency(before.get("aggregate_adjusted_bv")),
                _fmt_currency(after.get("aggregate_adjusted_bv")) if after_snapshot else "",
                self._format_delta(before.get("aggregate_adjusted_bv"), after.get("aggregate_adjusted_bv"), _fmt_currency) if after_snapshot else "",
                "Availability tab adjusted collateral value",
            ),
            (
                "Excess Concentration Haircut",
                _fmt_currency(before.get("excess_concentration")),
                _fmt_currency(after.get("excess_concentration")) if after_snapshot else "",
                self._format_delta(before.get("excess_concentration"), after.get("excess_concentration"), _fmt_currency) if after_snapshot else "",
                "Concentration haircut applied by the model",
            ),
            (
                "Net Adjusted Borrowing Value",
                _fmt_currency(before.get("net_adjusted_bv")),
                _fmt_currency(after.get("net_adjusted_bv")) if after_snapshot else "",
                self._format_delta(before.get("net_adjusted_bv"), after.get("net_adjusted_bv"), _fmt_currency) if after_snapshot else "",
                "Adjusted borrowing value net of concentrations",
            ),
        ]
        for row in availability_rows:
            self.availability_tree.insert("", "end", values=row)

        portfolio_rows = [
            ("Total SM Balance", _fmt_currency(values.get("total_sm_balance")), "Scenario input"),
            ("BDC Share %", share_pct, "BDC Balance / Total SM Balance"),
            ("Total Net Leverage", total_net, "(Revolver + First Out + Pari + Total SM - Cash) / EBITDA"),
            ("Net Detachment", total_net, "All debt / EBITDA"),
            ("Net Attachment", net_attachment, "(Revolver + First Out - Cash) / EBITDA"),
        ]
        for row in portfolio_rows:
            self.portfolio_tree.insert("", "end", values=row)

        if before.get("concentration_limits"):
            after_lookup = {row.get("limit_type"): row for row in after.get("concentration_limits", [])} if after_snapshot else {}
            for row in before["concentration_limits"]:
                after_row = after_lookup.get(row.get("limit_type"), {})
                self.concentration_tree.insert(
                    "",
                    "end",
                    values=(
                        row.get("limit_type", ""),
                        _fmt_percent(row.get("limit_percent")),
                        _fmt_currency(row.get("applicable_limit")),
                        _fmt_currency(row.get("actual")),
                        _fmt_currency(row.get("excess")),
                        _fmt_currency(after_row.get("actual")) if after_row else "",
                        _fmt_currency(after_row.get("excess")) if after_row else "",
                        self._format_delta(row.get("actual"), after_row.get("actual"), _fmt_currency) if after_row else "",
                        self._format_delta(row.get("excess"), after_row.get("excess"), _fmt_currency) if after_row else "",
                    ),
                )
        else:
            self.concentration_tree.insert(
                "",
                "end",
                values=("Current concentration snapshot not loaded yet.", "", "", "", "", "", "", "", ""),
            )

        model_rows = [
            (
                "Weighted Average Advance Rate",
                _fmt_percent(before.get("weighted_avg_advance_rate")),
                _fmt_percent(after.get("weighted_avg_advance_rate")) if after_snapshot else "",
                "Availability!L48",
            ),
            (
                "Credit Enhancement Test",
                str(before.get("credit_enhancement_test", "")).strip(),
                str(after.get("credit_enhancement_test", "")).strip() if after_snapshot else "",
                "Availability!L42",
            ),
            (
                "Current Advances",
                _fmt_currency(before.get("current_advances")),
                _fmt_currency(after.get("current_advances")) if after_snapshot else "",
                "Availability!L51",
            ),
        ]
        for row in model_rows:
            self.model_metrics_tree.insert("", "end", values=row)

        commentary_lines = [commentary]
        if eligibility:
            commentary_lines.append("")
            commentary_lines.append(f"Eligibility status: {eligibility.get('status', '')}")
            failed = eligibility.get("failed_tests", [])
            if failed:
                commentary_lines.append(f"Failed tests: {', '.join(failed)}")
        self._set_text_content(self.commentary_text, "\n".join(commentary_lines).strip())

    def _show_results_message(self, title: str, message: str) -> None:
        self._clear_results_panel()
        self.results_banner.set(title)
        self._set_text_content(self.commentary_text, message)

    def _tree_to_markdown(self, tree: ttk.Treeview) -> str:
        columns = tree["columns"]
        if not columns:
            return ""
        lines = [" | ".join(columns), " | ".join("---" for _ in columns)]
        for item_id in tree.get_children():
            values = [str(value) if value is not None else "" for value in tree.item(item_id, "values")]
            if any(values):
                lines.append(" | ".join(values))
        return "\n".join(lines)

    def _export_results(self) -> None:
        if not any(tree.get_children() for tree in [self.issue_tree, self.availability_tree, self.portfolio_tree, self.concentration_tree, self.model_metrics_tree]):
            messagebox.showwarning("Export results", "Run validation, current model read, or pro forma first.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".md",
            initialfile="borrowing_base_results.md",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt")],
        )
        if not path:
            return

        sections = [
            "# Borrowing Base Results",
            "",
            f"## {self.results_banner.get()}",
            "",
            "### Scenario Checks",
            "",
            self._tree_to_markdown(self.issue_tree),
            "",
            "### Availability Impact",
            "",
            self._tree_to_markdown(self.availability_tree),
            "",
            "### Portfolio Composition",
            "",
            self._tree_to_markdown(self.portfolio_tree),
            "",
            "### Concentration Limits",
            "",
            self._tree_to_markdown(self.concentration_tree),
            "",
            "### Workbook Metrics",
            "",
            self._tree_to_markdown(self.model_metrics_tree),
            "",
            "### Commentary",
            "",
            self.commentary_text.get("1.0", "end").strip(),
        ]
        Path(path).write_text("\n".join(section for section in sections if section is not None), encoding="utf-8")
        self.status_text.set(f"Results exported to {Path(path).name}.")

    def _validate_scenario(self) -> None:
        if not self.diagnosis:
            messagebox.showwarning("Validate scenario", "Load a workbook first.")
            return

        values = self._collect_form_values()
        issues = validate_scenario(values, self.diagnosis)
        commentary = build_commentary(values, issues)

        self._clear_tree(self.issue_tree)
        for issue in issues:
            self.issue_tree.insert("", "end", values=(issue.severity, issue.field, issue.message))

        snapshot = self._current_snapshot()
        self._populate_results_view(
            values,
            commentary,
            snapshot,
            banner="Scenario validation complete. Review the governed checks and baseline metrics below.",
        )
        self.notebook.select(self.results_tab)
        self.status_text.set("Scenario validated. Results view refreshed.")

    def _run_pro_forma(self) -> None:
        if not self.diagnosis:
            messagebox.showwarning("Run pro forma", "Load a workbook first.")
            return
        values = self._collect_form_values()
        issues = validate_scenario(values, self.diagnosis)
        hard_stops = [issue for issue in issues if issue.severity == "Hard Stop"]
        if hard_stops:
            self._validate_scenario()
            summary = "\n".join(f"- {issue.field}: {issue.message}" for issue in hard_stops[:5])
            if len(hard_stops) > 5:
                summary += f"\n- And {len(hard_stops) - 5} more"
            messagebox.showwarning(
                "Run pro forma",
                "Scenario has hard-stop issues. Fix them before running pro forma.\n\n"
                f"{summary}\n\n"
                "The full list is also shown on the Results tab under Scenario Checks.",
            )
            return
        warnings = [issue for issue in issues if issue.severity != "Hard Stop"]
        if warnings:
            proceed = messagebox.askyesno("Run pro forma", "The scenario has warning-level flags. Do you want to proceed with the workbook-copy run?")
            if not proceed:
                return

        if not self._current_snapshot():
            self._probe_excel()
            if not self._current_snapshot():
                return

        script_path = Path(__file__).with_name("run_pro_forma.ps1")
        try:
            self._show_loading_overlay("Running pro forma", "Writing the scenario to a staged workbook copy, recalculating Excel, and collecting before / after results.")
            result = run_pro_forma_workbook(self.workbook_path.get(), script_path, values)
        finally:
            self._hide_loading_overlay()
        self._clear_tree(self.issue_tree)
        for issue in issues:
            self.issue_tree.insert("", "end", values=(issue.severity, issue.field, issue.message))

        if result.get("status") == "ok":
            before_snapshot = dict(result.get("before", {}))
            after_snapshot = dict(result.get("after", {}))
            commentary = build_commentary(values, issues)
            commentary += "\nThe pro forma was run on a staged workbook copy, so the master workbook was not changed."
            self._populate_results_view(
                values,
                commentary,
                before_snapshot,
                after_snapshot=after_snapshot,
                eligibility=result.get("eligibility", {}),
                banner="Pro forma completed. Compare the current borrowing base against the staged scenario below.",
            )
            self.notebook.select(self.results_tab)
            self.status_text.set("Pro forma completed on a staged workbook copy.")
        else:
            lines = ["Run Pro Forma failed.", "", result.get("message", "Unknown error")]
            self._show_results_message("Run Pro Forma failed.", "\n".join(lines))
            self.notebook.select(self.results_tab)
            self.status_text.set("Run pro forma failed.")

    def _probe_excel(self) -> None:
        script_path = Path(__file__).with_name("excel_probe.ps1")
        try:
            self._show_loading_overlay("Reading current model", "Opening the governed workbook in Excel and capturing the live baseline metrics.")
            result = probe_excel_workbook(self.workbook_path.get(), script_path)
        finally:
            self._hide_loading_overlay()
        self.last_probe_result = result

        if result.get("status") == "ok":
            snapshot = dict(result.get("metrics", {}))
            snapshot["concentration_limits"] = result.get("concentration_limits", [])
            values = self._collect_form_values()
            commentary = (
                "This is the current live-model snapshot. It reads the workbook as it stands today; "
                "it does not add a new investment, clear an old investment, or write back to the source file."
            )
            self._populate_results_view(
                values,
                commentary,
                snapshot,
                banner="Current model snapshot loaded. This is the live workbook baseline before any new scenario is written.",
            )
            self._populate_overview()
            self.status_text.set(f"Current model read successfully from {Path(self.workbook_path.get()).name}.")
        else:
            lines = ["Read Current Model failed.", "", result.get("message", "Unknown error")]
            if result.get("remediation"):
                lines.append("")
                lines.append("Suggested remediation:")
                lines.extend(f"- {item}" for item in result["remediation"])
            self._show_results_message("Read Current Model failed.", "\n".join(lines))
            self.status_text.set("Current model read failed.")
        self.notebook.select(self.results_tab)

    def _export_diagnosis(self) -> None:
        if not self.diagnosis:
            messagebox.showwarning("Export diagnosis", "Load a workbook first.")
            return

        initial_name = f"{Path(self.workbook_path.get()).stem}_diagnosis.md"
        path = filedialog.asksaveasfilename(
            defaultextension=".md",
            initialfile=initial_name,
            filetypes=[("Markdown", "*.md"), ("JSON", "*.json")],
        )
        if not path:
            return

        payload = diagnosis_to_json(self.diagnosis) if path.lower().endswith(".json") else diagnosis_to_markdown(self.diagnosis)
        Path(path).write_text(payload, encoding="utf-8")
        self.status_text.set(f"Diagnosis exported to {Path(path).name}.")
        messagebox.showinfo("Export diagnosis", f"Saved {path}")


def main() -> None:
    app = BorrowingBaseWorkbench()
    app.mainloop()


if __name__ == "__main__":
    main()
