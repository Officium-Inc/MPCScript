"""
Main application window.

Layout:
  ┌──────────────────────────────────────────────────────────────┐
  │  Manila Polo Club — Instructor Fee Summary      [⚙ Settings] │
  ├──────────────────┬───────────────────────────────────────────┤
  │  STEP 1          │  [ARCHERY] [GOLF DRIVING RANGE]           │
  │  ISP File        │  Item | Unit | Qty | Amt | Comm | ...     │
  │  [Browse]        │  Florante Matan | ...                     │
  │  ✓ Loaded        │  Joan Tabanag   | ...                     │
  │                  │  Grand Total    | ...                     │
  │  STEP 2          │                                           │
  │  Raw Data        │                                           │
  │  [Browse]        │                                           │
  │  ✓ Loaded        │                                           │
  │                  │                                           │
  │ [Generate]       │                                           │
  │                  │                                           │
  │ [Export PDF]     │                                           │
  │ [Export Excel]   │                                           │
  ├──────────────────┴───────────────────────────────────────────┤
  │ Status: Ready                                                │
  └──────────────────────────────────────────────────────────────┘
"""
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import config.settings as settings_mod
import core.calculator as calculator
import core.excel_exporter as excel_exporter
import core.isp_reader as isp_reader
import core.pdf_exporter as pdf_exporter
import core.raw_data_reader as raw_data_reader
from ui.settings_dialog import SettingsDialog
from ui.alias_dialog import AliasDialog

# ---------------------------------------------------------------------------
# Colour / style constants
# ---------------------------------------------------------------------------
COLOR_NAVY   = "#1E3A5F"
COLOR_ORANGE = "#E8884C"
COLOR_WHITE  = "#FFFFFF"
COLOR_BG     = "#F0F2F5"
COLOR_PANEL  = "#FFFFFF"
COLOR_GREEN  = "#2E7D32"
COLOR_GREY   = "#9E9E9E"
COLOR_GT_BG  = "#FFF3E0"
COLOR_ALT    = "#FFF8F4"

TREE_COLUMNS = [
    ("item",          "Instructor",            200, "w"),
    ("min_unit_price","Min. Unit Price",         100, "e"),
    ("sum_qty",       "Qty",                     55, "e"),
    ("sum_amt",       "Amt. After Disc.",        120, "e"),
    ("comm",          "Commission",              100, "e"),
    ("total",         "Total",                   100, "e"),
    ("ewt_rate",      "EWT Rate",                 75, "e"),
    ("ewt",           "EWT",                      90, "e"),
    ("final_total",   "Net Total",               100, "e"),
]


class MainWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Manila Polo Club — Instructor Fee Summary")
        self.geometry("1150x680")
        self.minsize(900, 560)
        self.configure(bg=COLOR_BG)

        self._isp_path: str = ""
        self._raw_path: str = ""
        self._isp_dict: dict = {}
        self._summaries: dict = {}
        self._aliases: dict = settings_mod.load_aliases()
        self._settings: dict = settings_mod.load()

        self._apply_styles()
        self._build_header()
        self._build_body()
        self._build_statusbar()

    # ------------------------------------------------------------------
    # ttk styling
    # ------------------------------------------------------------------

    def _apply_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=COLOR_BG)
        style.configure("Panel.TFrame", background=COLOR_PANEL)
        style.configure(
            "Accent.TButton",
            background=COLOR_ORANGE,
            foreground=COLOR_WHITE,
            font=("Segoe UI", 9, "bold"),
            padding=(10, 6),
        )
        style.map("Accent.TButton", background=[("active", "#d4713a")])
        style.configure(
            "Primary.TButton",
            background=COLOR_NAVY,
            foreground=COLOR_WHITE,
            font=("Segoe UI", 9, "bold"),
            padding=(10, 6),
        )
        style.map("Primary.TButton", background=[("active", "#162c49")])
        style.configure(
            "Flat.TButton",
            background=COLOR_PANEL,
            foreground=COLOR_NAVY,
            font=("Segoe UI", 9),
            padding=(8, 5),
        )
        style.map("Flat.TButton", background=[("active", "#e8eaf0")])
        style.configure("TLabel", background=COLOR_PANEL, font=("Segoe UI", 9))
        style.configure("Header.TLabel", background=COLOR_NAVY, foreground=COLOR_WHITE,
                        font=("Segoe UI", 12, "bold"), padding=(12, 8))
        style.configure("Step.TLabel", background=COLOR_PANEL,
                        font=("Segoe UI", 9, "bold"), foreground=COLOR_NAVY)
        style.configure("Status.TLabel", background=COLOR_NAVY,
                        foreground=COLOR_WHITE, font=("Segoe UI", 8), padding=(6, 3))
        style.configure("OK.TLabel", background=COLOR_PANEL,
                        foreground=COLOR_GREEN, font=("Segoe UI", 9))
        style.configure("Warn.TLabel", background=COLOR_PANEL,
                        foreground="#B71C1C", font=("Segoe UI", 9))
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                        background=COLOR_ORANGE, foreground=COLOR_WHITE)
        style.map("Treeview", background=[("selected", COLOR_ORANGE)])

    # ------------------------------------------------------------------
    # Header bar
    # ------------------------------------------------------------------

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=COLOR_NAVY, height=50)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="Manila Polo Club  —  Instructor Fee Summary",
            bg=COLOR_NAVY, fg=COLOR_WHITE,
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left", padx=16, pady=10)

        tk.Button(
            header,
            text="⚙  Settings",
            bg=COLOR_NAVY, fg=COLOR_WHITE,
            activebackground="#162c49", activeforeground=COLOR_WHITE,
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9),
            command=self._open_settings,
        ).pack(side="right", padx=12, pady=10)

    # ------------------------------------------------------------------
    # Body (left panel + right notebook)
    # ------------------------------------------------------------------

    def _build_body(self) -> None:
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=0, pady=0)

        # Left control panel
        left = tk.Frame(body, bg=COLOR_PANEL, width=240)
        left.pack(side="left", fill="y", padx=0, pady=0)
        left.pack_propagate(False)
        self._build_left_panel(left)

        # Separator
        ttk.Separator(body, orient="vertical").pack(side="left", fill="y")

        # Right notebook area
        right = ttk.Frame(body)
        right.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        self._build_right_panel(right)

    def _build_left_panel(self, parent: tk.Frame) -> None:
        pad = {"padx": 14, "pady": 4}

        tk.Frame(parent, bg=COLOR_PANEL, height=12).pack()  # top spacer

        # ---- Step 1: ISP File ----
        ttk.Label(parent, text="Step 1 — ISP Reference List",
                  style="Step.TLabel").pack(anchor="w", **pad)

        self._isp_status = ttk.Label(parent, text="No file loaded", style="Warn.TLabel")
        self._isp_status.pack(anchor="w", padx=14, pady=0)

        ttk.Button(
            parent, text="Browse ISP File…",
            style="Flat.TButton", command=self._browse_isp,
        ).pack(anchor="w", **pad)

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # ---- Step 2: Raw Data ----
        ttk.Label(parent, text="Step 2 — POS Raw Data File",
                  style="Step.TLabel").pack(anchor="w", **pad)

        self._raw_status = ttk.Label(parent, text="Load ISP file first", style="Warn.TLabel")
        self._raw_status.pack(anchor="w", padx=14, pady=0)

        self._raw_btn = ttk.Button(
            parent, text="Browse Raw Data…",
            style="Flat.TButton", command=self._browse_raw,
            state="disabled",
        )
        self._raw_btn.pack(anchor="w", **pad)

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # ---- Generate ----
        self._gen_btn = ttk.Button(
            parent, text="▶  Generate Summary",
            style="Primary.TButton", command=self._generate,
            state="disabled",
        )
        self._gen_btn.pack(fill="x", padx=14, pady=4)

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # ---- Export ----
        ttk.Label(parent, text="Export", style="Step.TLabel").pack(anchor="w", **pad)

        self._pdf_btn = ttk.Button(
            parent, text="⬇  Export PDF",
            style="Accent.TButton", command=self._export_pdf,
            state="disabled",
        )
        self._pdf_btn.pack(fill="x", padx=14, pady=4)

        self._xlsx_btn = ttk.Button(
            parent, text="⬇  Export Excel",
            style="Accent.TButton", command=self._export_excel,
            state="disabled",
        )
        self._xlsx_btn.pack(fill="x", padx=14, pady=4)

    def _build_right_panel(self, parent: ttk.Frame) -> None:
        # Placeholder label shown before generation
        self._placeholder = tk.Label(
            parent,
            text=(
                "Load the ISP file and raw data,\n"
                "then click  ▶ Generate Summary."
            ),
            bg=COLOR_BG,
            fg=COLOR_GREY,
            font=("Segoe UI", 13),
            justify="center",
        )
        self._placeholder.pack(expand=True)

        # Notebook (hidden until generation)
        self._notebook = ttk.Notebook(parent)

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _build_statusbar(self) -> None:
        bar = tk.Frame(self, bg=COLOR_NAVY, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self._status_var = tk.StringVar(value="Ready")
        tk.Label(
            bar,
            textvariable=self._status_var,
            bg=COLOR_NAVY, fg=COLOR_WHITE,
            font=("Segoe UI", 8),
            anchor="w",
        ).pack(side="left", padx=10, fill="y")

    def _set_status(self, msg: str) -> None:
        self._status_var.set(msg)
        self.update_idletasks()

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self)
        self.wait_window(dlg)
        self._settings = settings_mod.load()

    def _apply_isp_dict(self, new_isp_dict: dict, path: str) -> None:
        """Update the stored ISP dict and clear aliases if the name list changed."""
        old_names = {v["display_name"].lower() for v in self._isp_dict.values()}
        new_names = {v["display_name"].lower() for v in new_isp_dict.values()}
        if old_names != new_names:
            self._aliases = {}
            settings_mod.save_aliases({})
        self._isp_dict = new_isp_dict
        self._isp_path = path

    def _browse_isp(self) -> None:
        path = filedialog.askopenfilename(
            title="Select ISP Reference List",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            self._set_status("Loading ISP list…")
            new_isp_dict = isp_reader.load(path)
            self._apply_isp_dict(new_isp_dict, path)
            fname = Path(path).name
            self._isp_status.configure(
                text=f"✓  {fname}  ({len(self._isp_dict)} instructors)",
                style="OK.TLabel",
            )
            self._raw_btn.configure(state="normal")
            self._raw_status.configure(text="No file loaded", style="Warn.TLabel")
            self._set_status(f"ISP list loaded: {len(self._isp_dict)} record(s).")
        except Exception as exc:
            messagebox.showerror("ISP File Error", str(exc))
            self._set_status("Error loading ISP file.")

    def _browse_raw(self) -> None:
        if not self._isp_dict:
            messagebox.showwarning("ISP Required", "Please load the ISP file first.")
            return
        path = filedialog.askopenfilename(
            title="Select POS Raw Data File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            self._set_status("Reading raw data…")
            df, warns = raw_data_reader.read(path)
            self._raw_df   = df
            self._raw_path = path
            fname = Path(path).name
            self._raw_status.configure(
                text=f"✓  {fname}  ({len(df)} rows)",
                style="OK.TLabel",
            )
            self._gen_btn.configure(state="normal")
            msg = f"Raw data loaded: {len(df)} transaction(s)."
            if warns:
                msg += "  Warnings: " + " | ".join(warns)
            self._set_status(msg)
        except Exception as exc:
            messagebox.showerror("Raw Data Error", str(exc))
            self._set_status("Error loading raw data file.")

    def _generate(self) -> None:
        if not self._isp_path:
            messagebox.showwarning("Missing Data", "Please load the ISP file first.")
            return
        if not hasattr(self, "_raw_df"):
            messagebox.showwarning("Missing Data", "Please load the raw data file first.")
            return
        try:
            # Always reload ISP list from disk so any edits to the Excel file
            # (tax rates, new instructors, etc.) are picked up immediately.
            self._set_status("Reloading ISP list from disk…")
            new_isp_dict = isp_reader.load(self._isp_path)
            self._apply_isp_dict(new_isp_dict, self._isp_path)
            fname = Path(self._isp_path).name
            self._isp_status.configure(
                text=f"✓  {fname}  ({len(self._isp_dict)} instructors)",
                style="OK.TLabel",
            )

            self._set_status("Generating summaries…")
            self._settings = settings_mod.load()
            commission_rate = self._settings.get("commission_rate", 5.0) / 100.0
            summaries, warns = calculator.summarise(
                self._raw_df, self._isp_dict, commission_rate, self._aliases
            )

            # --- ISP alias resolution ---
            # Collect names still unmatched after applying saved aliases
            import re as _re
            _pat = _re.compile(r"^'(.+)' was not found")
            unmatched = [
                m.group(1) for w in warns
                for m in [_pat.match(w)] if m
                and m.group(1).lower() not in self._aliases
            ]
            if unmatched:
                isp_names = [
                    v["display_name"] for v in self._isp_dict.values()
                ]
                dlg = AliasDialog(self, unmatched, isp_names, self._aliases)
                self.wait_window(dlg)
                if dlg.result:
                    self._aliases.update(dlg.result)
                    settings_mod.save_aliases(self._aliases)
                    # Re-run with the new aliases applied
                    summaries, warns = calculator.summarise(
                        self._raw_df, self._isp_dict,
                        commission_rate, self._aliases
                    )

            self._summaries = summaries
            self._populate_notebook(summaries)
            self._pdf_btn.configure(state="normal")
            self._xlsx_btn.configure(state="normal")

            # Remaining warnings (truly unmapped names)
            still_unmatched = [
                w for w in warns if "was not found" in w
            ]
            msg = f"Summary generated for {len(summaries)} group(s)."
            if still_unmatched:
                msg += "  Some instructors have no ISP match (EWT = 0)."
            self._set_status(msg)
        except Exception as exc:
            messagebox.showerror("Generation Error", str(exc))
            self._set_status("Error generating summary.")

    def _export_pdf(self) -> None:
        if not self._summaries:
            return
        path = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile="Instructor_Fee_Summary.pdf",
        )
        if not path:
            return
        try:
            self._set_status("Exporting PDF…")
            pdf_exporter.export(self._summaries, self._settings, path)
            self._set_status(f"PDF saved: {path}")
            messagebox.showinfo("Export Complete", f"PDF saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("PDF Export Error", str(exc))
            self._set_status("Error exporting PDF.")

    def _export_excel(self) -> None:
        if not self._summaries:
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel Report",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="Instructor_Fee_Summary.xlsx",
        )
        if not path:
            return
        try:
            self._set_status("Exporting Excel…")
            excel_exporter.export(self._summaries, self._settings, path)
            self._set_status(f"Excel saved: {path}")
            messagebox.showinfo("Export Complete", f"Excel saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Excel Export Error", str(exc))
            self._set_status("Error exporting Excel.")

    # ------------------------------------------------------------------
    # Notebook / Treeview population
    # ------------------------------------------------------------------

    def _populate_notebook(self, summaries: dict) -> None:
        # Hide placeholder, show notebook
        self._placeholder.pack_forget()
        # Destroy old tabs
        for tab in self._notebook.tabs():
            self._notebook.forget(tab)
        self._notebook.pack(fill="both", expand=True)

        for group_name, data in summaries.items():
            frame = ttk.Frame(self._notebook)
            from core.calculator import tab_label
            self._notebook.add(frame, text=tab_label(group_name))
            self._build_treeview(frame, data)

    def _build_treeview(self, parent: ttk.Frame, data: dict) -> None:
        from core.calculator import format_date_range, group_to_title

        # Sub-header with date range
        date_range = format_date_range(data["date_min"], data["date_max"])
        tk.Label(
            parent,
            text=date_range,
            bg=COLOR_BG, fg=COLOR_GREY,
            font=("Segoe UI", 9),
            anchor="e",
        ).pack(fill="x", padx=8, pady=(4, 0))

        # Scrollbar + Treeview
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True, padx=8, pady=6)

        vsb = ttk.Scrollbar(container, orient="vertical")
        hsb = ttk.Scrollbar(container, orient="horizontal")

        col_ids = [c[0] for c in TREE_COLUMNS]
        tree = ttk.Treeview(
            container,
            columns=col_ids,
            show="headings",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
        )
        vsb.configure(command=tree.yview)
        hsb.configure(command=tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        # Configure columns
        for col_id, heading, width, anchor in TREE_COLUMNS:
            tree.heading(col_id, text=heading)
            tree.column(col_id, width=width, anchor=anchor, minwidth=50)

        # Tags
        tree.tag_configure("alt",         background=COLOR_ALT)
        tree.tag_configure("grand_total", background=COLOR_GT_BG,
                           font=("Segoe UI", 9, "bold"))

        # Data rows
        rows_df = data["rows"]
        for idx, (_, row) in enumerate(rows_df.iterrows()):
            tag = "alt" if idx % 2 == 1 else ""
            tree.insert(
                "", "end",
                values=_format_row(row, is_total=False),
                tags=(tag,),
            )

        # Grand total row
        gt = data["grand_total"]
        tree.insert(
            "", "end",
            values=_format_grand_total(gt),
            tags=("grand_total",),
        )


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _fmt(value) -> str:
    if value is None or value == "":
        return ""
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value)


def _fmt_qty(value) -> str:
    """1.0 → '1',  0.5 → '0.5',  9.5 → '9.5'"""
    if value is None:
        return ""
    try:
        q = float(value)
        return str(int(q)) if q == int(q) else f"{q:g}"
    except (TypeError, ValueError):
        return str(value)


def _format_row(row, is_total: bool = False) -> tuple:
    # Show EWT rate as a percentage string (e.g. "5.00%")
    ewt_rate = row.get("ewt_rate") if hasattr(row, "get") else row["ewt_rate"]
    rate_str = f"{float(ewt_rate) * 100:.2f}%" if ewt_rate not in (None, "") else ""
    return (
        row["item"],
        _fmt(row["min_unit_price"]),
        _fmt_qty(row["sum_qty"]),
        _fmt(row["sum_amt"]),
        _fmt(row["comm"]),
        _fmt(row["total"]),
        rate_str,
        _fmt(row["ewt"]),
        _fmt(row["final_total"]),
    )


def _format_grand_total(gt: dict) -> tuple:
    # Grand Total row has no single EWT rate — leave blank
    return (
        gt["item"],
        "",
        _fmt_qty(gt["sum_qty"]),
        _fmt(gt["sum_amt"]),
        _fmt(gt["comm"]),
        _fmt(gt["total"]),
        "",                     # EWT Rate — blank for grand total
        _fmt(gt["ewt"]),
        _fmt(gt["final_total"]),
    )
