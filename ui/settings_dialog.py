"""
Settings dialog -- configure signatory names/titles and commission rate.
All values are persisted to %APPDATA%/MPCScript/settings.json.
"""
import tkinter as tk
from tkinter import messagebox, ttk

import config.settings as settings_mod


class SettingsDialog(tk.Toplevel):
    """Modal dialog for editing application settings."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.title("Settings")
        self.resizable(False, False)
        self.grab_set()                 # modal

        self._data = settings_mod.load()
        self._vars: dict = {}

        self._build_ui()
        self._populate()

        # Centre over parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 6}

        # ---- Signatory frame ----
        sig_frame = ttk.LabelFrame(self, text="Signatories", padding=10)
        sig_frame.pack(fill="x", padx=14, pady=(14, 6))

        headers = ["", "Name", "Title / Position"]
        for col_idx, hdr in enumerate(headers):
            ttk.Label(sig_frame, text=hdr, font=("Segoe UI", 9, "bold")).grid(
                row=0, column=col_idx, **pad, sticky="w"
            )

        sigs = [
            ("Prepared By",  "prepared_by_name",  "prepared_by_title"),
            ("Checked By",   "checked_by_name",   "checked_by_title"),
            ("Validated By", "validated_by_name",  "validated_by_title"),
        ]
        for row_idx, (label, name_key, title_key) in enumerate(sigs, start=1):
            ttk.Label(sig_frame, text=label).grid(
                row=row_idx, column=0, **pad, sticky="w"
            )
            name_var = tk.StringVar()
            title_var = tk.StringVar()
            self._vars[name_key]  = name_var
            self._vars[title_key] = title_var

            ttk.Entry(sig_frame, textvariable=name_var, width=28).grid(
                row=row_idx, column=1, **pad, sticky="ew"
            )
            ttk.Entry(sig_frame, textvariable=title_var, width=32).grid(
                row=row_idx, column=2, **pad, sticky="ew"
            )
        sig_frame.columnconfigure(1, weight=1)
        sig_frame.columnconfigure(2, weight=1)

        # ---- Commission rate frame ----
        rate_frame = ttk.LabelFrame(self, text="Commission Rate", padding=10)
        rate_frame.pack(fill="x", padx=14, pady=6)

        ttk.Label(rate_frame, text="Club Commission Rate (%)").pack(side="left")
        comm_var = tk.StringVar()
        self._vars["commission_rate"] = comm_var
        ttk.Entry(rate_frame, textvariable=comm_var, width=8).pack(side="left", padx=8)
        ttk.Label(rate_frame, text="(e.g. enter 5 for 5%)").pack(side="left")

        # ---- Buttons ----
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=14, pady=(6, 14))
        ttk.Button(btn_frame, text="Save", command=self._save).pack(side="right", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="right")

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------

    def _populate(self) -> None:
        for key, var in self._vars.items():
            if key == "commission_rate":
                var.set(str(self._data.get("commission_rate", 5.0)))
            else:
                var.set(self._data.get(key, ""))

    def _save(self) -> None:
        new_settings: dict = {}
        for key, var in self._vars.items():
            if key == "commission_rate":
                try:
                    new_settings[key] = float(var.get().strip())
                except ValueError:
                    messagebox.showerror(
                        "Invalid Input",
                        "Commission Rate must be a number (e.g. 5 for 5%).",
                        parent=self,
                    )
                    return
            else:
                new_settings[key] = var.get().strip()

        settings_mod.save(new_settings)
        self.destroy()
