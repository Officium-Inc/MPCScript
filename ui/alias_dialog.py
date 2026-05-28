"""
Dialog for mapping unmatched POS instructor names to ISP list entries.

Shown automatically after Generate Summary when one or more instructors
are not found in the ISP list.  The user selects the correct ISP entry
from a dropdown (or leaves it as "-- No match --" to keep EWT = 0).
Selections are remembered and auto-applied on future runs.
"""
import difflib
import tkinter as tk
from tkinter import ttk

_NO_MATCH = "-- No match  (EWT = 0) --"

_NAVY   = "#1E3A5F"
_ORANGE = "#E8884C"
_WHITE  = "#FFFFFF"
_BG     = "#F0F2F5"

_MAX_VISIBLE_ROWS = 12   # rows shown before scrollbar activates
_ROW_HEIGHT       = 30   # approximate px per row


class AliasDialog(tk.Toplevel):
    """
    Parameters
    ----------
    parent          : tk root/toplevel
    unmatched       : list of POS item names that weren't found in the ISP dict
    isp_names       : list of ISP display names (from isp_dict values)
    existing_aliases: previously saved {pos_name_lower: isp_display_name | None}
    """

    def __init__(
        self,
        parent: tk.Misc,
        unmatched: list,
        isp_names: list,
        existing_aliases: dict,
    ) -> None:
        super().__init__(parent)
        self.title("ISP Name Mapping")
        self.resizable(True, True)
        self.grab_set()
        self.configure(bg=_BG)

        self.result: dict = {}
        self._vars: dict  = {}

        sorted_isp = sorted(isp_names, key=str.casefold)
        choices    = [_NO_MATCH] + sorted_isp

        # ── title bar ────────────────────────────────────────────────
        title_bar = tk.Frame(self, bg=_NAVY, height=40)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(
            title_bar,
            text="ISP Name Mapping",
            bg=_NAVY, fg=_WHITE,
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=12, pady=8)

        # ── instruction ───────────────────────────────────────────────
        tk.Label(
            self,
            text=(
                "The instructors below were not found in the ISP list.\n"
                "Select the matching ISP entry for each, or leave as\n"
                "\"No match\" to keep EWT = 0 for that instructor."
            ),
            bg=_BG, fg=_NAVY,
            font=("Segoe UI", 9),
            justify="left",
        ).pack(anchor="w", padx=14, pady=(10, 4))

        # ── column headers (fixed, outside scroll) ────────────────────
        hdr_frame = tk.Frame(self, bg=_BG)
        hdr_frame.pack(fill="x", padx=14, pady=(0, 2))
        tk.Label(hdr_frame, text="POS Instructor Name",
                 bg=_BG, fg=_NAVY,
                 font=("Segoe UI", 9, "bold"), width=30, anchor="w").pack(side="left")
        tk.Label(hdr_frame, text="ISP List Entry",
                 bg=_BG, fg=_NAVY,
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left", padx=(8, 0))

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=(0, 4))

        # ── scrollable area ───────────────────────────────────────────
        scroll_height = min(len(unmatched), _MAX_VISIBLE_ROWS) * _ROW_HEIGHT

        scroll_container = tk.Frame(self, bg=_BG)
        scroll_container.pack(fill="both", expand=True, padx=10, pady=0)

        canvas = tk.Canvas(
            scroll_container, bg=_BG,
            highlightthickness=0,
            height=scroll_height,
        )
        scrollbar = ttk.Scrollbar(
            scroll_container, orient="vertical", command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=_BG)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_canvas_resize(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        def _on_inner_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_inner_resize)

        def _on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                # Canvas was destroyed (dialog closed); remove the stale binding
                self.unbind_all("<MouseWheel>")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        inner.bind("<MouseWheel>", _on_mousewheel)
        self.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind("<Destroy>", lambda e: self.unbind_all("<MouseWheel>") if e.widget is self else None)

        # ── mapping rows ──────────────────────────────────────────────
        for i, pos_name in enumerate(unmatched):
            row_bg = _BG if i % 2 == 0 else "#E8EBF0"
            row_frame = tk.Frame(inner, bg=row_bg)
            row_frame.pack(fill="x", pady=1)
            row_frame.bind("<MouseWheel>", _on_mousewheel)

            tk.Label(
                row_frame,
                text=pos_name,
                bg=row_bg, fg="#333333",
                font=("Segoe UI", 9),
                width=30, anchor="w",
            ).pack(side="left", padx=(4, 0), pady=3)

            var = tk.StringVar()

            key = pos_name.lower()
            if key in existing_aliases and existing_aliases[key] is not None:
                saved = existing_aliases[key]
                var.set(saved if saved in sorted_isp else _NO_MATCH)
            else:
                isp_lower = [n.casefold() for n in sorted_isp]
                matches   = difflib.get_close_matches(
                    key, isp_lower, n=1, cutoff=0.4
                )
                var.set(
                    sorted_isp[isp_lower.index(matches[0])]
                    if matches else _NO_MATCH
                )

            cb = ttk.Combobox(
                row_frame, textvariable=var,
                values=choices, width=36, state="readonly",
            )
            cb.pack(side="left", padx=(8, 4), pady=3)
            cb.bind("<MouseWheel>", _on_mousewheel)
            self._vars[pos_name] = var

        # ── buttons ───────────────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=(4, 0))

        btn_frame = tk.Frame(self, bg=_BG)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="  Apply & Regenerate  ",
            bg=_ORANGE, fg=_WHITE,
            activebackground="#c96e38", activeforeground=_WHITE,
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            command=self._apply,
        ).pack(side="left", padx=6)

        tk.Button(
            btn_frame,
            text="  Skip  ",
            bg="#CCCCCC", fg="#333333",
            activebackground="#AAAAAA",
            relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 9),
            command=self.destroy,
        ).pack(side="left", padx=6)

        self._center(parent)

    def _apply(self) -> None:
        for pos_name, var in self._vars.items():
            val = var.get()
            self.result[pos_name.lower()] = None if val == _NO_MATCH else val
        self.destroy()

    def _center(self, parent: tk.Misc) -> None:
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        dw, dh = self.winfo_width(), self.winfo_height()
        px = parent.winfo_x() + max((pw - dw) // 2, 0)
        py = parent.winfo_y() + max((ph - dh) // 2, 0)
        self.geometry(f"+{px}+{py}")

