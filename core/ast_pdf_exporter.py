"""
PDF exporter for AST (After School Tennis) fee reports.

One page per group (TRAINER or BALLBOY).

TRAINER table columns:
  Name | Hours | Rate | Total Amount | VAT (12%) | Ex-VAT | 5% Comm | Net Amount | EWT | Final Net

BALLBOY table columns:
  Name | Hours | Rate | Total Amount | VAT (12%) | Net Total
"""
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ── Colours (match POS report palette) ───────────────────────────────────────
C_HEADER_BG = colors.HexColor("#E8884C")
C_HEADER_FG = colors.white
C_TITLE_BG  = colors.HexColor("#C6EFCE")
C_GT_BG     = colors.HexColor("#FFF3E0")
C_BORDER    = colors.HexColor("#B0B0B0")
C_ROW_ALT   = colors.HexColor("#FFF8F4")
C_SIG_BG    = colors.HexColor("#E8F5E9")

# Landscape A4 usable width ≈ 756 pt  (842 - 2×43 pt margins)
_PAGE_W = 756

# ── TRAINER column layout (widths computed dynamically) ──────────────────────
# Fixed columns (no name, no day cols):
_T_HRS_W   = 40   # Total Hrs
_T_RATE_W  = 48   # Rate
_T_AMT_W   = 62   # Total Amount
_T_VAT_W   = 57   # VAT (12%)
_T_EXVAT_W = 62   # Ex-VAT
_T_COMM_W  = 53   # 5% Comm
_T_NET_W   = 62   # Net Amount
_T_EWT_W   = 48   # EWT
_T_FINAL_W = 62   # Final Net
_T_DAY_W   = 36   # width per day column
_T_FIXED_SUM = (
    _T_HRS_W + _T_RATE_W + _T_AMT_W + _T_VAT_W +
    _T_EXVAT_W + _T_COMM_W + _T_NET_W + _T_EWT_W + _T_FINAL_W
)  # 494

def _trainer_col_widths(n_days: int) -> list:
    name_w = max(100, _PAGE_W - _T_FIXED_SUM - n_days * _T_DAY_W)
    return (
        [name_w]
        + [_T_DAY_W] * n_days
        + [_T_HRS_W, _T_RATE_W, _T_AMT_W, _T_VAT_W,
           _T_EXVAT_W, _T_COMM_W, _T_NET_W, _T_EWT_W, _T_FINAL_W]
    )

def _trainer_hdr_labels(day_labels: list) -> list:
    day_hdrs = [lbl.replace(' ', '\n') for lbl in day_labels]
    return (
        ["Name"]
        + day_hdrs
        + ["Total\nHrs", "Rate", "Total\nAmount", "VAT\n(12%)",
           "Ex-VAT", "5%\nComm", "Net\nAmount", "EWT", "Final\nNet"]
    )

# ── BALLBOY column layout ─────────────────────────────────────────────────────
_B_HRS_W   = 42   # Total Hrs
_B_RATE_W  = 60   # Rate
_B_AMT_W   = 90   # Total Amount
_B_VAT_W   = 80   # VAT (12%)
_B_NET_W   = 90   # Net Total
_B_DAY_W   = 40   # width per day column
_B_FIXED_SUM = _B_HRS_W + _B_RATE_W + _B_AMT_W + _B_VAT_W + _B_NET_W  # 362

def _ballboy_col_widths(n_days: int) -> list:
    name_w = max(120, _PAGE_W - _B_FIXED_SUM - n_days * _B_DAY_W)
    return (
        [name_w]
        + [_B_DAY_W] * n_days
        + [_B_HRS_W, _B_RATE_W, _B_AMT_W, _B_VAT_W, _B_NET_W]
    )

def _ballboy_hdr_labels(day_labels: list) -> list:
    day_hdrs = [lbl.replace(' ', '\n') for lbl in day_labels]
    return (
        ["Name"]
        + day_hdrs
        + ["Total\nHrs", "Rate", "Total\nAmount", "VAT\n(12%)", "Net Total"]
    )


# ── Public API ────────────────────────────────────────────────────────────────

def export(summaries: dict, settings: dict, output_path: str) -> None:
    """
    Write a multi-page AST PDF to output_path.

    Parameters
    ----------
    summaries  : result from ast_reader.load()
    settings   : dict from config.settings.load()
    output_path: destination file path
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    styles = getSampleStyleSheet()

    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    story = []
    keys = list(summaries.keys())
    for i, label in enumerate(keys):
        data = summaries[label]
        _build_page(story, label, data, settings, styles)
        if i < len(keys) - 1:
            story.append(PageBreak())

    doc.build(story)


# ── Internal builders ─────────────────────────────────────────────────────────

def _build_page(story, label, data, settings, styles):
    grp_type  = data["type"]
    date_str  = data.get("date_str", "")
    day_cols  = data.get("day_cols", [])   # [(key, label), ...]
    day_labels = [lbl for (_, lbl) in day_cols]

    if grp_type == "TRAINER":
        col_widths = _trainer_col_widths(len(day_cols))
        table_data = _trainer_table_data(data["rows"], data["grand_total"], day_cols)
    else:
        col_widths = _ballboy_col_widths(len(day_cols))
        table_data = _ballboy_table_data(data["rows"], data["grand_total"], day_cols)

    total_w    = sum(col_widths)
    report_title = _make_title(label, grp_type)

    # ── Banner ──
    ts = ParagraphStyle("T", fontSize=11, fontName="Helvetica-Bold", alignment=TA_CENTER)
    ss = ParagraphStyle("S", fontSize=10, fontName="Helvetica-Bold", alignment=TA_CENTER)
    ds = ParagraphStyle("D", fontSize=9,  alignment=TA_RIGHT)

    banner = Table(
        [
            [Paragraph("Manila Polo Club Inc.", ts), ""],
            [Paragraph(report_title, ss), Paragraph(date_str, ds)],
        ],
        colWidths=[total_w * 0.65, total_w * 0.35],
    )
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_TITLE_BG),
        ("SPAN",       (0, 0), (1, 0)),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("BOX",        (0, 0), (-1, -1), 0.5, C_BORDER),
    ]))
    story.append(banner)
    story.append(Spacer(1, 0.3 * cm))

    # ── Summary table ──
    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(_table_style(len(table_data)))
    story.append(tbl)
    story.append(Spacer(1, 0.8 * cm))

    # ── Signatories ──
    story.append(_sig_table(settings, total_w))


def _trainer_table_data(df, grand_total: float, day_cols: list) -> list:
    ns = ParagraphStyle("N", alignment=TA_RIGHT,  fontSize=8)
    ls = ParagraphStyle("L", alignment=TA_LEFT,   fontSize=8)
    hs = ParagraphStyle("H", alignment=TA_CENTER, fontSize=8,
                        fontName="Helvetica-Bold", textColor=C_HEADER_FG)
    gs = ParagraphStyle("G", alignment=TA_LEFT,   fontSize=8, fontName="Helvetica-Bold")
    gn = ParagraphStyle("GN",alignment=TA_RIGHT,  fontSize=8, fontName="Helvetica-Bold")

    day_labels = [lbl for (_, lbl) in day_cols]
    rows = [[Paragraph(h, hs) for h in _trainer_hdr_labels(day_labels)]]

    for _, r in df.iterrows():
        day_vals = [Paragraph(_fq(r.get(key, 0)), ns) for (key, _) in day_cols]
        rows.append([
            Paragraph(str(r["name"]), ls),
            *day_vals,
            Paragraph(_fq(r["hours"]),         ns),
            Paragraph(_fm(r["rate"]),           ns),
            Paragraph(_fm(r["total_amount"]),   ns),
            Paragraph(_fm(r["vat"]),            ns),
            Paragraph(_fm(r["ex_vat"]),         ns),
            Paragraph(_fm(r["commission"]),     ns),
            Paragraph(_fm(r["net_amount"]),     ns),
            Paragraph(_fm(r["ewt"]),            ns),
            Paragraph(_fm(r["net_final"]),      ns),
        ])

    # Grand-total row
    gt_hours  = float(df["hours"].sum())      if not df.empty else 0.0
    gt_amount = float(df["total_amount"].sum()) if not df.empty else 0.0
    gt_vat    = float(df["vat"].sum())        if not df.empty else 0.0
    gt_exvat  = float(df["ex_vat"].sum())     if not df.empty else 0.0
    gt_comm   = float(df["commission"].sum()) if not df.empty else 0.0
    gt_net    = float(df["net_amount"].sum()) if not df.empty else 0.0
    gt_ewt    = float(df["ewt"].sum())        if not df.empty else 0.0
    day_sums  = [
        Paragraph(_fq(float(df[key].sum())) if key in df.columns and not df.empty else "", gn)
        for (key, _) in day_cols
    ]

    rows.append([
        Paragraph("GRAND TOTAL", gs),
        *day_sums,
        Paragraph(_fq(gt_hours),  gn),
        Paragraph("",             gn),
        Paragraph(_fm(gt_amount), gn),
        Paragraph(_fm(gt_vat),    gn),
        Paragraph(_fm(gt_exvat),  gn),
        Paragraph(_fm(gt_comm),   gn),
        Paragraph(_fm(gt_net),    gn),
        Paragraph(_fm(gt_ewt),    gn),
        Paragraph(_fm(grand_total), gn),
    ])
    return rows


def _ballboy_table_data(df, grand_total: float, day_cols: list) -> list:
    ns = ParagraphStyle("N", alignment=TA_RIGHT,  fontSize=8)
    ls = ParagraphStyle("L", alignment=TA_LEFT,   fontSize=8)
    hs = ParagraphStyle("H", alignment=TA_CENTER, fontSize=8,
                        fontName="Helvetica-Bold", textColor=C_HEADER_FG)
    gs = ParagraphStyle("G", alignment=TA_LEFT,   fontSize=8, fontName="Helvetica-Bold")
    gn = ParagraphStyle("GN",alignment=TA_RIGHT,  fontSize=8, fontName="Helvetica-Bold")

    day_labels = [lbl for (_, lbl) in day_cols]
    rows = [[Paragraph(h, hs) for h in _ballboy_hdr_labels(day_labels)]]

    for _, r in df.iterrows():
        day_vals = [Paragraph(_fq(r.get(key, 0)), ns) for (key, _) in day_cols]
        rows.append([
            Paragraph(str(r["name"]),    ls),
            *day_vals,
            Paragraph(_fq(r["hours"]),   ns),
            Paragraph(_fm(r["rate"]),    ns),
            Paragraph(_fm(r["total"]),   ns),
            Paragraph(_fm(r["vat"]),     ns),
            Paragraph(_fm(r["net_total"]), ns),
        ])

    gt_hours  = float(df["hours"].sum())     if not df.empty else 0.0
    gt_total  = float(df["total"].sum())     if not df.empty else 0.0
    gt_vat    = float(df["vat"].sum())       if not df.empty else 0.0
    day_sums  = [
        Paragraph(_fq(float(df[key].sum())) if key in df.columns and not df.empty else "", gn)
        for (key, _) in day_cols
    ]

    rows.append([
        Paragraph("GRAND TOTAL",     gs),
        *day_sums,
        Paragraph(_fq(gt_hours),     gn),
        Paragraph("",                gn),
        Paragraph(_fm(gt_total),     gn),
        Paragraph(_fm(gt_vat),       gn),
        Paragraph(_fm(grand_total),  gn),
    ])
    return rows


def _table_style(n_rows: int) -> TableStyle:
    cmds = [
        ("BACKGROUND",   (0, 0),      (-1, 0),         C_HEADER_BG),
        ("TEXTCOLOR",    (0, 0),      (-1, 0),         C_HEADER_FG),
        ("FONTNAME",     (0, 0),      (-1, 0),         "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0),      (-1, 0),         8),
        ("ALIGN",        (0, 0),      (-1, 0),         "CENTER"),
        ("VALIGN",       (0, 0),      (-1, -1),        "MIDDLE"),
        ("FONTSIZE",     (0, 1),      (-1, -1),        8),
        ("LEFTPADDING",  (0, 0),      (-1, -1),        5),
        ("RIGHTPADDING", (0, 0),      (-1, -1),        5),
        ("TOPPADDING",   (0, 0),      (-1, -1),        4),
        ("BOTTOMPADDING",(0, 0),      (-1, -1),        4),
        ("GRID",         (0, 0),      (-1, -1),        0.4, C_BORDER),
        ("BACKGROUND",   (0, n_rows - 1), (-1, n_rows - 1), C_GT_BG),
        ("FONTNAME",     (0, n_rows - 1), (-1, n_rows - 1), "Helvetica-Bold"),
        ("LINEABOVE",    (0, n_rows - 1), (-1, n_rows - 1), 1, C_BORDER),
    ]
    for idx in range(1, n_rows - 1):
        if idx % 2 == 0:
            cmds.append(("BACKGROUND", (0, idx), (-1, idx), C_ROW_ALT))
    return TableStyle(cmds)


def _sig_table(settings: dict, total_w: float) -> Table:
    lb = ParagraphStyle("SL", fontSize=8, fontName="Helvetica-Bold")
    nm = ParagraphStyle("SN", fontSize=9, fontName="Helvetica-Bold")
    tt = ParagraphStyle("ST", fontSize=8)

    def col(lbl, name, title):
        return [Paragraph(lbl, lb), Spacer(1, 0.4 * cm),
                Paragraph(name, nm), Paragraph(title, tt)]

    tbl = Table([[
        col("Prepared By:",  settings.get("prepared_by_name", ""),  settings.get("prepared_by_title", "")),
        col("Checked by:",   settings.get("checked_by_name", ""),   settings.get("checked_by_title", "")),
        col("Validated by:", settings.get("validated_by_name", ""), settings.get("validated_by_title", "")),
    ]], colWidths=[total_w / 3] * 3)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), C_SIG_BG),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING",   (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
        ("BOX",          (0, 0), (-1, -1), 0.5, C_BORDER),
        ("INNERGRID",    (0, 0), (-1, -1), 0.3, C_BORDER),
    ]))
    return tbl


def _make_title(label: str, grp_type: str) -> str:
    """Build report title from the sheet label."""
    import re
    # Strip trailing "(AST)" or "AST" if already in label, then append neatly
    clean = re.sub(r'\s*\(?\bast\b\)?', '', label, flags=re.IGNORECASE).strip()
    kind  = "Instructor Fee" if grp_type == "TRAINER" else "Fee"
    return f"{clean.title()} (AST) {kind}"


def _fm(v) -> str:
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return str(v) if v is not None else ""


def _fq(v) -> str:
    """Format hours: whole numbers without decimal."""
    try:
        q = float(v)
        return str(int(q)) if q == int(q) else f"{q:g}"
    except (TypeError, ValueError):
        return str(v) if v is not None else ""
