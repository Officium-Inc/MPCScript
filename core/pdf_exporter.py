"""
Generates a PDF summary report using ReportLab.

One page per Income Center Group:
  ┌──────────────────────────────────────────┐
  │  Manila Polo Club Inc.                   │
  │  Archery Instructor Fee    May 1-8, 2026 │
  │                                          │
  │  [Summary table]                         │
  │                                          │
  │  Prepared By:  Checked by:  Validated by:│
  │  Name          Name          Name        │
  │  Title         Title         Title       │
  └──────────────────────────────────────────┘
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

from core.calculator import format_date_range, group_to_title

# ---------------------------------------------------------------------------
# Colour palette (matching the POS report aesthetic)
# ---------------------------------------------------------------------------
C_HEADER_BG   = colors.HexColor("#E8884C")   # orange header
C_HEADER_FG   = colors.white
C_TITLE_BG    = colors.HexColor("#C6EFCE")   # light green title band
C_GT_BG       = colors.HexColor("#FFF3E0")   # soft amber grand-total row
C_BORDER      = colors.HexColor("#B0B0B0")
C_ROW_ALT     = colors.HexColor("#FFF8F4")   # very light alternating row
C_SIG_BG      = colors.HexColor("#E8F5E9")   # light green signatory band

# ---------------------------------------------------------------------------
# Column definitions
# ---------------------------------------------------------------------------
COL_HEADERS = [
    "Item",
    "Min of\nUnit Price",
    "Sum of\nQty",
    "Sum of Amt.\nAfter Disc.",
    "5% Comm",
    "Total",
    "Ewt",
    "Total",          # Final total
]

# Column widths in points (landscape A4 content ≈ 756 pt)
COL_WIDTHS = [160, 75, 55, 95, 75, 80, 65, 80]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def export(summaries: dict, settings: dict, output_path: str) -> None:
    """
    Write a multi-page PDF (one page per group) to output_path.

    Args:
        summaries:   result from calculator.summarise()
        settings:    dict from config.settings.load()
        output_path: full file path for the output PDF
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
    group_names = list(summaries.keys())

    for i, group_name in enumerate(group_names):
        data = summaries[group_name]
        _build_group_page(story, group_name, data, settings, styles)
        if i < len(group_names) - 1:
            story.append(PageBreak())

    doc.build(story)


# ---------------------------------------------------------------------------
# Internal builders
# ---------------------------------------------------------------------------

def _build_group_page(
    story: list,
    group_name: str,
    data: dict,
    settings: dict,
    styles,
) -> None:
    """Append all flowables for one group to story."""
    title = group_to_title(group_name)
    date_range = format_date_range(data["date_min"], data["date_max"])

    # --- Title block ---
    title_style = ParagraphStyle(
        "MPCTitle",
        parent=styles["Normal"],
        fontSize=11,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
    )
    sub_style = ParagraphStyle(
        "MPCSub",
        parent=styles["Normal"],
        fontSize=10,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
    )
    date_style = ParagraphStyle(
        "MPCDate",
        parent=styles["Normal"],
        fontSize=9,
        alignment=TA_RIGHT,
    )

    # Company name + report title in a 2-row banner table
    banner_data = [
        [Paragraph("Manila Polo Club Inc.", title_style), ""],
        [Paragraph(title, sub_style), Paragraph(date_range, date_style)],
    ]
    banner_table = Table(banner_data, colWidths=[sum(COL_WIDTHS) * 0.65, sum(COL_WIDTHS) * 0.35])
    banner_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), C_TITLE_BG),
            ("SPAN", (0, 0), (1, 0)),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("BOX", (0, 0), (-1, -1), 0.5, C_BORDER),
        ])
    )
    story.append(banner_table)
    story.append(Spacer(1, 0.3 * cm))

    # --- Summary table ---
    table_data = _build_table_data(data["rows"], data["grand_total"])
    summary_table = Table(table_data, colWidths=COL_WIDTHS, repeatRows=1)
    summary_table.setStyle(_build_table_style(len(table_data)))
    story.append(summary_table)
    story.append(Spacer(1, 0.8 * cm))

    # --- Signatory footer ---
    story.append(_build_signatory_table(settings, styles))


def _build_table_data(rows_df, grand_total: dict) -> list:
    """Build the 2-D list that ReportLab's Table expects."""
    num_style = ParagraphStyle("Num", alignment=TA_RIGHT, fontSize=8)
    item_style = ParagraphStyle("Item", alignment=TA_LEFT, fontSize=8)
    hdr_style = ParagraphStyle(
        "Hdr",
        alignment=TA_CENTER,
        fontSize=8,
        fontName="Helvetica-Bold",
        textColor=C_HEADER_FG,
    )
    gt_style = ParagraphStyle(
        "GT",
        alignment=TA_LEFT,
        fontSize=8,
        fontName="Helvetica-Bold",
    )
    gt_num_style = ParagraphStyle(
        "GTNum",
        alignment=TA_RIGHT,
        fontSize=8,
        fontName="Helvetica-Bold",
    )

    table_data = [[Paragraph(h, hdr_style) for h in COL_HEADERS]]

    for _, row in rows_df.iterrows():
        table_data.append([
            Paragraph(str(row["item"]), item_style),
            Paragraph(_fmt(row["min_unit_price"]), num_style),
            Paragraph(_fmt_qty(row["sum_qty"]), num_style),
            Paragraph(_fmt(row["sum_amt"]), num_style),
            Paragraph(_fmt(row["comm"]), num_style),
            Paragraph(_fmt(row["total"]), num_style),
            Paragraph(_fmt(row["ewt"]), num_style),
            Paragraph(_fmt(row["final_total"]), num_style),
        ])

    gt = grand_total
    table_data.append([
        Paragraph(gt["item"], gt_style),
        Paragraph("", gt_num_style),
        Paragraph(_fmt_qty(gt["sum_qty"]), gt_num_style),
        Paragraph(_fmt(gt["sum_amt"]), gt_num_style),
        Paragraph(_fmt(gt["comm"]), gt_num_style),
        Paragraph(_fmt(gt["total"]), gt_num_style),
        Paragraph(_fmt(gt["ewt"]), gt_num_style),
        Paragraph(_fmt(gt["final_total"]), gt_num_style),
    ])

    return table_data


def _build_table_style(num_rows: int) -> TableStyle:
    """Build the full table style (header + alternating rows + grand total)."""
    commands = [
        # Header row
        ("BACKGROUND", (0, 0), (-1, 0), C_HEADER_BG),
        ("TEXTCOLOR", (0, 0), (-1, 0), C_HEADER_FG),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        # All cells
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        # Grid
        ("GRID", (0, 0), (-1, -1), 0.4, C_BORDER),
        # Grand total row
        ("BACKGROUND", (0, num_rows - 1), (-1, num_rows - 1), C_GT_BG),
        ("FONTNAME", (0, num_rows - 1), (-1, num_rows - 1), "Helvetica-Bold"),
        ("LINEABOVE", (0, num_rows - 1), (-1, num_rows - 1), 1, C_BORDER),
    ]

    # Alternating row colours (data rows only, skip header at index 0)
    for row_idx in range(1, num_rows - 1):
        if row_idx % 2 == 0:
            commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), C_ROW_ALT))

    return TableStyle(commands)


def _build_signatory_table(settings: dict, styles) -> Table:
    """Three-column signatory block at the bottom of the page."""
    label_style = ParagraphStyle(
        "SigLabel",
        fontSize=8,
        fontName="Helvetica-Bold",
        alignment=TA_LEFT,
    )
    name_style = ParagraphStyle("SigName", fontSize=9, fontName="Helvetica-Bold")
    title_style = ParagraphStyle("SigTitle", fontSize=8, fontName="Helvetica")

    def sig_col(label, name, title):
        return [
            Paragraph(label, label_style),
            Spacer(1, 0.4 * cm),
            Paragraph(name, name_style),
            Paragraph(title, title_style),
        ]

    sig_data = [[
        sig_col("Prepared By:", settings.get("prepared_by_name", ""),
                 settings.get("prepared_by_title", "")),
        sig_col("Checked by:", settings.get("checked_by_name", ""),
                 settings.get("checked_by_title", "")),
        sig_col("Validated by:", settings.get("validated_by_name", ""),
                 settings.get("validated_by_title", "")),
    ]]

    total_w = sum(COL_WIDTHS)
    sig_table = Table(sig_data, colWidths=[total_w / 3] * 3)
    sig_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), C_SIG_BG),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("BOX", (0, 0), (-1, -1), 0.5, C_BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.3, C_BORDER),
        ])
    )
    return sig_table


def _fmt(value) -> str:
    """Format a number with 2 decimal places and thousands separator."""
    if value is None:
        return ""
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value)


def _fmt_qty(value) -> str:
    """Format a quantity: whole numbers show without decimal, fractions show as-is.
    e.g.  1.0 → '1',  0.5 → '0.5',  9.5 → '9.5'
    """
    if value is None:
        return ""
    try:
        q = float(value)
        return str(int(q)) if q == int(q) else f"{q:g}"
    except (TypeError, ValueError):
        return str(value)
