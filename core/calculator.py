"""
Computes per-instructor fee summaries from parsed transaction data.

Calculations per instructor (grouped by income_center_group → item):
    Sum Amt After Disc = sum of amt_after_disc
    Comm              = Sum Amt After Disc × commission_rate  (default 5 %)
    Total             = Sum Amt After Disc − Comm
    EWT               = Total × ISP tax_rate
    Final Total       = Total − EWT

Returns a dict keyed by group name, each value being a result dict:
    {
        "rows": DataFrame,       # one row per instructor (sorted by name)
        "grand_total": dict,     # aggregated totals row
        "date_min": datetime,
        "date_max": datetime,
    }
"""
from datetime import datetime
import re as _re

import pandas as pd


COMMISSION_RATE_DEFAULT = 0.05


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def summarise(
    df: pd.DataFrame,
    isp_dict: dict,
    commission_rate: float = COMMISSION_RATE_DEFAULT,
    aliases: dict | None = None,
) -> tuple[dict, list[str]]:
    """
    Args:
        df:               flat DataFrame from raw_data_reader.read()
        isp_dict:         dict from isp_reader.load()
        commission_rate:  fixed commission fraction (e.g. 0.05 for 5 %)
        aliases:          {pos_name_lower: isp_display_name} saved mappings;
                          None means no aliases.

    Returns:
        (summaries, warnings)
        summaries: {group_name: {"rows": DataFrame, "grand_total": dict,
                                 "date_min": datetime, "date_max": datetime}}
    """
    if aliases is None:
        aliases = {}
    warnings: list = []
    summaries: dict = {}
    isp_matchers = _build_isp_matchers(isp_dict)

    for (group_name, center_name), group_df in df.groupby(
        ["income_center_group", "income_center"], sort=True
    ):
        rows = []

        for instructor, inst_df in group_df.groupby("item", sort=True):
            instructor_stripped = str(instructor).strip()
            if not instructor_stripped:
                continue

            # ISP lookup: aliases, exact ISP name, then ISP-name prefix.
            isp_info, trainer_name = _resolve_isp_info(
                instructor_stripped, isp_dict, aliases, isp_matchers
            )
            if isp_info is None:
                warnings.append(
                    f"'{instructor_stripped}' was not found in the ISP list. "
                    "EWT set to 0."
                )
                ewt_rate = 0.0
                trainer_name = instructor_stripped
            else:
                ewt_rate = isp_info["tax_rate"]

            min_unit_price = _r2(float(inst_df["unit_price"].min()))
            sum_qty        = float(inst_df["qty"].sum())   # keep as float to preserve 0.5 etc.
            sum_amt        = _r2(float(inst_df["amt_after_disc"].sum()))
            comm           = _r2(sum_amt * commission_rate)
            total          = _r2(sum_amt - comm)
            ewt            = _r2(total * ewt_rate)
            final_total    = _r2(total - ewt)

            rows.append({
                "item":             instructor_stripped,
                "trainer":          trainer_name,
                "min_unit_price":   min_unit_price,
                "sum_qty":          sum_qty,
                "sum_amt":          sum_amt,
                "comm":             comm,
                "total":            total,
                "ewt_rate":         ewt_rate,
                "ewt":              ewt,
                "final_total":      final_total,
            })

        if not rows:
            continue

        summary_df = pd.DataFrame(rows)
        trainer_summary_df = _build_trainer_summary(summary_df)

        grand_total = _build_grand_total(summary_df)

        # Date range from the raw transactions in this group
        dates = pd.to_datetime(group_df["date"], errors="coerce").dropna()
        date_min = dates.min().to_pydatetime() if not dates.empty else datetime.today()
        date_max = dates.max().to_pydatetime() if not dates.empty else datetime.today()

        summaries[center_name] = {
            "rows":               summary_df,
            "trainer_summary":    trainer_summary_df,
            "trainer_grand_total": _build_grand_total(trainer_summary_df),
            "grand_total":        grand_total,
            "date_min":           date_min,
            "date_max":           date_max,
            "income_center_group": group_name,
            "income_center":       center_name,
        }

    return summaries, warnings


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _r2(value: float) -> float:
    """Round to 2 decimal places using standard rounding."""
    return round(value, 2)


def _normalise_name(value: str) -> str:
    """Lowercase and collapse whitespace for name matching."""
    return _re.sub(r"\s+", " ", str(value).strip().lower())


def _build_isp_matchers(isp_dict: dict) -> list[tuple[str, dict]]:
    """Return ISP records sorted longest-name first for prefix matching."""
    return sorted(
        ((_normalise_name(v["display_name"]), v) for v in isp_dict.values()),
        key=lambda item: len(item[0]),
        reverse=True,
    )


def _resolve_isp_info(
    instructor_name: str,
    isp_dict: dict,
    aliases: dict,
    isp_matchers: list[tuple[str, dict]],
) -> tuple[dict | None, str]:
    """
    Resolve a POS item name to an ISP record and display trainer name.

    Exact aliases keep their existing behavior. If no alias exists, a POS item
    can still match an ISP name when the ISP name is the leading part of the
    item, e.g. "Catalino Casas 2kids" -> "Catalino Casas".
    """
    lookup_key = _normalise_name(instructor_name)
    if lookup_key in aliases:
        alias_value = aliases[lookup_key]
        if alias_value is None:
            return None, instructor_name
        lookup_key = _normalise_name(alias_value)
        isp_info = isp_dict.get(lookup_key)
        return isp_info, isp_info["display_name"] if isp_info else str(alias_value)

    isp_info = isp_dict.get(lookup_key)
    if isp_info:
        return isp_info, isp_info["display_name"]

    for isp_name, candidate_info in isp_matchers:
        if _is_name_prefix(lookup_key, isp_name):
            return candidate_info, candidate_info["display_name"]

    return None, instructor_name


def _is_name_prefix(item_name: str, isp_name: str) -> bool:
    """Match exact names or names followed by a separator/suffix."""
    if item_name == isp_name:
        return True
    if not item_name.startswith(isp_name):
        return False
    suffix = item_name[len(isp_name):]
    return bool(suffix) and suffix[0] in {" ", "-", "/", "(", "["}


def _build_trainer_summary(summary_df: pd.DataFrame) -> pd.DataFrame:
    """Collapse detail rows to one row per resolved ISP trainer."""
    summary_rows = []
    for trainer, trainer_df in summary_df.groupby("trainer", sort=True):
        rates = trainer_df["ewt_rate"].dropna().unique()
        ewt_rate = float(rates[0]) if len(rates) else 0.0
        summary_rows.append({
            "item":             trainer,
            "trainer":          trainer,
            "min_unit_price":   None,
            "sum_qty":          float(trainer_df["sum_qty"].sum()),
            "sum_amt":          _r2(float(trainer_df["sum_amt"].sum())),
            "comm":             _r2(float(trainer_df["comm"].sum())),
            "total":            _r2(float(trainer_df["total"].sum())),
            "ewt_rate":         ewt_rate,
            "ewt":              _r2(float(trainer_df["ewt"].sum())),
            "final_total":      _r2(float(trainer_df["final_total"].sum())),
        })
    return pd.DataFrame(summary_rows)


def _build_grand_total(summary_df: pd.DataFrame) -> dict:
    """Build a grand-total row for any calculator summary DataFrame."""
    return {
        "item":           "Grand Total",
        "min_unit_price": None,
        "sum_qty":        float(summary_df["sum_qty"].sum()),
        "sum_amt":        _r2(float(summary_df["sum_amt"].sum())),
        "comm":           _r2(float(summary_df["comm"].sum())),
        "total":          _r2(float(summary_df["total"].sum())),
        "ewt_rate":       None,
        "ewt":            _r2(float(summary_df["ewt"].sum())),
        "final_total":    _r2(float(summary_df["final_total"].sum())),
    }


def format_date_range(date_min: datetime, date_max: datetime) -> str:
    """
    Format two dates as a compact range string.
        Same month/year:  "May 1-8, 2026"
        Different months: "Dec 28, 2025 - Jan 3, 2026"
    """
    if date_min.year == date_max.year and date_min.month == date_max.month:
        return f"{date_min.strftime('%b')} {date_min.day}-{date_max.day}, {date_max.year}"
    # Use .day (int) to avoid platform-specific strftime flags like %-d
    return (
        f"{date_min.strftime('%b')} {date_min.day}, {date_min.year} - "
        f"{date_max.strftime('%b')} {date_max.day}, {date_max.year}"
    )


def group_to_title(center_name: str) -> str:
    """
    Convert an income_center name to a report title.
        "Archery - Instructor Fees"  → "Archery Instructor Fee"
        "Pickleball Coach"           → "Pickleball Coach Instructor Fee"
        "ARCHERY"                    → "Archery Instructor Fee"  (legacy)
    """
    cleaned = _re.sub(
        r'\s*-?\s*instructor fees?\s*$', '', center_name, flags=_re.IGNORECASE
    ).strip(' -')
    label = cleaned.title() if cleaned else center_name.title()
    return label + " Instructor Fee"


def tab_label(center_name: str) -> str:
    """
    Short label for a notebook tab (strips 'Instructor Fees' suffix).
        "Archery - Instructor Fees"  → "Archery"
        "Pickleball Coach"           → "Pickleball Coach"
    """
    cleaned = _re.sub(
        r'\s*-?\s*instructor fees?\s*$', '', center_name, flags=_re.IGNORECASE
    ).strip(' -')
    return cleaned.title() if cleaned else center_name.title()
