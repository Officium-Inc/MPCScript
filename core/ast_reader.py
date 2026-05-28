"""
Reader for AST (After School Tennis) trainer/ballboy fee Excel files.

Recognises two summary sheet types:
  - TRAINER sheet  (sheet name contains "TRAIN")
      Columns: NAMES | ... | Total hours | Rate | Total Amount | Vat 12% |
               Total | 5%Comm | Total Net Amount | 2%Tax | Net Amt
  - BALLBOY sheet  (sheet name contains "BALLBOY")
      Columns: NAME | ... | Total Hours | Rate | Total | Vat 12% |
               Total Net Amount

Both sheets have 3-4 title rows at the top before the header row.

Returns
-------
dict keyed by display label, each value:
    {
        "type":       "TRAINER" | "BALLBOY",
        "title":      str,          # e.g. "Tennis Trainers (Ast)"
        "date_str":   str,          # e.g. "May 1-8, 2026"
        "date_min":   datetime,
        "date_max":   datetime,
        "rows":       pd.DataFrame, # one row per person
        "grand_total":float,
    }
"""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd

# ── Sheet-name patterns ──────────────────────────────────────────────────────
_RE_TRAINER = re.compile(r'train', re.IGNORECASE)
_RE_BALLBOY = re.compile(r'ball\s*boy', re.IGNORECASE)
_RE_PREV    = re.compile(r'\(\d+\)')   # e.g. "BALLBOY 2026 (1)"

# Per-day log sheet detection (sheet contains a 'Trainer & Coach' section header)
_RE_LOG_TRAINER = re.compile(r'^trainer', re.IGNORECASE)
_RE_LOG_BALLBOY = re.compile(r'^ball\s*boy', re.IGNORECASE)
_RE_DATE_HDR    = re.compile(r'date:\s*(.+)', re.IGNORECASE)

# ── Month lookup ─────────────────────────────────────────────────────────────
_MONTHS = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5,  'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10,'nov': 11, 'dec': 12,
}

# e.g. "May 1-8, 2026" or "April 25-29, 2026"
_RE_DATE_RANGE = re.compile(
    r'([A-Za-z]+)\s+(\d+)\s*[-–]\s*(\d+),?\s*(\d{4})',
    re.IGNORECASE,
)
# e.g. "April 29, 2026" or "APRIL 9 2026" or "MAY 1,2026" (no space after comma)
_RE_SINGLE_DATE = re.compile(
    r'([A-Za-z]+)\s+(\d{1,2}),?\s*(\d{4})',
    re.IGNORECASE,
)


def _parse_date_range(text: str) -> tuple[datetime, datetime]:
    m = _RE_DATE_RANGE.search(str(text))
    if m:
        mon_str, d1, d2, yr = m.groups()
        mon = _MONTHS.get(mon_str[:3].lower(), 1)
        y   = int(yr)
        try:
            return datetime(y, mon, int(d1)), datetime(y, mon, int(d2))
        except ValueError:
            pass
    today = datetime.today()
    return today, today


def _parse_any_date(text: str) -> tuple[datetime, datetime] | None:
    """Try date-range first, then single date.  Returns None if no match."""
    m = _RE_DATE_RANGE.search(str(text))
    if m:
        mon_str, d1, d2, yr = m.groups()
        mon = _MONTHS.get(mon_str[:3].lower(), 1)
        try:
            return datetime(int(yr), mon, int(d1)), datetime(int(yr), mon, int(d2))
        except ValueError:
            pass
    m = _RE_SINGLE_DATE.search(str(text))
    if m:
        mon_str, day, yr = m.groups()
        mon = _MONTHS.get(mon_str[:3].lower(), 1)
        try:
            d = datetime(int(yr), mon, int(day))
            return d, d
        except ValueError:
            pass
    return None


def _hdr_vals(sh, row: int) -> list[str]:
    return [str(sh.cell_value(row, c)).strip() for c in range(sh.ncols)]


def _find_hdr_row(sh, keyword: str) -> int:
    """Return 0-based index of the first row whose col-0 matches keyword."""
    kw = keyword.lower()
    for r in range(min(12, sh.nrows)):
        if str(sh.cell_value(r, 0)).strip().lower() == kw:
            return r
    return -1


def _flt(val) -> float:
    try:
        return float(val) if val not in ('', None) else 0.0
    except (TypeError, ValueError):
        return 0.0


def _col(hdrs: list[str], *candidates: str, after: int = 0) -> int:
    """Return the index of the first header matching any candidate (case-insensitive)."""
    for i, h in enumerate(hdrs):
        if i <= after:
            continue
        for cand in candidates:
            if h.lower() == cand.lower():
                return i
    # fallback: first match from position 0
    for i, h in enumerate(hdrs):
        for cand in candidates:
            if h.lower() == cand.lower():
                return i
    return -1


def _day_key(label: str) -> str:
    """'May 1' → 'may_1'  (safe DataFrame column name)."""
    return re.sub(r'\W+', '_', label.strip().lower()).strip('_')


# ── TRAINER sheet ────────────────────────────────────────────────────────────

def _read_trainer(sh) -> dict:
    title_raw = str(sh.cell_value(1, 0)).strip()
    title     = re.sub(r'\s+', ' ', title_raw).title()
    date_str  = str(sh.cell_value(2, 0)).strip()
    date_min, date_max = _parse_date_range(date_str)

    hdr_row = _find_hdr_row(sh, 'names')
    if hdr_row < 0:
        raise ValueError("Cannot find NAMES header row in trainer sheet")

    hdrs = _hdr_vals(sh, hdr_row)

    # Column indices
    c_name    = 0
    c_hours   = _col(hdrs, 'total hours', 'total hrs')
    c_rate    = _col(hdrs, 'rate')
    c_amount  = _col(hdrs, 'total amount')
    c_vat     = _col(hdrs, 'vat 12%')
    # "Total" after "Total Amount" = ex-VAT total
    c_exvat   = _col(hdrs, 'total', after=c_amount if c_amount >= 0 else 0)
    c_comm    = _col(hdrs, '5%comm', '5% comm')
    c_netcomm = _col(hdrs, 'total net amount')
    c_ewt     = _col(hdrs, '2%tax', '2% tax', 'ewt')
    # "Net Amt" / "Net  Amt" is the last column
    c_net     = sh.ncols - 1  # last col

    # Day columns: non-empty headers between description cols and c_hours
    day_col_defs = []  # list of (col_index, key, label)
    for i in range(1, c_hours if c_hours >= 0 else sh.ncols):
        h = hdrs[i].strip()
        if h:
            day_col_defs.append((i, _day_key(h), h))

    rows = []
    for r in range(hdr_row + 1, sh.nrows):
        name = str(sh.cell_value(r, c_name)).strip()
        if not name:
            continue
        if name.upper() in ('TOTAL', 'TOTALS'):
            continue
        # Skip signatory rows
        if any(kw in name.lower() for kw in (
            'prepared', 'checked', 'validated',
            'supervisor', 'manager', 'assistant', 'secretar',
        )):
            continue

        hours    = _flt(sh.cell_value(r, c_hours))    if c_hours    >= 0 else 0.0
        rate     = _flt(sh.cell_value(r, c_rate))     if c_rate     >= 0 else 0.0
        amount   = _flt(sh.cell_value(r, c_amount))   if c_amount   >= 0 else 0.0
        vat      = _flt(sh.cell_value(r, c_vat))      if c_vat      >= 0 else 0.0
        ex_vat   = _flt(sh.cell_value(r, c_exvat))    if c_exvat    >= 0 else amount - vat
        comm     = _flt(sh.cell_value(r, c_comm))     if c_comm     >= 0 else 0.0
        net_comm = _flt(sh.cell_value(r, c_netcomm))  if c_netcomm  >= 0 else ex_vat - comm
        ewt      = _flt(sh.cell_value(r, c_ewt))      if c_ewt      >= 0 else 0.0
        # Net Amt in sheet is 0 when EWT=0 (formula quirk); compute directly
        net_final = net_comm - ewt

        row_dict = {'name': name}
        for (ci, key, _) in day_col_defs:
            row_dict[key] = _flt(sh.cell_value(r, ci))
        row_dict.update({
            'hours':        hours,
            'rate':         rate,
            'total_amount': amount,
            'vat':          vat,
            'ex_vat':       ex_vat,
            'commission':   comm,
            'net_amount':   net_comm,
            'ewt':          ewt,
            'net_final':    net_final,
        })
        rows.append(row_dict)

    day_keys = [key for (_, key, _) in day_col_defs]
    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=['name'] + day_keys +
                ['hours', 'rate', 'total_amount', 'vat', 'ex_vat',
                 'commission', 'net_amount', 'ewt', 'net_final']
    )
    grand_total = float(df['net_final'].sum()) if not df.empty else 0.0

    return {
        'type':        'TRAINER',
        'title':       title,
        'date_str':    date_str,
        'date_min':    date_min,
        'date_max':    date_max,
        'day_cols':    [(key, lbl) for (_, key, lbl) in day_col_defs],
        'rows':        df,
        'grand_total': grand_total,
    }


# ── BALLBOY sheet ─────────────────────────────────────────────────────────────

def _read_ballboy(sh) -> dict:
    title_raw = str(sh.cell_value(2, 0)).strip()
    title     = re.sub(r'\s+', ' ', title_raw).title()
    date_str  = str(sh.cell_value(3, 0)).strip()
    date_min, date_max = _parse_date_range(date_str)

    hdr_row = _find_hdr_row(sh, 'name')
    if hdr_row < 0:
        raise ValueError("Cannot find NAME header row in ballboy sheet")

    hdrs = _hdr_vals(sh, hdr_row)

    c_name  = 0
    c_hours = _col(hdrs, 'total hours', 'total hrs')
    c_rate  = _col(hdrs, 'rate')
    c_total = _col(hdrs, 'total')
    c_vat   = _col(hdrs, 'vat 12%')
    # "Total Net  Amount" (note possible double space)
    c_net   = next(
        (i for i, h in enumerate(hdrs) if 'total net' in h.lower()), -1
    )

    # Day columns: non-empty headers between col 1 and c_hours
    day_col_defs = []
    for i in range(1, c_hours if c_hours >= 0 else sh.ncols):
        h = hdrs[i].strip()
        if h:
            day_col_defs.append((i, _day_key(h), h))

    rows = []
    for r in range(hdr_row + 1, sh.nrows):
        name = str(sh.cell_value(r, c_name)).strip()
        if not name:
            continue
        if name.upper() in ('TOTAL', 'TOTALS'):
            continue
        if any(kw in name.lower() for kw in (
            'prepared', 'checked', 'validated',
            'supervisor', 'manager', 'assistant', 'secretar',
        )):
            continue

        hours     = _flt(sh.cell_value(r, c_hours)) if c_hours >= 0 else 0.0
        rate      = _flt(sh.cell_value(r, c_rate))  if c_rate  >= 0 else 0.0
        total     = _flt(sh.cell_value(r, c_total)) if c_total >= 0 else 0.0
        vat       = _flt(sh.cell_value(r, c_vat))   if c_vat   >= 0 else 0.0
        net_total = _flt(sh.cell_value(r, c_net))   if c_net   >= 0 else total - vat

        row_dict = {'name': name}
        for (ci, key, _) in day_col_defs:
            row_dict[key] = _flt(sh.cell_value(r, ci))
        row_dict.update({
            'hours':     hours,
            'rate':      rate,
            'total':     total,
            'vat':       vat,
            'net_total': net_total,
        })
        rows.append(row_dict)

    day_keys = [key for (_, key, _) in day_col_defs]
    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=['name'] + day_keys + ['hours', 'rate', 'total', 'vat', 'net_total']
    )
    grand_total = float(df['net_total'].sum()) if not df.empty else 0.0

    return {
        'type':        'BALLBOY',
        'title':       title,
        'date_str':    date_str,
        'date_min':    date_min,
        'date_max':    date_max,
        'day_cols':    [(key, lbl) for (_, key, lbl) in day_col_defs],
        'rows':        df,
        'grand_total': grand_total,
    }


# ── Per-day log format (one or more daily sections stacked in a single sheet) ──────
# Structure per day: optional date row, court section, then:
#   Row N:   "Trainer & Coach" | Time In | Time out | Total | Rate | Total Amount
#   Row M:   "Ballboy"         | Time In | Time out | Total | Rate | Total Amount
# Multiple days may be stacked vertically in the same sheet.

_SKIP_NAMES = frozenset((
    'total', 'totals', 'prepared', 'checked', 'validated',
    'supervisor', 'manager', 'assistant', 'secretar',
))

_DAILY_HDRS = frozenset((
    'court', 'time in', 'time out', 'total w/o', 'total w/l',
    'name', 'names', 'ballboy', 'trainer', 'trainer & coach',
))


def _is_valid_daily_name(name: str) -> bool:
    """Return True only if 'name' looks like a real person (reject numbers,
    date rows, court/column headers, and admin keywords)."""
    if not name:
        return False
    # Reject numeric cell values (court numbers, time fractions stored as text)
    try:
        float(name)
        return False
    except ValueError:
        pass
    # Reject "Date: ..." prefix rows
    if _RE_DATE_HDR.match(name):
        return False
    lwr = name.lower().strip()
    # Reject known section / column headers
    if lwr in _DAILY_HDRS:
        return False
    # Reject administrative keywords (substring match)
    if any(kw in lwr for kw in _SKIP_NAMES):
        return False
    return True


def _is_daily_log(sh) -> bool:
    """Return True if any cell in the first 20 rows matches the 'Trainer & Coach' header."""
    for r in range(min(20, sh.nrows)):
        for c in range(sh.ncols):
            if _RE_LOG_TRAINER.match(str(sh.cell_value(r, c)).strip()):
                return True
    return False


def _daily_hours(sh, row: int, c_total: int) -> float:
    """Return hours from Total cell, or compute from Time In/Out (day fractions)."""
    hours = _flt(sh.cell_value(row, c_total)) if c_total >= 0 else 0.0
    if hours == 0.0:
        t_in  = _flt(sh.cell_value(row, 1))
        t_out = _flt(sh.cell_value(row, 2))
        if t_out > t_in:
            hours = round((t_out - t_in) * 24, 2)
    return hours


def _read_daily_log(sh, sheet_name: str) -> dict:
    """
    Read a per-day AST log sheet.  Handles both single-column and multi-column
    (side-by-side) layouts by locating every column that contains a
    "Trainer & Coach" header and treating each as an independent data block.
    Returns {label: data_dict} with up to two entries (TRAINER, BALLBOY).
    """
    # ---- Find all data-block base columns ----
    block_cols_set: set[int] = set()
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            if _RE_LOG_TRAINER.match(str(sh.cell_value(r, c)).strip()):
                block_cols_set.add(c)
    if not block_cols_set:
        return {}

    block_cols = sorted(block_cols_set)

    def _block_width(bi: int) -> int:
        """Columns belonging to block bi (distance to next block or sheet end)."""
        return (block_cols[bi + 1] - block_cols[bi]
                if bi + 1 < len(block_cols)
                else sh.ncols - block_cols[bi])

    def _hdr_cols_in_block(row: int, c0: int, bw: int) -> dict[str, int]:
        """Map lowercase header name → absolute column index within this block."""
        result: dict[str, int] = {}
        for c in range(c0, min(c0 + bw, sh.ncols)):
            h = str(sh.cell_value(row, c)).strip().lower()
            if h and h not in result:
                result[h] = c
        return result

    def _hours_from_row(row: int, c_tot: int, c_in: int, c_out: int) -> float:
        hours = _flt(sh.cell_value(row, c_tot)) if c_tot >= 0 else 0.0
        if hours == 0.0 and c_in >= 0 and c_out >= 0:
            t_in  = _flt(sh.cell_value(row, c_in))
            t_out = _flt(sh.cell_value(row, c_out))
            if t_out > t_in:
                hours = round((t_out - t_in) * 24, 2)
        return hours

    # ---- Per-block: collect date markers and section header rows ----
    # all_sections  : (t_row, c0, date_label, date_dt)
    # all_b_sections: (b_row, c0, date_label, date_dt)
    all_sections:   list[tuple] = []
    all_b_sections: list[tuple] = []

    for bi, c0 in enumerate(block_cols):
        # Date markers in this block's name column
        date_at_b: dict[int, tuple[str, datetime]] = {}
        for r in range(sh.nrows):
            raw = str(sh.cell_value(r, c0)).strip()
            m   = _RE_DATE_HDR.match(raw)
            candidate = m.group(1).strip() if m else raw
            p = _parse_any_date(candidate)
            if p is not None:
                lbl = f"{p[0].strftime('%B')} {p[0].day}, {p[0].year}"
                date_at_b[r] = (lbl, p[0])

        def _sec_date(hdr_row: int,
                      _dat: dict = date_at_b,
                      _sn: str   = sheet_name) -> tuple[str, datetime]:
            best = max((r for r in _dat if r <= hdr_row), default=None)
            if best is not None:
                return _dat[best]
            p = _parse_any_date(_sn)
            if p:
                return f"{p[0].strftime('%B')} {p[0].day}, {p[0].year}", p[0]
            return _sn, datetime.today()

        for r in range(sh.nrows):
            raw = str(sh.cell_value(r, c0)).strip()
            if _RE_LOG_TRAINER.match(raw):
                lbl, dt = _sec_date(r)
                all_sections.append((r, c0, lbl, dt))
            elif _RE_LOG_BALLBOY.match(raw):
                lbl, dt = _sec_date(r)
                all_b_sections.append((r, c0, lbl, dt))

    if not all_sections:
        return {}

    # ---- Unique day definitions (sorted by date) ----
    day_defs: dict[str, tuple[str, datetime]] = {}
    for (_, __, lbl, dt) in all_sections:
        k = _day_key(lbl)
        if k not in day_defs:
            day_defs[k] = (lbl, dt)

    day_col_list = sorted(day_defs.items(), key=lambda x: x[1][1])
    day_cols     = [(k, v[0]) for k, v in day_col_list]
    day_keys     = [k for k, _ in day_cols]

    # ---- Overall period string ----
    date_min = min(v[1] for v in day_defs.values())
    date_max = max(v[1] for v in day_defs.values())

    if date_min.date() == date_max.date():
        period_str = f"{date_min.strftime('%B')} {date_min.day}, {date_min.year}"
    elif date_min.year == date_max.year and date_min.month == date_max.month:
        period_str = (f"{date_min.strftime('%B')} "
                      f"{date_min.day}-{date_max.day}, {date_min.year}")
    else:
        period_str = (f"{date_min.strftime('%B')} {date_min.day} \u2013 "
                      f"{date_max.strftime('%B')} {date_max.day}, {date_max.year}")

    # ---- Aggregate trainer data ----
    trainer_agg: dict[str, dict] = {}

    for t_row, c0, lbl, _ in all_sections:
        dk = _day_key(lbl)
        bi = block_cols.index(c0)
        bw = _block_width(bi)

        next_b = next((br for (br, bc, _, __) in all_b_sections
                       if bc == c0 and br > t_row), None)
        next_t = next((tr for (tr, bc, _, __) in all_sections
                       if bc == c0 and tr > t_row), sh.nrows)
        t_end  = next_b if next_b is not None and next_b < next_t else next_t

        hcols  = _hdr_cols_in_block(t_row, c0, bw)
        c_tot  = hcols.get('total',    -1)
        c_rate = hcols.get('rate',     -1)
        c_in   = hcols.get('time in',  -1)
        c_out  = hcols.get('time out', -1)

        for r in range(t_row + 1, t_end):
            name = str(sh.cell_value(r, c0)).strip()
            if not _is_valid_daily_name(name):
                continue
            norm  = name.upper().strip()
            hours = _hours_from_row(r, c_tot, c_in, c_out)
            rate  = _flt(sh.cell_value(r, c_rate)) if c_rate >= 0 else 0.0
            if norm not in trainer_agg:
                trainer_agg[norm] = {'display_name': name, 'rate': rate}
            if rate > 0.0:
                trainer_agg[norm]['rate'] = rate
            trainer_agg[norm][dk] = trainer_agg[norm].get(dk, 0.0) + hours

    # ---- Build trainer DataFrame ----
    t_rows = []
    for p in trainer_agg.values():
        if p.get('rate', 0.0) == 0.0:   # exclude trainers with no rate/amount
            continue
        day_hrs = {dk: p.get(dk, 0.0) for dk in day_keys}
        total_h = sum(day_hrs.values())
        rate    = p.get('rate', 0.0)
        amount  = total_h * rate
        vat     = amount * 12 / 112
        ex_vat  = amount - vat
        comm    = ex_vat * 0.05
        net_amt = ex_vat - comm
        row     = {'name': p['display_name']}
        row.update(day_hrs)
        row.update({'hours': total_h, 'rate': rate, 'total_amount': amount,
                    'vat': vat, 'ex_vat': ex_vat, 'commission': comm,
                    'net_amount': net_amt, 'ewt': 0.0, 'net_final': net_amt})
        t_rows.append(row)

    t_cols     = (['name'] + day_keys +
                  ['hours', 'rate', 'total_amount', 'vat', 'ex_vat',
                   'commission', 'net_amount', 'ewt', 'net_final'])
    trainer_df = (pd.DataFrame(t_rows, columns=t_cols) if t_rows
                  else pd.DataFrame(columns=t_cols))
    trainer_gt = float(trainer_df['net_final'].sum()) if not trainer_df.empty else 0.0

    # ---- Aggregate ballboy data ----
    ballboy_agg: dict[str, dict] = {}

    for b_row, c0, lbl, _ in all_b_sections:
        dk = _day_key(lbl)
        bi = block_cols.index(c0)
        bw = _block_width(bi)

        next_t = next((tr for (tr, bc, _, __) in all_sections
                       if bc == c0 and tr > b_row), sh.nrows)
        b_end  = next_t

        hcols  = _hdr_cols_in_block(b_row, c0, bw)
        c_tot  = hcols.get('total',    -1)
        c_rate = hcols.get('rate',     -1)
        c_in   = hcols.get('time in',  -1)
        c_out  = hcols.get('time out', -1)

        for r in range(b_row + 1, b_end):
            name = str(sh.cell_value(r, c0)).strip()
            if not _is_valid_daily_name(name):
                continue
            norm  = name.upper().strip()
            hours = _hours_from_row(r, c_tot, c_in, c_out)
            rate  = _flt(sh.cell_value(r, c_rate)) if c_rate >= 0 else 0.0
            if norm not in ballboy_agg:
                ballboy_agg[norm] = {'display_name': name, 'rate': rate}
            if rate > 0.0:
                ballboy_agg[norm]['rate'] = rate
            ballboy_agg[norm][dk] = ballboy_agg[norm].get(dk, 0.0) + hours

    # ---- Build ballboy DataFrame ----
    b_rows = []
    for p in ballboy_agg.values():
        day_hrs   = {dk: p.get(dk, 0.0) for dk in day_keys}
        total_h   = sum(day_hrs.values())
        rate      = p.get('rate', 0.0)
        total     = total_h * rate
        vat       = total * 12 / 112
        net_total = total - vat
        row       = {'name': p['display_name']}
        row.update(day_hrs)
        row.update({'hours': total_h, 'rate': rate, 'total': total,
                    'vat': vat, 'net_total': net_total})
        b_rows.append(row)

    b_cols     = ['name'] + day_keys + ['hours', 'rate', 'total', 'vat', 'net_total']
    ballboy_df = (pd.DataFrame(b_rows, columns=b_cols) if b_rows
                  else pd.DataFrame(columns=b_cols))
    ballboy_gt = float(ballboy_df['net_total'].sum()) if not ballboy_df.empty else 0.0

    # ---- Return results ----
    results: dict = {}
    t_label = f"{period_str} (AST) Trainers"
    results[t_label] = {
        'type': 'TRAINER', 'title': t_label,
        'date_str': period_str, 'date_min': date_min, 'date_max': date_max,
        'day_cols': day_cols, 'rows': trainer_df, 'grand_total': trainer_gt,
    }
    b_label = f"{period_str} (AST) Ballboys"
    results[b_label] = {
        'type': 'BALLBOY', 'title': b_label,
        'date_str': period_str, 'date_min': date_min, 'date_max': date_max,
        'day_cols': day_cols, 'rows': ballboy_df, 'grand_total': ballboy_gt,
    }
    return results


# ── Public API ────────────────────────────────────────────────────────────────

def load(filepath: str) -> dict[str, dict]:
    """
    Load an AST fee Excel file.

    Returns a dict ordered by sheet order, keyed by display label.
    Raises ValueError if no recognisable AST sheets are found.
    """
    import xlrd
    wb = xlrd.open_workbook(filepath)

    summaries: dict[str, dict] = {}
    errors: list[str] = []

    for name in wb.sheet_names():
        if _RE_PREV.search(name):          # skip "(1)", "(2)" repeat sheets
            continue
        sh = wb.sheet_by_name(name)
        try:
            if _RE_TRAINER.search(name):
                data  = _read_trainer(sh)
                label = data['title'] or name.title()
                summaries[label] = data
            elif _RE_BALLBOY.search(name):
                data  = _read_ballboy(sh)
                label = data['title'] or name.title()
                summaries[label] = data
            elif _is_daily_log(sh):
                for lbl, d in _read_daily_log(sh, name).items():
                    summaries[lbl] = d
        except Exception as exc:
            errors.append(f"{name}: {exc}")

    if not summaries:
        detail = '; '.join(errors) if errors else 'no TRAINERS or BALLBOY sheets found'
        raise ValueError(f"No AST data found in '{filepath}': {detail}")

    return summaries
