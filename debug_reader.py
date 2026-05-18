"""
Diagnostic script — run this to inspect the raw structure of the POS .xls file.
Usage: python debug_reader.py "path\\to\\your\\file.xls"
"""
import sys
from pathlib import Path

import pandas as pd

if len(sys.argv) < 2:
    print("Usage: python debug_reader.py <path_to_xls_file>")
    sys.exit(1)

filepath = sys.argv[1]
path = Path(filepath)
engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"

print(f"Reading: {filepath}")
print(f"Engine:  {engine}\n")

raw = pd.read_excel(filepath, engine=engine, header=None)

print(f"Shape: {raw.shape}  ({raw.shape[0]} rows x {raw.shape[1]} cols)\n")
print("=" * 80)
print("First 40 rows — value | type for each non-empty cell:")
print("=" * 80)

for row_idx in range(min(40, len(raw))):
    row = raw.iloc[row_idx]
    parts = []
    for col_idx, val in enumerate(row):
        try:
            is_na = pd.isna(val)
        except Exception:
            is_na = False
        if not is_na and str(val).strip() not in ("", "nan"):
            parts.append(f"  [{col_idx}] {repr(val)} ({type(val).__name__})")
    if parts:
        print(f"\nRow {row_idx}:")
        for p in parts:
            print(p)
    else:
        print(f"\nRow {row_idx}: <empty>")
