# MPCScript Project Documentation

## 1. Project Overview

MPCScript is a desktop Python application for Manila Polo Club instructor fee reporting. It reads Excel-based source files, calculates instructor fee summaries, previews the calculated rows in a Tkinter GUI, and exports styled reports as PDF or Excel workbooks.

The application supports two reporting workflows:

1. POS Instructor Fee Summary
   - Uses an ISP Reference List.
   - Uses a POS Sales-Income Center raw Excel export.
   - Groups transactions by income center.
   - Computes commission, EWT, and final net total per instructor.

2. AST Fee Report
   - Uses an ISP Reference List as a required gate in the UI, although current AST calculations are read from the AST file itself and do not match AST names against the ISP list.
   - Uses an After School Tennis trainer/ballboy Excel file.
   - Supports summary sheets and daily-log sheets.
   - Computes trainer and ballboy totals.

The main entry point is `main.py`, which starts the Tkinter window defined in `ui/main_window.py`.

## 2. Repository Structure

```text
MPCScript/
  main.py
  requirements.txt
  MPCScript.spec
  debug_reader.py
  config/
    __init__.py
    settings.py
  core/
    __init__.py
    raw_data_reader.py
    isp_reader.py
    calculator.py
    excel_exporter.py
    pdf_exporter.py
    ast_reader.py
    ast_excel_exporter.py
    ast_pdf_exporter.py
  ui/
    __init__.py
    main_window.py
    settings_dialog.py
    alias_dialog.py
    src/
      mpc_icon.ico
      MPC-LOGO-white-text-1.png
```

## 3. Runtime Dependencies

Dependencies are listed in `requirements.txt`:

```text
pandas>=2.0.0
openpyxl>=3.1.0
xlrd>=2.0.1
reportlab>=4.0.0
```

The code also imports:

- `tkinter`, included with most standard Python Windows installs.
- `PIL.Image` and `PIL.ImageTk` in `ui/main_window.py`. This requires Pillow, but Pillow is not currently listed in `requirements.txt`.

Recommended setup:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install Pillow
python main.py
```

## 4. Application Entry Point

`main.py`:

- Adds the project root to `sys.path`.
- Imports `MainWindow` from `ui.main_window`.
- Creates the main Tkinter application.
- Calls `mainloop()`.

Execution:

```powershell
python main.py
```

## 5. Persistent Application Data

The application stores user-level configuration under:

```text
%APPDATA%/MPCScript/
```

Files:

- `settings.json`
  - Signatory names and titles.
  - Club commission rate.

- `aliases.json`
  - Saved mappings from unmatched POS instructor names to ISP Reference List names.
  - Used only by the POS workflow.

Default settings from `config/settings.py`:

```json
{
  "prepared_by_name": "Gilfred C. Sale",
  "prepared_by_title": "Sports Supervisor",
  "checked_by_name": "Jazmin Montealegre",
  "checked_by_title": "S&A Executive Secretary",
  "validated_by_name": "Maricel Balingbing",
  "validated_by_title": "Billing Assistant",
  "commission_rate": 5.0
}
```

`commission_rate` is stored as a percentage value, for example `5.0` means 5 percent. During calculation it is converted to a decimal fraction, for example `0.05`.

## 6. User Interface Workflow

The main window is implemented in `ui/main_window.py`.

Main UI areas:

- Header
  - Manila Polo Club title and logo.
  - Settings button.

- Left control panel
  - Step 1: load ISP Reference List.
  - Step 2: load POS Raw Data File.
  - Generate Summary.
  - Export PDF and Export Excel.
  - Alternative AST Fee File section.
  - Generate AST Report.
  - Export AST PDF and Export AST Excel.

- Right content panel
  - Shows a placeholder before generation.
  - Shows a tabbed notebook after generation.
  - Each tab contains a Treeview preview table.

- Status bar
  - Displays load, generation, export, and error status.

## 7. Settings Dialog

Implemented in `ui/settings_dialog.py`.

The Settings dialog edits:

- Prepared By name and title.
- Checked By name and title.
- Validated By name and title.
- Club Commission Rate in percent.

When the user clicks Save:

1. Values are read from the form.
2. Commission rate is validated as a number.
3. Settings are saved to `%APPDATA%/MPCScript/settings.json`.
4. The dialog closes.

These settings are reloaded before POS summary generation and AST report generation.

## 8. POS Workflow

### 8.1 POS Workflow Summary

The POS workflow converts raw transaction-level POS data into per-instructor fee summaries.

High-level process:

1. User loads the ISP Reference List.
2. User loads the POS Raw Data File.
3. Application parses raw POS transactions.
4. User clicks Generate Summary.
5. Application reloads the ISP Reference List from disk.
6. Application groups transactions by income center and instructor.
7. Application matches POS instructor names to ISP names.
8. Application calculates commission, EWT, and net totals.
9. If unmatched names exist, an alias dialog lets the user map POS names to ISP names.
10. Application displays one tab per income center.
11. User exports PDF or Excel.

### 8.2 POS Input 1: ISP Reference List

Reader: `core/isp_reader.py`

Accepted file types:

- `.xlsx`
- `.xls`

Engine selection:

- `.xls` uses `xlrd`.
- Other Excel files use `openpyxl`.

The ISP file may contain title rows above the real headers. The reader scans the first 20 rows and looks for a header row with at least two of these keywords:

- `isp name`
- `tax rate`
- `sworn`
- `tin no`
- `acct no`

Expected logical fields:

- ISP name
- Tax rate
- TIN number
- Sworn declaration
- VAT registered
- Account number

Column matching uses header keywords first. If names are non-standard, there are positional fallbacks:

- ISP name fallback: third column, index 2.
- Tax rate fallback: sixth column, index 5.

Output from `isp_reader.load(filepath)`:

```python
{
    "joan tabanag": {
        "display_name": "Joan Tabanag",
        "tax_rate": 0.05,
        "tin_no": "227-123-513-000",
        "sworn_decl": "B",
        "vat_registered": "No",
        "acct_no": "ISP0181",
    }
}
```

Keys are lowercase instructor names. This enables fast lookup by POS instructor name.

Tax rate parsing:

- `5%` becomes `0.05`.
- `5` becomes `0.05`.
- `0.05` remains `0.05`.
- Blank, invalid, or NaN values become `0.0`.

Errors:

- If no header row is found, a `ValueError` explains that the file does not appear to be an ISP Reference List.
- If no valid ISP records are loaded, a `ValueError` explains that the list appears empty or could not be parsed.

### 8.3 POS Input 2: Raw POS Sales-Income Center Export

Reader: `core/raw_data_reader.py`

Accepted file types:

- `.xlsx`
- `.xls`

Engine selection:

- `.xls` uses `xlrd`.
- Other Excel files use `openpyxl`.

The POS file is not treated as a simple flat table. It is parsed as a formatted report with repeating sections:

```text
Report header rows
Income Center Group : ARCHERY
Income Center : Archery - Instructor Fees
Column header row
Transaction rows
Income Center Total row
Income Center Group Total row
```

The parser tracks the current:

- `income_center_group`
- `income_center`

It skips:

- Empty rows.
- Income center total rows.
- Income center group total rows.
- Grand total rows.

Header row detection:

A row is considered a column header row when it contains:

- `date`
- `item`
- `qty`

Recognized POS columns:

- Date
- Receipt No.
- Pax
- Item
- Guest
- Pay Guest
- Unit Price
- Qty
- Gross Price
- Disc
- Amt. After Disc.
- Svc
- Tax
- Total

Transaction row detection:

- A row is a data row only when the mapped Date column contains a recognized date.
- Dates can be pandas timestamps, Python datetimes, `m/d/yyyy`, `mm/dd/yyyy`, or `yyyy-mm-dd`.
- Accepted year range is 2015 through 2035.

Output from `raw_data_reader.read(filepath)`:

```python
(
    pandas.DataFrame,
    warnings_list
)
```

The DataFrame contains one row per transaction:

```text
date
receipt_no
pax
item
guest
unit_price
qty
gross_price
disc
amt_after_disc
svc
tax
total
income_center_group
income_center
```

Important behavior:

- Receipt numbers are preserved as text.
- A trailing `.0` is removed if pandas loaded the receipt as a float-like value.
- Numeric fields are converted safely to floats.
- Invalid numeric values become `0.0`.
- Blank Item values are included, but a warning is returned.

Errors:

- If no transaction rows are found, a `ValueError` tells the user to verify that the file is the correct POS Sales-Income Center export.

### 8.4 POS Calculation

Calculator: `core/calculator.py`

Function:

```python
summarise(df, isp_dict, commission_rate=0.05, aliases=None)
```

Grouping:

1. Group raw transaction rows by:
   - `income_center_group`
   - `income_center`
2. Inside each income center, group by:
   - `item`, which represents the instructor name.

Instructor matching:

1. Strip the POS Item value.
2. Convert it to lowercase.
3. If a saved alias exists for that POS name, use the alias as the lookup key.
4. Look up the instructor in `isp_dict`.
5. If no ISP match exists, EWT rate is set to `0.0` and a warning is generated.

Per-instructor calculations:

```text
min_unit_price = minimum unit_price
sum_qty        = sum(qty)
sum_amt        = sum(amt_after_disc)
comm           = sum_amt * commission_rate
total          = sum_amt - comm
ewt            = total * isp_tax_rate
final_total    = total - ewt
```

Amounts are rounded to two decimal places using Python `round(value, 2)`.

Output summary structure:

```python
{
    "Archery - Instructor Fees": {
        "rows": pandas.DataFrame,
        "grand_total": {
            "item": "Grand Total",
            "min_unit_price": None,
            "sum_qty": 10.0,
            "sum_amt": 10000.00,
            "comm": 500.00,
            "total": 9500.00,
            "ewt_rate": None,
            "ewt": 475.00,
            "final_total": 9025.00,
        },
        "date_min": datetime,
        "date_max": datetime,
        "income_center_group": "ARCHERY",
        "income_center": "Archery - Instructor Fees",
    }
}
```

Each `rows` DataFrame contains:

```text
item
min_unit_price
sum_qty
sum_amt
comm
total
ewt_rate
ewt
final_total
```

### 8.5 POS Alias Mapping

Dialog: `ui/alias_dialog.py`

The alias dialog appears automatically during POS summary generation if any POS instructor names are not found in the ISP list and do not already have saved aliases.

The dialog shows:

- POS Instructor Name.
- ISP List Entry dropdown.

Choices:

- `-- No match  (EWT = 0) --`
- All ISP display names.

The dialog uses `difflib.get_close_matches()` to preselect a likely ISP name when possible.

When the user clicks Apply & Regenerate:

1. Selected mappings are stored in memory.
2. They are saved to `%APPDATA%/MPCScript/aliases.json`.
3. POS summaries are recalculated with aliases applied.

When the user leaves a name as No match:

- The alias value is saved as `None`.
- Future runs will keep EWT at `0` for that POS name unless aliases are reset.

Alias reset behavior:

- When a newly loaded ISP list has a different set of display names from the previous ISP list, saved aliases are cleared.

### 8.6 POS Preview Output

After generation, the right panel displays one notebook tab per income center.

Tab labels are created by stripping `Instructor Fee` or `Instructor Fees` suffixes from the income center name and title-casing the result.

Preview columns:

```text
Instructor
Min. Unit Price
Qty
Amt. After Disc.
Commission
Total
EWT Rate
EWT
Net Total
```

The last row is a highlighted Grand Total row.

### 8.7 POS PDF Export

Exporter: `core/pdf_exporter.py`

Default output filename:

```text
Instructor_Fee_Summary.pdf
```

Output:

- Landscape A4 PDF.
- One page per income center.
- Styled Manila Polo Club title banner.
- Report title derived from income center name.
- Date range from transaction dates.
- Summary table.
- Grand total row.
- Signatory block.

PDF table columns:

```text
Item
Min of Unit Price
Sum of Qty
Sum of Amt. After Disc.
5% Comm
Total
Ewt
Total
```

The final `Total` column is the final net total.

### 8.8 POS Excel Export

Exporter: `core/excel_exporter.py`

Default output filename:

```text
Instructor_Fee_Summary.xlsx
```

Output:

- One worksheet per income center.
- Sheet names are title-cased and truncated to Excel's 31-character limit.
- Styled title rows, headers, alternating data rows, grand total row, and signatory block.

Worksheet layout:

```text
Row 1: blank
Row 2: Manila Polo Club Inc.
Row 3: report title and date range
Row 4: column headers
Row 5+: data rows
Next row: grand total
Next rows: signatories
```

Columns:

```text
Item
Min of Unit Price
Sum of Qty
Sum of Amt. After Disc.
5% Comm
Total
Ewt
Total
```

## 9. AST Workflow

### 9.1 AST Workflow Summary

The AST workflow reads After School Tennis fee files and exports trainer and ballboy reports.

High-level process:

1. User loads the ISP Reference List.
2. User loads the AST Fee File.
3. Application parses trainer and ballboy data from the AST workbook.
4. User clicks Generate AST Report.
5. Application reloads settings.
6. Application displays AST tabs.
7. User exports AST PDF or AST Excel.

The UI requires an ISP file before AST files can be selected. However, the current AST reader and AST exporters do not use ISP tax rates or ISP aliases.

### 9.2 AST Input File

Reader: `core/ast_reader.py`

The AST reader opens the workbook with `xlrd`:

```python
wb = xlrd.open_workbook(filepath)
```

Important implication:

- The current AST reader is built around `xlrd`.
- Modern `xlrd` versions support `.xls` but not `.xlsx`.
- The UI allows `.xlsx` and `.xls`, but `.xlsx` AST files may fail depending on the installed `xlrd` behavior.

The reader skips duplicate or previous-copy sheets whose names contain patterns like:

```text
(1)
(2)
```

Recognized AST formats:

1. Trainer summary sheet.
2. Ballboy summary sheet.
3. Per-day daily log sheet.

### 9.3 AST Trainer Summary Sheet

Detected when the sheet name contains:

```text
train
```

Expected structure:

- Title at row 2, column A in Excel terms.
- Date string at row 3, column A.
- Header row where column A is `NAMES`.

Recognized columns:

```text
NAMES
day columns
Total hours / Total hrs
Rate
Total Amount
Vat 12%
Total
5%Comm / 5% Comm
Total Net Amount
2%Tax / 2% Tax / EWT
Net Amt
```

Day columns are all non-empty headers between the name column and the total-hours column.

Rows skipped:

- Blank names.
- `TOTAL` or `TOTALS`.
- Signatory/admin rows containing words such as prepared, checked, validated, supervisor, manager, assistant, or secretary.

Trainer row output:

```text
name
one column per day
hours
rate
total_amount
vat
ex_vat
commission
net_amount
ewt
net_final
```

Important calculation:

- `net_final` is computed as `net_amount - ewt`.
- This is done because the source sheet can have a formula quirk where Net Amt is zero when EWT is zero.

Grand total:

```text
sum(net_final)
```

### 9.4 AST Ballboy Summary Sheet

Detected when the sheet name contains:

```text
ball boy
```

Expected structure:

- Title at row 3, column A in Excel terms.
- Date string at row 4, column A.
- Header row where column A is `NAME`.

Recognized columns:

```text
NAME
day columns
Total Hours / Total Hrs
Rate
Total
Vat 12%
Total Net Amount
```

Ballboy row output:

```text
name
one column per day
hours
rate
total
vat
net_total
```

If `Total Net Amount` is not available, net total falls back to:

```text
total - vat
```

Grand total:

```text
sum(net_total)
```

### 9.5 AST Daily Log Sheet

Detected when any cell in the first 20 rows matches a trainer header such as:

```text
Trainer & Coach
```

Supported layout:

- One or more day sections stacked vertically.
- One or more blocks side by side in the same sheet.

Typical per-day structure:

```text
Date: May 1, 2026
Trainer & Coach | Time In | Time out | Total | Rate | Total Amount
...
Ballboy         | Time In | Time out | Total | Rate | Total Amount
...
```

Date detection:

- Reads date markers from the block's first column.
- Supports date range text and single-date text.
- Falls back to parsing dates from the sheet name.
- If no date can be parsed, it uses today's date.

Hours detection:

1. Uses the `Total` column if present.
2. If total hours are zero, computes hours from `Time In` and `Time out` Excel time fractions:

```text
(time_out - time_in) * 24
```

Trainer daily-log calculations:

```text
total_h      = sum(day hours)
amount       = total_h * rate
vat          = amount * 12 / 112
ex_vat       = amount - vat
commission   = ex_vat * 0.05
net_amount   = ex_vat - commission
ewt          = 0.0
net_final    = net_amount
```

Trainer rows with zero rate are excluded.

Ballboy daily-log calculations:

```text
total_h      = sum(day hours)
total        = total_h * rate
vat          = total * 12 / 112
net_total    = total - vat
```

Output keys:

- `<period> (AST) Trainers`
- `<period> (AST) Ballboys`

### 9.6 AST Reader Output

`ast_reader.load(filepath)` returns:

```python
{
    "May 1-8, 2026 (AST) Trainers": {
        "type": "TRAINER",
        "title": "...",
        "date_str": "May 1-8, 2026",
        "date_min": datetime,
        "date_max": datetime,
        "day_cols": [("may_1", "May 1"), ...],
        "rows": pandas.DataFrame,
        "grand_total": 12345.67,
    },
    "May 1-8, 2026 (AST) Ballboys": {
        "type": "BALLBOY",
        "title": "...",
        "date_str": "May 1-8, 2026",
        "date_min": datetime,
        "date_max": datetime,
        "day_cols": [("may_1", "May 1"), ...],
        "rows": pandas.DataFrame,
        "grand_total": 12345.67,
    }
}
```

If no recognizable AST sheets are found, it raises `ValueError`.

### 9.7 AST Preview Output

After AST generation, the app creates notebook tabs for each AST summary group.

Trainer preview columns:

```text
Name
day columns
Total Hrs
Rate
Total Amt
VAT (12%)
Ex-VAT
5% Comm
Net Amount
EWT
Final Net
```

Ballboy preview columns:

```text
Name
day columns
Total Hrs
Rate
Total Amt
VAT (12%)
Net Total
```

Each table ends with a highlighted Grand Total row.

### 9.8 AST PDF Export

Exporter: `core/ast_pdf_exporter.py`

Default output filename:

```text
AST_Fee_Report.pdf
```

Output:

- Landscape A4 PDF.
- One page per AST group.
- Manila Polo Club title banner.
- Dynamic day columns.
- Trainer or ballboy-specific calculation columns.
- Grand total row.
- Signatory block.

Trainer PDF columns:

```text
Name
day columns
Total Hrs
Rate
Total Amount
VAT (12%)
Ex-VAT
5% Comm
Net Amount
EWT
Final Net
```

Ballboy PDF columns:

```text
Name
day columns
Total Hrs
Rate
Total Amount
VAT (12%)
Net Total
```

### 9.9 AST Excel Export

Exporter: `core/ast_excel_exporter.py`

Default output filename:

```text
AST_Fee_Report.xlsx
```

Output:

- One worksheet per AST group.
- Sheet names are sanitized for Excel and truncated to 31 characters.
- Dynamic day columns.
- Styled title rows, headers, data rows, grand total row, and signatories.

Trainer Excel columns:

```text
Name
day columns
Hours
Rate
Total Amount
VAT (12%)
Ex-VAT
5% Comm
Net Amount
EWT
Final Net
```

Ballboy Excel columns:

```text
Name
day columns
Hours
Rate
Total Amount
VAT (12%)
Net Total
```

## 10. Report Titles and Date Ranges

Report title helpers live in `core/calculator.py`.

### POS Date Range Format

Function:

```python
format_date_range(date_min, date_max)
```

Same month and year:

```text
May 1-8, 2026
```

Different months:

```text
Dec 28, 2025 - Jan 3, 2026
```

### POS Report Title

Function:

```python
group_to_title(center_name)
```

Examples:

```text
Archery - Instructor Fees -> Archery Instructor Fee
Pickleball Coach -> Pickleball Coach Instructor Fee
ARCHERY -> Archery Instructor Fee
```

### POS Tab Label

Function:

```python
tab_label(center_name)
```

Examples:

```text
Archery - Instructor Fees -> Archery
Pickleball Coach -> Pickleball Coach
```

## 11. Debug Utility

`debug_reader.py` is a diagnostic script for inspecting the raw structure of an Excel file.

Usage:

```powershell
python debug_reader.py "path\to\file.xls"
```

Behavior:

1. Selects `xlrd` for `.xls`, otherwise `openpyxl`.
2. Reads the file with no header row.
3. Prints the DataFrame shape.
4. Prints the first 40 rows.
5. For each non-empty cell, prints:
   - Column index.
   - Raw value representation.
   - Python type.

This is useful when a POS export or ISP file fails to parse because headers or date cells differ from expected layouts.

## 12. Packaging

PyInstaller spec file:

```text
MPCScript.spec
```

Purpose:

- Builds a GUI distribution named `MPCScript`.
- Uses `main.py` as the entry script.
- Bundles the Manila Polo Club logo PNG.
- Uses `ui/src/mpc_icon.ico` as the app icon.
- Hides the console window with `console=False`.
- Excludes large unused libraries such as matplotlib, scipy, IPython, notebook, pygments, docutils, and sphinx.

Build command:

```powershell
pyinstaller MPCScript.spec
```

Expected output folder:

```text
dist\MPCScript\
```

The spec includes hidden imports for pandas, openpyxl, xlrd, reportlab, and Pillow/Tkinter integration.

## 13. Module Responsibilities

### `main.py`

Starts the application.

### `config/settings.py`

Handles:

- Default settings.
- Loading and saving settings JSON.
- Loading and saving alias JSON.
- Ensuring `%APPDATA%/MPCScript` exists.

### `core/isp_reader.py`

Parses ISP Reference List Excel files into a lowercase-name lookup dictionary.

### `core/raw_data_reader.py`

Parses formatted POS Sales-Income Center Excel exports into a flat transaction DataFrame.

### `core/calculator.py`

Converts raw POS transactions and ISP lookup data into per-income-center summary tables.

### `core/excel_exporter.py`

Writes POS summaries to styled Excel workbooks.

### `core/pdf_exporter.py`

Writes POS summaries to styled landscape PDF reports.

### `core/ast_reader.py`

Parses AST trainer and ballboy Excel files, including summary sheets and daily-log sheets.

### `core/ast_excel_exporter.py`

Writes AST summaries to styled Excel workbooks.

### `core/ast_pdf_exporter.py`

Writes AST summaries to styled landscape PDF reports.

### `ui/main_window.py`

Builds the main Tkinter application, coordinates file selection, generation, previews, and exports.

### `ui/settings_dialog.py`

Provides the settings editor for signatories and commission rate.

### `ui/alias_dialog.py`

Provides the POS name-to-ISP-name mapping dialog for unmatched instructor names.

### `debug_reader.py`

Prints diagnostic information about raw Excel workbook structure.

### `MPCScript.spec`

PyInstaller build configuration.

## 14. Data Flow Diagrams

### POS Data Flow

```text
ISP Reference List Excel
        |
        v
core.isp_reader.load()
        |
        v
isp_dict

POS Raw Data Excel
        |
        v
core.raw_data_reader.read()
        |
        v
raw transactions DataFrame
        |
        v
core.calculator.summarise()
        |
        +-- uses isp_dict
        +-- uses saved aliases
        +-- uses commission rate setting
        |
        v
summaries dict
        |
        +--> UI Treeview preview
        +--> core.pdf_exporter.export()
        +--> core.excel_exporter.export()
```

### AST Data Flow

```text
AST Fee Excel
        |
        v
core.ast_reader.load()
        |
        v
AST summaries dict
        |
        +--> UI Treeview preview
        +--> core.ast_pdf_exporter.export()
        +--> core.ast_excel_exporter.export()
```

## 15. Inputs and Outputs Summary

### POS Inputs

```text
ISP Reference List: .xlsx or .xls
POS Raw Data File: .xlsx or .xls
Settings: %APPDATA%/MPCScript/settings.json
Aliases: %APPDATA%/MPCScript/aliases.json
```

### POS Outputs

```text
UI preview tabs by income center
Instructor_Fee_Summary.pdf
Instructor_Fee_Summary.xlsx
```

### AST Inputs

```text
ISP Reference List: required by UI before AST file can be selected
AST Fee File: .xls expected by current reader; .xlsx may fail with xlrd
Settings: %APPDATA%/MPCScript/settings.json
```

### AST Outputs

```text
UI preview tabs by AST group
AST_Fee_Report.pdf
AST_Fee_Report.xlsx
```

## 16. Error Handling and Warnings

Common error paths:

- ISP file has no recognizable header row.
- ISP file parses but contains no valid ISP names.
- POS raw data file contains no transaction rows.
- AST file contains no recognizable trainer, ballboy, or daily-log sheets.
- Export destination cannot be written.
- Commission rate setting is not numeric.

Warnings:

- POS rows with blank Item names are included and reported as a warning.
- POS instructors missing from the ISP list are reported and get EWT set to `0`.
- If aliases are applied, summaries are regenerated automatically.

## 17. Known Implementation Notes

1. Pillow is imported by the UI but is not listed in `requirements.txt`.

2. AST file selection allows `.xlsx`, but the AST reader uses `xlrd.open_workbook()` directly. With modern `xlrd`, `.xlsx` files are usually unsupported.

3. The POS PDF and Excel headers always say `5% Comm` in their column labels, even though the actual commission rate is configurable in settings.

4. AST daily-log calculations use a fixed 5 percent commission for trainers. They do not currently use the configurable commission rate from settings.

5. AST generation requires an ISP file in the UI, but AST calculations do not currently use ISP data.

6. POS summaries are stored in a dictionary keyed by income center name. If two different income center groups contain the same income center name, the later one can overwrite the earlier one.

7. Some source comments and UI strings appear to contain mojibake characters, likely from an encoding mismatch. This does not necessarily affect runtime behavior, but it can affect readability.

## 18. Maintenance Guide

### Adding a New POS Column

1. Update `_build_col_map()` in `core/raw_data_reader.py`.
2. Update `_extract_record()` if the column should be included in raw transaction rows.
3. Update `core/calculator.py` if the column affects calculations.
4. Update UI and exporters if the column should appear in reports.

### Changing POS Calculations

1. Update `summarise()` in `core/calculator.py`.
2. Verify the output DataFrame columns still match:
   - `ui/main_window.py`
   - `core/pdf_exporter.py`
   - `core/excel_exporter.py`
3. Update this documentation.

### Changing AST Calculations

1. Update `core/ast_reader.py`.
2. Verify trainer columns still match:
   - `core/ast_pdf_exporter.py`
   - `core/ast_excel_exporter.py`
   - AST preview table in `ui/main_window.py`
3. Verify ballboy columns still match those same modules.
4. Update this documentation.

### Adding a New Export Format

1. Add a new exporter module under `core/`.
2. Add buttons and handlers in `ui/main_window.py`.
3. Pass the existing summary dictionary into the exporter.
4. Keep the exporter responsible only for file generation, not parsing or calculation.

### Updating Signatory Fields

1. Update defaults in `config/settings.py`.
2. Update `ui/settings_dialog.py`.
3. Update all exporters that render signatory blocks.

## 19. Suggested Manual Test Checklist

POS workflow:

1. Start the app with `python main.py`.
2. Load a valid ISP Reference List.
3. Load a valid POS raw data file.
4. Generate Summary.
5. Confirm one tab appears per income center.
6. Confirm unmatched instructors open the alias dialog.
7. Apply an alias and confirm totals regenerate.
8. Export PDF.
9. Export Excel.
10. Open both exported files and verify totals, dates, and signatories.

AST workflow:

1. Start the app with `python main.py`.
2. Load a valid ISP Reference List.
3. Load a valid AST fee file.
4. Generate AST Report.
5. Confirm trainer and/or ballboy tabs appear.
6. Confirm day columns and totals are correct.
7. Export AST PDF.
8. Export AST Excel.
9. Open both exported files and verify totals, dates, and signatories.

Settings:

1. Open Settings.
2. Change signatory names and titles.
3. Change commission rate.
4. Save.
5. Generate a POS summary and verify the commission calculation uses the new rate.
6. Export reports and verify signatory values appear.

## 20. Quick Reference

Run app:

```powershell
python main.py
```

Install dependencies:

```powershell
pip install -r requirements.txt
pip install Pillow
```

Inspect an Excel file:

```powershell
python debug_reader.py "path\to\file.xls"
```

Build Windows distribution:

```powershell
pyinstaller MPCScript.spec
```

Persistent settings location:

```text
%APPDATA%/MPCScript/settings.json
%APPDATA%/MPCScript/aliases.json
```
