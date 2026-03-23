# Diff Tool

Compare two `.xlsx` or `.csv` files with the same column structure and generate a detailed diff report in HTML, CSV, or XLSX format.

## Features

### File Support
- Supports `.xlsx` and `.csv` files (auto-detects delimiter for CSV)
- Reads the first sheet of each `.xlsx` file
- Automatically ignores fully-empty columns and trailing empty rows
- Prompts to choose files if more than 2 are found in the `source/` folder

### Comparison Modes
- **By unique key column** — matches rows by a key, detects added, deleted, and changed rows
- **By row position** — compares rows at the same index (requires equal row counts)
- **Case-sensitive or case-insensitive** comparison (user choice)
- Skip specific columns from comparison

### Value Normalization
- **Aliases** (`aliases.txt`) — treat two different values as equivalent for a given column
- **Substring substitution** (`ignore_substring.txt`) — strip or replace substrings per file per column before comparing

### Report Formats
- **HTML** — styled, color-coded report with sticky header
- **CSV** — flat file with `diff_type` and `row_key` columns
- **XLSX** — color-coded spreadsheet with frozen header row

### Report Output
- Color coding: yellow highlight for changed cells, red for OLD/deleted rows, green for NEW/added rows
- Summary box with: file names, comparison mode, total rows compared, changed/added/deleted counts, per-column diff counts
- **Split report** (HTML only) — split into multiple files by a fixed number of rows, saved in a `reports/` folder, with prev/next navigation links and `X/Y` counts per part
- **Post-substitution report** — additional report showing cell values after substitution rules were applied (what the tool actually compared)

### Column Display (HTML)
- **Minimum column width** — configurable via `column_widths.txt`; falls back to prompting the user
- Maximum column width with text wrapping (280px)
- Sticky table header when scrolling

---

## Requirements

- Python 3.8+
- `openpyxl`

## Setup

```bash
# 1. Clone or download this project

# 2. Create and activate a virtual environment

# Windows (PowerShell)
python -m venv venv
& venv\Scripts\Activate.ps1

# Windows (CMD)
python -m venv venv
venv\Scripts\activate.bat

# macOS / Linux
python -m venv venv
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Usage

1. Place your files in the `source/` folder:
   ```
   source/
   ├── file_a.xlsx
   └── file_b.xlsx
   ```
   More than 2 files? The tool will prompt you to choose which two to compare.

2. (Optional) Add configuration files in the project root — see [Configuration Files](#configuration-files).

3. Run:
   ```bash
   python diff_xlsx.py
   ```

4. Answer the prompts (all have defaults — just press Enter to skip):
   ```
   Export format (html/csv/xlsx) [html]:
   Does the data have a unique key column? (yes/no) [no]:
   Use case-sensitive comparison? (yes/no) [yes]:
   Enter column(s) to skip during comparison:
   Apply minimum column width in HTML report? (yes/no) [no]:
   Export additional report with post-substitution values? (yes/no) [no]:
   Split report into multiple files? (yes/no) [no]:
   ```

5. Open the generated report in the project root:
   ```
   diff_report.html       (or .csv / .xlsx)
   diff_report_substituted.html   (if extra report was requested)
   reports/               (if split was chosen)
   ```

---

## Configuration Files

All configuration files are optional and live in the project root alongside `diff_xlsx.py`.

### `aliases.txt` — Treat two values as equivalent

Format: `ColumnName:(value1,value2)`

```
status:(OPENING,opening)
gender:(Male,male)
```

### `ignore_substring.txt` — Strip or replace substrings before comparing

Format: `filename:ColumnName:(find,replacement)`
Use an empty replacement to remove the substring.

```
file_a.xlsx:mobile:(+84 ,)
file_b.csv:mobile:(+84,)
file_b.csv:mobile:( ,)
```

> **Tip:** If a rule is not being applied, double-check that the filename in the rule exactly matches the filename in `source/` (including case and extension).

### `column_widths.txt` — Set minimum column widths in the HTML report

Format: `ColumnName:width`

```
name:150px
street:200px
mobile:120px
```

Columns not listed fall back to a default of `100px` when min-width is enabled.

---

## Output Example

### Summary Box (HTML)

```
File 1: file_a.xlsx
File 2: file_b.xlsx
Comparison mode: Case-insensitive
Total rows compared: 50

Changed rows: 12
Added rows (only in File 2): 3
Deleted rows (only in File 1): 1

Changed cells per column:
  - mobile: 8 row(s) differ
  - street: 4 row(s) differ
```

### Diff Table

| Row / Key | Type    | id  | name        | mobile      |
|-----------|---------|-----|-------------|-------------|
| Key: 2    | OLD     | 2   | Nguyen Van A | +84 090...  |
| Key: 2    | NEW     | 2   | Nguyen Van A | 090...      |
| Key: 99   | ADDED   | 99  | Tran Thi B  | 091...      |
| Key: 5    | DELETED | 5   | Le Van C    | 093...      |

### Split Report Navigation

When splitting is enabled, each file includes:
```
Part 2 of 5  |  <- Previous  |  Next ->

Changed rows: 3/12   Added rows: 1/3   Deleted rows: 0/1
```

---

## Notes

- Without a unique key column, both files must have the same number of rows.
- The `source/` folder must be in the same directory where you run the script.
- Split reports are always saved inside a `reports/` subfolder (cleared on each run).
- The post-substitution report only includes `changed` rows and only the columns that had differences.
