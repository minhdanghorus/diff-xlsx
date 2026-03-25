# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the tool

```bash
# Activate the virtual environment first
source venv/bin/activate          # macOS/Linux
venv\Scripts\activate.bat         # Windows CMD
& venv\Scripts\Activate.ps1       # Windows PowerShell

# Run
python diff_xlsx.py
```

Only dependency: `openpyxl==3.1.5` (install via `pip install -r requirements.txt`).

There are no tests and no lint configuration.

## Architecture

The entire tool lives in a single file: `diff_xlsx.py`. Execution flows strictly top-to-bottom through `main()`.

### `main()` prompt sequence

Steps run in this fixed order:
1. Discover files in `source/` → `get_files()`
2. Read both files → `read_sheet()` (dispatches to `read_xlsx` / `read_csv`)
3. Validate matching column headers → `check_format()`
4. Ask export format (`html` / `csv` / `xlsx`)
5. Load `aliases.txt` and `ignore_substring.txt`
6. Ask unique key column → `ask_unique_key()`
7. Ask case sensitivity → `ask_case_sensitive()`
8. Ask columns to skip from comparison → `ask_skip_columns()`
9. Ask min-width columns (HTML only) → `ask_min_width_columns()`
10. Compare rows → `compare_by_key()` or `compare_by_position()`
11. Ask columns to hide from report output → `ask_hide_columns()`
12. Ask extra post-substitution report → `ask_extra_report()`
13. Ask split report (HTML only) → `ask_split_report()`
14. Generate and write the main report
15. Generate and write the extra report (if requested)

### `diffs` data structure

The comparison functions return a list of dicts:
```python
{
    "type":    "changed" | "added" | "deleted",
    "label":   str,          # e.g. "Key: 42" or "Row 5"
    "row1":    list | None,  # values from file 1 (None for added)
    "row2":    list | None,  # values from file 2 (None for deleted)
    "changed": set[int],     # positional column indices that differ (changed rows only)
}
```
All indices in `changed` refer to positions in the **original full headers list**.

### Report generation pattern

Each format has two generators:
- **Main report**: `generate_html` / `generate_csv_report` / `generate_xlsx_report`
- **Extra (post-substitution) report**: `generate_extra_html` / `generate_extra_csv` / `generate_extra_xlsx`

The extra report is debug-only: it shows post-substitution values, only changed rows, only columns that differed. It intentionally does **not** receive `hide_columns`.

All main generators receive a `shared_kwargs` dict from `main()` containing: `file1_name`, `file2_name`, `case_sensitive`, `ignore_substrings`, `total_rows_compared`, `skip_columns`, `hide_columns`, `keep_in_summary`.

### Key helpers

- **`_filter_for_report(diffs, headers, hide_columns)`** — strips hidden columns from row data and remaps `changed` index sets. Called inside each main generator for the table/data section. The summary section always receives the original unfiltered diffs so it can show hidden-column counts when `keep_in_summary=True`.
- **`_summary_rows(...)`** — builds the summary section for CSV and XLSX. HTML builds its own inline equivalent. Both respect `hide_columns` + `keep_in_summary` via `summary_exclude = skip_columns if keep_in_summary else (skip_columns | hide_columns)`.
- **`apply_substitutions(value, sub_rules)`** — applied per-cell during comparison (not stored back into rows), and again inside extra-report generators.
- **`_build_col_subs(header_strs, ignore_substrings, filename)`** — builds a per-column list of substitution rules for one file.

### Configuration files (project root)

| File | Format | Purpose |
|---|---|---|
| `aliases.txt` | `ColName:(val1,val2)` | Treat two values as equivalent |
| `ignore_substring.txt` | `filename:ColName:(find,replacement)` | Substitute substrings before comparing |
| `column_widths.txt` | `ColName:widthpx` | Min column widths in HTML report |

### Split HTML reports

When splitting is enabled, reports go into a `reports/` subdirectory (cleared on each run). `grand_col_counts` is pre-computed from all diffs before chunking, then passed to each `generate_html` call so every part shows full-dataset totals alongside per-part counts.
