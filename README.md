# Excel Diff Tool

Compare two `.xlsx` files with the same format and generate a color-coded HTML diff report.

## Features

- Auto-detects the two files in the `source/` folder (prompts if more than 2 are found)
- Ignores fully-empty columns and trailing empty rows
- Compares only the first sheet of each file
- Supports two comparison modes:
  - **By unique key column** — detects added, deleted, and changed rows
  - **By row position** — detects changed rows at the same position
- Exact numeric and case-sensitive text comparison
- Outputs a styled HTML report with:
  - 🟡 Yellow highlight on changed cells
  - 🔴 Red background for deleted rows (only in File 1)
  - 🟢 Green background for added rows (only in File 2)

## Requirements

- Python 3.8+

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

## Usage

1. Place exactly 2 `.xlsx` files inside the `source/` folder:
   ```
   source/
   ├── file_a.xlsx
   └── file_b.xlsx
   ```
   If there are more than 2 files, the script will prompt you to choose.

2. Run the script:
   ```bash
   python diff_xlsx.py
   ```

3. Answer the prompts:
   ```
   Does the data have a unique key column? (yes/no) [no]:
   Enter the key column name [default: 'ID']:
   ```

4. Open the generated report:
   ```
   diff_report.html   ← saved in the project root
   ```

## Output Example

| Row / Key | Type | ID | Name | Score |
|-----------|------|----|------|-------|
| Key: 101  | OLD  | 101 | Alice | **85** |
| Key: 101  | NEW  | 101 | Alice | **90** |
| Key: 102  | DELETED | 102 | Bob | 70 |
| Key: 103  | ADDED   | 103 | Carol | 95 |

## Notes

- If no unique key column is used and the two files have a different number of rows, the script will raise an error and suggest using a key column.
- The `source/` folder must be in the same directory where you run the script.
