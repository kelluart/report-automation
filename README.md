# CSV to Excel Report Generator

Turn a client CSV into a clean, client-ready Excel report with exactly two sheets. The output is structured for fast review and consistent delivery.

## What you get

- **Clean_Data**: cleaned and de-duplicated rows, sorted by category -> date -> amount (when columns exist).
- **Summary**: metrics (rows_total, columns_total, duplicates_removed, null_columns_total, null_cells_total), plus Category Totals and Month Totals (when dates are available).

## Input format

- **File path:** `data/input.csv`
- CSV = comma-separated values.
- Headerless CSV is supported with columns: date, category, amount, description.
- Expected date format: `YYYY-MM-DD`.

## How to run (Windows PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python src/main.py
```

Output is written to `output/report.xlsx`.

## Verification

```powershell
python -c "from openpyxl import load_workbook; wb=load_workbook('output/report.xlsx'); print(wb.sheetnames)"
python -c "import pandas as pd; df=pd.read_excel('output/report.xlsx', sheet_name='Summary', header=None); print(df.head(30).to_string(index=False, header=False))"
```

## Screenshots (placeholders)
![Clean_Data screenshot](assets/clean_data.png)
![Summary screenshot](assets/summary.png)
