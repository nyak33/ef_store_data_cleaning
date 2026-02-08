# EF Store Data Cleaning

Deduplicate EF Store Excel files by keeping **one row per SN** (Serial Number), selecting the row with the **highest Scan Count**, and saving a cleaned `.xlsx` next to the original file.

## What It Does
- Prompts you to select an Excel file (`.xlsx`, `.xls`, `.xlsm`)
- Validates required columns: `SN` and `Scan Count`
- Keeps the highest `Scan Count` row per `SN`
- Writes a new file named `*_cleaned.xlsx`

## Requirements
- Python 3.8+
- `pandas`
- `openpyxl` (for writing `.xlsx`)
- Optional: `tqdm` (progress bars)

Install dependencies:
```bash
pip install pandas openpyxl tqdm
```

## Usage
Run the script:
```bash
python ef_store_deduplicate.py
```

Select your Excel file when the file dialog appears. The cleaned file is saved in the same folder as the original, with `_cleaned` appended to the filename.

## Input Columns
The script expects these columns (case-sensitive after trimming spaces):
- `SN`
- `Scan Count`

`Scan Count` is coerced to numeric. Invalid values become `NaN`, which are treated as the lowest values for deduplication.

## Output
If your input is:
```
EF_Store_Export.xlsx
```

You get:
```
EF_Store_Export_cleaned.xlsx
```

## Notes
- If multiple rows for the same `SN` have the same `Scan Count`, the first row encountered is kept.
- Set `DETAILED_PROGRESS = False` in `ef_store_deduplicate.py` to use a faster, vectorized deduplication path.
