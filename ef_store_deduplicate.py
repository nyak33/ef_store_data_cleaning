import os
import sys
import time
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Purpose:
# - Load an Excel file selected by the user
# - Keep one row per SN (Serial Number), choosing the highest Scan Count
# - Save the cleaned data to a new Excel file next to the original

# Optional: per-SN detailed progress (slower but shows a progress bar per group)
DETAILED_PROGRESS = True

# Try tqdm for progress bars; fall back to a no-op shim if missing
try:
    from tqdm import tqdm
except Exception:
    class tqdm:  # simple shim
        def __init__(self, iterable=None, total=None, desc=None):
            self.iterable = iterable
        def update(self, n=1): pass
        def close(self): pass
        def __iter__(self):
            return iter(self.iterable) if self.iterable is not None else iter(())
    print("Note: tqdm not found. Install with: pip install tqdm  (progress bar will be minimal)")

def main():
    start = time.perf_counter()

    # Step 1: Pick the Excel file using a native file picker dialog
    stepbar = tqdm(total=6, desc="Overall")
    root = Tk(); root.withdraw()
    input_file = askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
    )
    if not input_file:
        print("ERROR: No file selected. Exiting.")
        sys.exit(0)
    stepbar.update(1)

    # Step 2: Read the Excel file into a DataFrame
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        stepbar.close()
        print(f"ERROR: Failed to read Excel: {e}")
        sys.exit(1)
    stepbar.update(1)

    # Step 3: Validate required columns and normalize types
    df.columns = df.columns.str.strip()
    required = {"SN", "Scan Count"}
    missing = required - set(df.columns)
    if missing:
        stepbar.close()
        print(f"ERROR: Missing required column(s): {', '.join(sorted(missing))}")
        sys.exit(1)

    # Coerce Scan Count to numeric (invalid values become NaN)
    df["Scan Count"] = pd.to_numeric(df["Scan Count"], errors="coerce")
    stepbar.update(1)

    # Step 4: Deduplicate (keep the highest Scan Count per SN)
    if DETAILED_PROGRESS:
        # Slower but shows progress over groups; preserves original dtypes
        selected_idx = []
        gb = df.groupby("SN", sort=False)
        for sn, g in tqdm(gb, desc="Per-SN selection"):
            # Within each SN, pick the row with max Scan Count (NaN treated as smallest)
            g_sorted = g.sort_values("Scan Count", ascending=False, na_position="last", kind="mergesort")
            selected_idx.append(g_sorted.index[0])
        # Preserve original dtypes by selecting from the original df
        df_filtered = df.loc[selected_idx]
    else:
        # Fast, vectorized approach
        df_sorted = df.sort_values(["SN", "Scan Count"], ascending=[True, False], na_position="last")
        df_filtered = df_sorted.drop_duplicates(subset="SN", keep="first").sort_index(kind="stable")
    stepbar.update(1)

    # Step 5: Save output as a new .xlsx next to the original
    base, _ext = os.path.splitext(input_file)
    output_file = f"{base}_cleaned.xlsx"
    try:
        df_filtered.to_excel(output_file, index=False, engine="openpyxl")
    except Exception as e:
        stepbar.close()
        print(f"ERROR: Failed to write Excel: {e}")
        sys.exit(1)
    stepbar.update(1)

    # Step 6: Wrap up and report summary
    removed = len(df) - len(df_filtered)
    stepbar.update(1)
    stepbar.close()

    elapsed = time.perf_counter() - start
    print(f"OK: Cleaned file saved as: {output_file}")
    print(f"INFO: {removed} rows removed (kept highest 'Scan Count' per SN).")
    print(f"DONE in {elapsed:.2f}s")

if __name__ == "__main__":
    main()
