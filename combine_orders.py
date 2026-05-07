import pandas as pd
import xlwings as xw
from datetime import date

# ── Configuration ─────────────────────────────────────────────────────────────
SOURCE_FOLDER  = r"C:\Daily Sales Tracker Automation\Raw Data"   # <-- folder with 3 source files
TEMPLATE_FILE  = r"C:\Daily Sales Tracker Automation\Tracker\Daily Sales Status - May BLANK Tracker.xlsx" # <-- template workbook (never overwritten)

OPEN_KEYWORD   = "Open"    # files whose name contains this word are open orders
CLOSED_KEYWORD = "Closed"  # files whose name contains this word are closed orders

OPEN_SHEET     = "Open Orders"    # exact tab name in the output file
CLOSED_SHEET   = "Closed Orders"  # exact tab name in the output file

EXCLUDE_DIVISIONS = [
    "9-MEX/CA/CAR/PR/SA",
    "9 - CANADA",
    "8 - MISCELLANEOUS BILLING",
]
# ──────────────────────────────────────────────────────────────────────────────

import os
from pathlib import Path

OPEN_COLS   = 32   # columns A–AF
CLOSED_COLS = 42   # columns A–AP


def find_source_files(folder: str):
    """Return (list_of_open_files, list_of_closed_files) from the source folder."""
    csv_files    = list(Path(folder).glob("*.csv"))
    open_files   = [f for f in csv_files if OPEN_KEYWORD.lower()   in f.stem.lower()]
    closed_files = [f for f in csv_files if CLOSED_KEYWORD.lower() in f.stem.lower()]

    if len(open_files) == 0:
        raise FileNotFoundError(f"No files containing '{OPEN_KEYWORD}' found in {folder}")
    if len(closed_files) == 0:
        raise FileNotFoundError(f"No files containing '{CLOSED_KEYWORD}' found in {folder}")

    print(f"Open-order files found   : {[f.name for f in open_files]}")
    print(f"Closed-order files found : {[f.name for f in closed_files]}")
    return open_files, closed_files


def read_and_stack(file_list: list, num_cols: int) -> pd.DataFrame:
    """Read CSV files, keep only the first num_cols columns, and stack."""
    frames = []
    for f in file_list:
        df = pd.read_csv(f, usecols=range(num_cols))
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def write_to_sheet(ws, df: pd.DataFrame):
    """Write DataFrame rows starting from A2. Template is always clean on open."""
    ws.range("A2").value = df.values.tolist()
    print(f"  Written {len(df)} rows to '{ws.name}'")


def main():
    open_files, closed_files = find_source_files(SOURCE_FOLDER)

    print("\nStacking open orders...")
    open_df = read_and_stack(open_files, OPEN_COLS)

    print("Stacking closed orders...")
    closed_df = read_and_stack(closed_files, CLOSED_COLS)

    print("Filtering divisions...")
    open_df   = open_df[~open_df["Division"].isin(EXCLUDE_DIVISIONS)]
    closed_df = closed_df[~closed_df["DIVISION"].isin(EXCLUDE_DIVISIONS)]
    print(f"  Open rows after filter   : {len(open_df)}")
    print(f"  Closed rows after filter : {len(closed_df)}")

    open_df["Bill Of Lading No"] = open_df["Bill Of Lading No"].apply(
        lambda x: "No" if pd.isna(x) or str(x).strip() == "" else "Yes"
    )

    template_path = Path(TEMPLATE_FILE)
    today = date.today()
    output_name = f"Daily Sales Status - {today:%B} {today.day} Tracker.xlsx"
    output_path = template_path.parent / output_name

    print(f"\nLoading template: {template_path}")
    with xw.App(visible=False) as app:
        wb = app.books.open(str(template_path))

        print("Writing data...")
        write_to_sheet(wb.sheets[OPEN_SHEET],   open_df)
        write_to_sheet(wb.sheets[CLOSED_SHEET], closed_df)

        wb.save(str(output_path))
        wb.close()

    print(f"\nDone. Saved: {output_path}")


if __name__ == "__main__":
    main()
