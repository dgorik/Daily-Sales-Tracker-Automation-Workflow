import json
import pandas as pd
import xlwings as xw
from datetime import date, datetime
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────
SOURCE_FOLDER  = r"C:\Daily Sales Tracker Automation\Raw Data"   # <-- folder with 3 source files
TEMPLATE_FILE  = r"C:\Daily Sales Tracker Automation\Tracker\Daily Sales Status - May BLANK Tracker.xlsx" # <-- template workbook (never overwritten)

OPEN_KEYWORD   = "Open"    # files whose name contains this word are open orders
CLOSED_KEYWORD = "Closed"  # files whose name contains this word are closed orders

OPEN_SHEET     = "Open Orders"    # exact tab name in the output file
CLOSED_SHEET   = "Closed Orders"  # exact tab name in the output file
RAW_SHEET      = "RAW"            # exact tab name for the stacked raw output
RAW_NO_BOL_SHEET = "RAW - Open Only No BOL"  # tab for open orders with no BOL

EXCLUDE_DIVISIONS = [
    "9-MEX/CA/CAR/PR/SA",
    "9 - CANADA",
    "8 - MISCELLANEOUS BILLING",
]

CURRENT_MONTH = "May"
# ──────────────────────────────────────────────────────────────────────────────

import os

OPEN_COLS   = 32   # columns A–AF
CLOSED_COLS = 42   # columns A–AP

FORMULA_START = "AW"  # first formula column to copy into RAW
FORMULA_END   = "BX"  # last formula column to copy into RAW
BA_COL_INDEX  = 4    # 0-based index of column BA within AW:BX (AW=0, AX=1 … BA=4)
BE_COL_INDEX  = 8    # 0-based index of column BE within AW:BX (AW=0, AX=1 … BE=8)
BT_COL_INDEX  = 23   # 0-based index of column BT within AW:BX (AW=0, AX=1 … BT=23)


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

        print("Calculating formulas...")
        app.calculate()

        print("Building RAW tab...")
        open_ws   = wb.sheets[OPEN_SHEET]
        closed_ws = wb.sheets[CLOSED_SHEET]

        open_last   = len(open_df)   + 1  # +1 for header row
        closed_last = len(closed_df) + 1

        closed_vals = closed_ws.range(f"{FORMULA_START}2:{FORMULA_END}{closed_last}").value or []
        open_vals   = open_ws.range(f"{FORMULA_START}2:{FORMULA_END}{open_last}").value or []

        # Load fiscal calendar to get last Tuesday
        calendar_path = Path(__file__).parent / "fiscal_calendar_2026.json"
        with open(calendar_path, "r") as f:
            fiscal_cal = json.load(f)
        
        last_tuesday_str = fiscal_cal["2026"][CURRENT_MONTH]["last_tuesday"]
        last_tuesday = datetime.strptime(last_tuesday_str, "%Y-%m-%d").date()
        fiscal_end_str = fiscal_cal["2026"][CURRENT_MONTH]["fiscal_end"]
        fiscal_end = datetime.strptime(fiscal_end_str, "%Y-%m-%d").date()
        print(f"  Filtering open orders: Main report < {last_tuesday}, Excluded [{last_tuesday} to {fiscal_end}]")

        # keep only open rows where column BA = "Yes" and column BE < last_tuesday
        # for No BOL (BA = "No"), we don't filter on column BE
        open_filtered = []
        open_no_bol_filtered = []
        excluded_bol_sales = 0.0   # sum of BT for BA=Yes rows in [last_tuesday, fiscal_end]

        for row in open_vals:
            ba_val = row[BA_COL_INDEX]
            
            if ba_val == "No":
                open_no_bol_filtered.append(row)
                continue

            if ba_val == "Yes":
                be_val = row[BE_COL_INDEX]
                if be_val is None:
                    continue
                
                # Convert be_val to date if it's a datetime object
                if isinstance(be_val, datetime):
                    be_date = be_val.date()
                elif isinstance(be_val, str):
                    try:
                        be_date = datetime.strptime(be_val, "%Y-%m-%d").date()
                    except ValueError:
                        continue
                else:
                    continue
                    
                if be_date < last_tuesday:
                    open_filtered.append(row)
                elif last_tuesday <= be_date <= fiscal_end:
                    # Accumulate gross sales for orders in the gap period
                    bt_val = row[BT_COL_INDEX]
                    if isinstance(bt_val, (int, float)):
                        excluded_bol_sales += bt_val

        combined = closed_vals + open_filtered
        print(f"  RAW rows: {len(closed_vals)} closed + {len(open_filtered)} open = {len(combined)}")
        print(f"  RAW No BOL rows: {len(open_no_bol_filtered)}")

        raw_ws = wb.sheets[RAW_SHEET]
        if combined:
            raw_ws.range("A2").value = combined

        raw_no_bol_ws = wb.sheets[RAW_NO_BOL_SHEET]
        if open_no_bol_filtered:
            raw_no_bol_ws.range("A2").value = open_no_bol_filtered

        # Write sidecar JSON so send_email.py can build the BOL note sentence
        email_notes = {
            "last_tuesday": last_tuesday.strftime("%Y-%m-%d"),
            "fiscal_end": fiscal_end_str,
            "excluded_bol_sales": excluded_bol_sales,
        }
        notes_path = output_path.with_suffix(".json")
        with open(notes_path, "w") as nf:
            json.dump(email_notes, nf, indent=2)
        print(f"  Email notes saved → {notes_path.name}")
        print(f"  Excluded BOL sales (BE > {last_tuesday}): ${excluded_bol_sales:,.0f}")

        print("Refreshing pivot tables...")
        # Format: { "Pivot Table Name": {"field": "Field Name", "value": "Filter Value"} }
        pivot_configs = {
            "Pivot Table Open Orders": {"field": "Type", "value": "Open Orders"},
            "Pivot Table Closed Orders": {"field": "Type", "value": "Actual Shipped"},
            "Current Month Bol (Take No's Only)": {"field": "Current Month Bol (Take No's Only)", "value": "No"},
            "Past Current Month Bol (Take No's Only)": {"field": "Past Current Month Bol (Take No's Only)", "value": "No"},
        }
        for sheet in wb.sheets:
            pts = sheet.api.PivotTables()
            for i in range(1, pts.Count + 1):
                pt = pts.Item(i)
                if pt.Name in pivot_configs:
                    pt.RefreshTable()
                    config = pivot_configs[pt.Name]
                    try:
                        pf = pt.PivotFields(config["field"])
                        pf.EnableMultiplePageItems = False
                        pf.CurrentPage = config["value"]
                        print(f"  Refreshed '{pt.Name}' → filtered '{config['field']}' to '{config['value']}'")
                    except Exception as e:
                        print(f"  Warning: Could not set filter on '{pt.Name}': {e}")

        wb.save(str(output_path))
        wb.close()

    print(f"\nDone. Saved: {output_path}")


if __name__ == "__main__":
    main()
