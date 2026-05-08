import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import os
from pathlib import Path
from datetime import date, timedelta
import xlwings as xw
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# ── Configuration ─────────────────────────────────────────────────────────────
# This folder should match where combine_orders.py saves its output
TRACKER_FOLDER = r"C:\Daily Sales Tracker Automation\Tracker"

# Gmail Configuration (Loaded from .env)
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASSWORD = os.getenv("GMAIL_PASSWORD")

# Email Recipients (Loaded from .env, split by comma)
RECIPIENTS = [email.strip() for email in os.getenv("RECIPIENTS", "").split(",") if email.strip()]
CC_RECIPIENTS = [email.strip() for email in os.getenv("CC_RECIPIENTS", "").split(",") if email.strip()]

# Subject Line Prefix
SUBJECT_PREFIX = "Daily Sales Status"

# Table extraction settings
TABLE_SHEET = "Email" 
TABLE_RANGE = "A3:H6" 
# ──────────────────────────────────────────────────────────────────────────────

def find_todays_tracker(folder: str):
    """Find the tracker file created for today's date."""
    path = Path(folder)
    today = date.today()
    
    # Match the naming convention from combine_orders.py:
    # f"Daily Sales Status - {today:%B} {today.day} Tracker.xlsx"
    filename = f"Daily Sales Status - {today:%B} {today.day} Tracker.xlsx"
    file_path = path / filename
    
    if file_path.exists():
        return file_path
    
    print(f"Error: Today's tracker file not found: {filename}")
    return None

def load_email_notes(tracker_path: Path) -> dict:
    """Load the BOL metadata written by combine_orders.py (sidecar .json file)."""
    notes_path = tracker_path.with_suffix(".json")
    if notes_path.exists():
        with open(notes_path, "r") as f:
            return json.load(f)
    print(f"Warning: Email notes file not found: {notes_path.name}")
    return {}

def format_bol_note(notes: dict) -> str:
    """
    Build the BOL disclaimer sentence and additional bullet points from the sidecar metadata.
    """
    last_tuesday_str = notes.get("last_tuesday")
    fiscal_end_str   = notes.get("fiscal_end")
    excluded_sales   = notes.get("excluded_bol_sales", 0) or 0
    
    # New fields for additional bullets
    no_bol_current = notes.get("no_bol_current_total", 0) or 0
    no_bol_past    = notes.get("no_bol_past_total", 0) or 0
    fiscal_start_str = notes.get("fiscal_start")
    prev_fiscal_end_str = notes.get("prev_fiscal_end")
    prev_month_name = notes.get("prev_fiscal_month_name", "previous month")

    if not last_tuesday_str:
        return ""

    from datetime import datetime as dt
    last_tuesday = dt.strptime(last_tuesday_str, "%Y-%m-%d").date()
    report_end   = last_tuesday - timedelta(days=1)

    cutoff_fmt = report_end.strftime("%m/%d")
    start_fmt  = last_tuesday.strftime("%m/%d/%Y")
    
    if excluded_sales >= 1_000_000:
        sales_fmt = f"${excluded_sales / 1_000_000:.1f}M"
    else:
        sales_fmt = f"${excluded_sales / 1_000:.1f}K"

    main_note = ""
    if fiscal_end_str and excluded_sales > 0:
        fiscal_end = dt.strptime(fiscal_end_str, "%Y-%m-%d").date()
        end_fmt    = fiscal_end.strftime("%m/%d/%Y")
        main_note = (
            f"* Please note that the month orders only include open orders through "
            f"{cutoff_fmt}. There is {sales_fmt} of open orders with BOLs with request "
            f"dates between {start_fmt} and {end_fmt}."
        )
    else:
        main_note = f"* Please note that the month orders only include open orders through {cutoff_fmt}."

    # Build the two new bullet points
    bullets = []
    if fiscal_start_str and fiscal_end_str:
        f_start = dt.strptime(fiscal_start_str, "%Y-%m-%d").date()
        f_end   = dt.strptime(fiscal_end_str, "%Y-%m-%d").date()
        f_start_fmt = f_start.strftime("%m/%d/%y")
        f_end_fmt   = f_end.strftime("%m/%d/%y")
        
        current_val_fmt = f"${no_bol_current / 1_000:.0f}K"
        bullets.append(f"&bull; {current_val_fmt} with request dates in this fiscal month ({f_start_fmt} &ndash; {f_end_fmt})")

    if prev_fiscal_end_str:
        p_end = dt.strptime(prev_fiscal_end_str, "%Y-%m-%d").date()
        p_end_fmt = p_end.strftime("%m/%d/%y")
        
        past_val_fmt = f"${no_bol_past / 1_000:.0f}K"
        bullets.append(f"&bull; {past_val_fmt} with request dates in this fiscal month (prior or equal to {p_end_fmt})")

    bullets_html = ""
    if bullets:
        bullets_html = "<br>Open Orders without BOLs:<br>" + "<br>".join(bullets)

    return f"{main_note}<br>{bullets_html}"

def get_table_as_html(file_path: Path, sheet_name: str, range_address: str):
    """
    Opens the Excel file, reads the specified range, and converts it to a
    styled HTML table for the email body using inline CSS.
    """
    if not sheet_name or not range_address:
        return ""
        
    print(f"Extracting table from {sheet_name}!{range_address}...")
    try:
        with xw.App(visible=False) as app:
            wb = app.books.open(str(file_path), read_only=True)
            sheet = wb.sheets[sheet_name]
            # Convert range to DataFrame
            df = sheet.range(range_address).options(pd.DataFrame, header=1, index=False).value
            wb.close()
            
        if df is None or df.empty:
            return ""

        # Format columns based on user request:
        # 2nd, 3rd, 5th columns (indices 1, 2, 4) -> Currency $
        # 4th, 6th, 8th columns (indices 3, 5, 7) -> Percentage %
        for i, col_name in enumerate(df.columns):
            if i in [1, 2, 4, 6]: # 2nd, 3rd, 5th
                df[col_name] = df[col_name].apply(lambda x: f"${x:,.0f}")
            elif i in [3, 5, 7]: # 4th, 6th, 8th
                df[col_name] = df[col_name].apply(lambda x: f"{x:.1%}" if isinstance(x, (int, float)) else x)

        # Convert DataFrame to HTML without the CSS class
        html_table = df.to_html(index=False)
        
        # Define inline styles
        table_style = "border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 10pt; width: auto; margin: 10px 0;"
        th_style = "background-color: #4472C4; color: white; padding: 8px; text-align: center; border: 1px solid #ccc;"
        td_style = "font-weight: bold, padding: 8px; border: 1px solid #ccc; text-align: center"
        
        # Manually inject inline styles into the HTML string
        html_table = html_table.replace('<table', f'<table style="{table_style}"')
        html_table = html_table.replace('<th>', f'<th style="{th_style}">')
        html_table = html_table.replace('<td>', f'<td style="{td_style}">')
        
        # Add alternating row colors (zebra striping)
        rows = html_table.split('<tr>')
        styled_rows = []
        for i, row in enumerate(rows):
            # i=0 is everything before the first <tr>
            # i=1 is the header row
            # i=2 is the first data row (even child)
            if i > 1 and i % 2 == 0: # Apply to even rows (2, 4, 6...)
                # Replace the already injected td_style with one that includes the background color
                styled_rows.append(row.replace(f'style="{td_style}"', f'style="{td_style} background-color: #f2f2f2;"'))
            else:
                styled_rows.append(row)
        
        # Reconstruct the table
        html_table = '<tr>'.join(styled_rows)
        
        return f'<div class="sales-table-container">{html_table}</div>'
    except Exception as e:
        print(f"Warning: Failed to extract table: {e}")
        return ""

def send_via_gmail(attachment_path: Path, table_html: str = "", bol_note: str = ""):
    """
    Sends an email via Gmail SMTP with the attachment, HTML table, and BOL note.
    """
    try:
        today = date.today()
        
        # Create the message
        msg = MIMEMultipart()
        msg['From'] = GMAIL_USER
        msg['To'] = ", ".join(RECIPIENTS)
        msg['Cc'] = ", ".join(CC_RECIPIENTS)
        msg['Subject'] = f"{SUBJECT_PREFIX} - {today:%B} {today.day}"
        
        bol_note_html = f"<p><i>{bol_note}</i> </p>" if bol_note else ""

        # Construct the HTML Body
        html_body = f"""
        <html>
            <body style="font-family: Calibri, sans-serif; font-size: 11pt;">
                <p>Hi team,</p>
                <p>Please find the <b>Daily Sales Status</b> tracker for {today:%B %d} attached.</p>
                {table_html}
                {bol_note_html}
                <p>Best regards,<br><b>Sales Automation Bot</b></p>
            </body>
        </html>
        """
        msg.attach(MIMEText(html_body, 'html'))
        
        # Attach the file
        print(f"Attaching file: {attachment_path.name}...")
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {attachment_path.name}",
        )
        msg.attach(part)
        
        # Send the email
        print("Connecting to Gmail SMTP server...")
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        
        all_recipients = RECIPIENTS + CC_RECIPIENTS
        server.sendmail(GMAIL_USER, all_recipients, msg.as_string())
        server.quit()
        
        print(f"Email successfully sent for: {attachment_path.name}")
        
    except Exception as e:
        print(f"Error while sending Gmail email: {e}")
        print("\nNote: If you get a login error, ensure you are using an 'App Password'.")
        print("Go to: Google Account > Security > 2-Step Verification > App Passwords.")

if __name__ == "__main__":
    # 1. Find today's tracker file
    todays_tracker = find_todays_tracker(TRACKER_FOLDER)
    
    if todays_tracker:
        print(f"Found today's file: {todays_tracker.name}")
        
        # 2. Extract table from 'Email' tab A3:H6
        table_content = get_table_as_html(todays_tracker, TABLE_SHEET, TABLE_RANGE)
        
        # 3. Load BOL note metadata written by combine_orders.py
        notes = load_email_notes(todays_tracker)
        bol_note = format_bol_note(notes)
        if bol_note:
            print(f"BOL note: {bol_note}")
        
        # 4. Send the email
        send_via_gmail(todays_tracker, table_content, bol_note)
    else:
        print(f"Could not find a tracker file for today in {TRACKER_FOLDER}.")
        print("Please run combine_orders.py first.")
