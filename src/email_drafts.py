import os
import smtplib
import imaplib
import time
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import Dict, List

# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv not installed, rely on system env vars

import sys
from pathlib import Path

# Add parent directory to path to import config
sys.path.append(str(Path(__file__).parent.parent))
# Add current directory to path for local imports
sys.path.append(str(Path(__file__).parent))
from config import get_excel_path, get_excel_data_sheet, get_images_folder, bill_date_to_month_end
from excel import latest_month_totals, get_tenant_data

# Email configuration
YAHOO_USER = os.getenv("YAHOO_USER")
YAHOO_APP_PASSWORD = os.getenv("YAHOO_APP_PASSWORD")

HOUSE_POLICIES = {
    "1705": {"required_vendors": ["ENMAX", "ATCO"], "template": "dual_vendor"},
    "1707": {"required_vendors": ["ENMAX", "ATCO"], "template": "dual_vendor"},
    "default": {"required_vendors": ["ENMAX"], "template": "single_vendor"},
}

def _get_house_policy(house: str) -> dict:
    return HOUSE_POLICIES.get(str(house), HOUSE_POLICIES["default"])

def get_email_template(tenant_name: str, rent_date: str, base_rent: float, 
                      utility_share: int, total_utilities: float, final_amount: float, 
                      template_type: str = "single_vendor") -> str:
    """Generate email template with placeholders filled in."""
    
    if template_type == "dual_vendor":
        # Template for houses with multiple vendors (ENMAX + ATCO)
        return f"""Hi everyone,

Attached are last month's utility bills.

The {rent_date} rent amount is:
${base_rent} + {utility_share}% * ${total_utilities:.2f} = ${final_amount:.2f}

Best regards."""
    else:
        # Template for houses with single vendor (ENMAX only)
        return f"""Hi everyone,

Attached is last month's utilities bill.

The {rent_date} rent amount is:
${base_rent} + {utility_share}% * ${total_utilities:.2f} = ${final_amount:.2f}

Best regards."""

def find_utility_images(house: str, month_date: str) -> List[str]:
    """Resolve utility bill images for a house and month, enforcing policy.

    - Uses Data sheet to ensure all required vendors exist for the exact month.
    - Returns image paths (one per vendor) in a stable order.
    - Raises ValueError if a required vendor is missing for that month.
    """
    policy = _get_house_policy(house)
    required_vendors = policy.get("required_vendors")

    try:
        df = pd.read_excel(get_excel_path(), sheet_name=get_excel_data_sheet())
    except Exception as e:
        print(f"Error reading Data sheet: {e}")
        return []

    df.columns = [c.strip().lower() for c in df.columns]
    needed = {"file", "house_number", "bill_date", "vendor"}
    missing = needed - set(df.columns)
    if missing:
        print(f"Data sheet missing columns: {sorted(missing)}")
        return []

    # Normalize
    df["house_number"] = df["house_number"].astype(str)
    df["vendor"] = df["vendor"].astype(str).str.strip().str.upper()
    df["bill_date"] = pd.to_datetime(df["bill_date"], errors="coerce")
    
    # Compute month-end for each bill_date and filter
    df["_month_key"] = df["bill_date"].apply(bill_date_to_month_end)
    df_f = df[(df["house_number"] == str(house)) & (df["_month_key"] == str(month_date))]
    if df_f.empty:
        print(f"No data rows for house {house} in month {month_date}")
        return []

    # For each vendor present, pick the row with latest bill_date
    image_folder = Path(get_images_folder())
    vendor_to_image: dict[str, str] = {}
    for vendor, grp in df_f.groupby("vendor"):
        grp_sorted = grp.sort_values("bill_date")
        row = grp_sorted.iloc[-1]
        bill_dt = row["bill_date"]
        if pd.isna(bill_dt):
            # Fallback: try to derive bill_date from file name if needed (skip if not possible)
            continue
        iso_date = str(bill_dt.date())
        img_name = f"{house}_{iso_date}_{vendor}.png"
        img_path = image_folder / img_name
        if img_path.exists():
            vendor_to_image[vendor] = str(img_path)

    if required_vendors:
        missing_required = [v for v in required_vendors if v not in vendor_to_image]
        if missing_required:
            raise ValueError(
                f"Missing required vendors for house {house} {month_date}: {missing_required}"
            )

        # return in required order
        return [vendor_to_image[v] for v in required_vendors if v in vendor_to_image]

    # default: return all images we found, sorted by vendor name for stability
    return [vendor_to_image[v] for v in sorted(vendor_to_image.keys())]

def create_email_draft(house: str, month_date: str, total_utilities: float) -> MIMEMultipart:
    """Create email draft for a specific house."""
    # First, check if we can find the required utility images
    # This will raise ValueError if required vendors are missing
    image_paths = find_utility_images(house, month_date)
    
    # Additional check: ensure we actually have some images to attach
    if not image_paths:
        raise ValueError(f"No utility bill images found for house {house} {month_date}")
    
    print(f"Found {len(image_paths)} images for {house}: {[Path(p).name for p in image_paths]}")
    
    # Get tenant data for this house
    tenant_data = get_tenant_data(str(house))
    tenant_name = tenant_data["tenant_name"]
    base_rent = tenant_data["base_rent"]
    utility_share_percent = tenant_data["utility_share_percent"]
    
    # Calculate final amount
    utility_share_amount = (utility_share_percent / 100) * total_utilities
    final_amount = base_rent + utility_share_amount
    
    # Format rent date (month after bill month, without day)
    try:
        from datetime import datetime, timedelta
        import calendar
        
        # Parse the bill month and add one month
        bill_dt = datetime.strptime(month_date, "%Y-%m-%d")
        
        # Add approximately one month (handles year rollover automatically)
        next_month_dt = bill_dt.replace(day=1) + timedelta(days=32)
        next_month_dt = next_month_dt.replace(day=1)  # First of next month
        
        # Format as "Month YYYY" (no day)
        rent_date = next_month_dt.strftime("%B %Y")
    except:
        rent_date = month_date
    
    # Get house policy to determine template type
    house_policy = _get_house_policy(house)
    template_type = house_policy.get("template", "single_vendor")
    
    # Generate email content
    email_content = get_email_template(
        tenant_name, rent_date, base_rent, utility_share_percent, 
        total_utilities, final_amount, template_type
    )
    
    # Create email message
    msg = MIMEMultipart()
    msg['From'] = YAHOO_USER
    msg['To'] = f"tenant_{house}@example.com"  # Placeholder email
    msg['Subject'] = f"Rent Statement - {house} - {rent_date}"
    
    # Add email body
    msg.attach(MIMEText(email_content, 'plain'))
    
    # Attach utility bill images for the month
    for image_path in image_paths:
        try:
            if os.path.exists(image_path):
                with open(image_path, 'rb') as img:
                    mime_img = MIMEImage(img.read())
                    filename = Path(image_path).name
                    mime_img.add_header('Content-Disposition', 'attachment', filename=filename)
                    msg.attach(mime_img)
                    print(f"Attached: {filename}")
        except Exception as e:
            print(f"Warning: Could not attach image {image_path} for {house}: {e}")
    
    return msg

def list_attachments(msg: MIMEMultipart) -> List[str]:
    names: List[str] = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        filename = part.get_filename()
        if filename:
            names.append(filename)
    return names

def dry_run_print() -> None:
    """Print what would be drafted (no IMAP), using Data-based latest-month totals."""
    df_latest = latest_month_totals(get_excel_path(), get_excel_data_sheet())
    if df_latest.empty:
        print("No data found in Data sheet")
        return
    print(f"Latest-month rows: {len(df_latest)}")
    for _, row in df_latest.iterrows():
        house = str(row['house_number'])
        month_date = str(row['latest_month'].date())
        total_utilities = float(row['total'])
        print(f"--- Dry-run: house={house} month={month_date} total=${total_utilities:.2f}")
        try:
            msg = create_email_draft(house, month_date, total_utilities)
            attach_names = list_attachments(msg)
            print(f"Would draft. Attachments ({len(attach_names)}): {attach_names}")
        except Exception as e:
            print(f"SKIP {house} {month_date}: {e}")

def save_draft_via_imap(msg: MIMEMultipart) -> bool:
    try:
        imap = imaplib.IMAP4_SSL("imap.mail.yahoo.com", 993)
        imap.login(YAHOO_USER, YAHOO_APP_PASSWORD)

        # Find the mailbox that has the \Drafts special-use flag
        drafts_mailbox = None
        typ, boxes = imap.list()
        if typ == 'OK' and boxes:
            for line in boxes:
                decoded = line.decode('utf-8', errors='ignore')
                if '\\Drafts' in decoded:
                    # last quoted token is the mailbox name
                    drafts_mailbox = decoded.split(' "/" ')[-1].strip().strip('"')
                    break

        if not drafts_mailbox:
            drafts_mailbox = '[Yahoo]/Drafts'  # common Yahoo default; fallback if detection fails

        flags = r'(\Draft)'
        date_time = imaplib.Time2Internaldate(time.time())
        typ, resp = imap.append(drafts_mailbox, flags, date_time, msg.as_bytes())
        imap.logout()
        if typ == 'OK':
            print("Draft saved to Yahoo Drafts mailbox")
            return True
        print(f"IMAP append returned non-OK: {(typ, resp)}")
        return False
    except Exception as e:
        print(f"Error saving draft via IMAP: {e}")
        return False

def generate_email_drafts(dry_run: bool = False) -> None:
    """
    Generate email drafts for all houses using latest month data.
    
    Args:
        dry_run: If True, only print what would be generated without creating actual drafts
    """
    if dry_run:
        dry_run_print()
        return
        
    if not YAHOO_USER or not YAHOO_APP_PASSWORD:
        print("Error: YAHOO_USER and YAHOO_APP_PASSWORD environment variables required")
        return
    
    try:
        # Compute latest-month totals directly from the Data sheet
        df_latest = latest_month_totals(get_excel_path(), get_excel_data_sheet())
        if df_latest.empty:
            print("No data found in Data sheet")
            return

        print(f"Generating email drafts for {len(df_latest)} houses...")

        for _, row in df_latest.iterrows():
            house = str(row['house_number'])
            month_date = str(row['latest_month'].date())
            total_utilities = float(row['total'])
            
            print(f"Processing {house} for {month_date} (utilities: ${total_utilities:.2f})")
            
            try:
                # Create email draft (will enforce required vendors when applicable)
                msg = create_email_draft(house, month_date, total_utilities)
            except Exception as e:
                print(f"Skipping {house} for {month_date}: {e}")
                continue
            
            # Save draft to Yahoo via IMAP
            saved = save_draft_via_imap(msg)
            if not saved:
                print(f"Failed to save draft for {house} via IMAP")
        
        print("Email drafts generated successfully!")
        
    except Exception as e:
        print(f"Error generating email drafts: {e}")

def main():
    """CLI entry point - checks for dry run environment variable."""
    dry_run = os.getenv("EMAIL_DRAFTS_DRY_RUN", "").lower() in {"1","true","yes"}
    generate_email_drafts(dry_run=dry_run)

if __name__ == "__main__":
    main()
