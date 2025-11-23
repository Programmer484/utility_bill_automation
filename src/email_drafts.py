import os
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
from config import get_excel_path, get_excel_data_sheet, get_images_folder, bill_date_to_month_end, get_config
from excel import get_tenant_data

# Email configuration
YAHOO_USER = os.getenv("YAHOO_USER")
YAHOO_APP_PASSWORD = os.getenv("YAHOO_APP_PASSWORD")

HOUSE_POLICIES = {
    "1705": {"required_vendors": ["ENMAX", "ATCO"], "template": "dual_vendor"},
    "1707": {"required_vendors": ["ENMAX", "ATCO"], "template": "dual_vendor"},
    "default": {"required_vendors": ["ENMAX"], "template": "single_vendor"},
}

def get_email_template(tenant_name: str, rent_date: str, base_rent: float, 
                      utility_share: int, total_utilities: float, final_amount: float, 
                      template_type: str = "single_vendor", vendor_breakdown: Dict = None) -> str:
    """Generate email template with placeholders filled in."""
    
    if template_type == "dual_vendor" and vendor_breakdown:
        # Template for houses with multiple vendors (ENMAX + ATCO)
        enmax_amount = vendor_breakdown.get("ENMAX", 0)
        atco_amount = vendor_breakdown.get("ATCO", 0)
        
        return f"""Hi everyone

            Attached are last month's utilities bills.

            The {rent_date} rent & utilities
            ${base_rent:.0f} + {utility_share}%*(${enmax_amount:.2f} [Water&Waste] + ${atco_amount:.2f} [Atco]) = ${final_amount:.2f}

            Thanks,
            Linda"""

    else:
        # Template for houses with single vendor (ENMAX only)
        return f"""Hi everyone

        Attached are last month's utilities bills.

        The {rent_date} rent & utilities
        ${base_rent:.0f} + {utility_share}%*(${total_utilities:.2f}) = ${final_amount:.2f}

        Thanks,
        Linda"""


def _get_house_policy(house: str) -> dict:
    return HOUSE_POLICIES.get(str(house), HOUSE_POLICIES["default"])

def find_house_utility_images(house: str, month_date: str, bills_data: List[Dict]) -> List[str]:
    """Find utility images for a specific house from bill data."""
    image_folder = Path(get_images_folder())
    image_paths = []
    
    # Get house policy for vendor requirements
    policy = _get_house_policy(house)
    required_vendors = policy.get("required_vendors", [])
    
    vendors_found = []
    for bill in bills_data:
        vendor = bill['vendor']
        bill_date = bill['date']
        
        # Create expected image filename
        img_name = f"{house}_{bill_date}_{vendor}.png"
        img_path = image_folder / img_name
        
        if img_path.exists():
            image_paths.append(str(img_path))
            vendors_found.append(vendor)
    
    # Check if required vendors are present
    if required_vendors:
        missing_vendors = [v for v in required_vendors if v not in vendors_found]
        if missing_vendors:
            raise ValueError(f"Missing required vendors for house {house} {month_date}: {missing_vendors}")
    
    if not image_paths:
        raise ValueError(f"No utility bill images found for house {house} {month_date}")
    
    return sorted(image_paths)  # Return in consistent order

def create_email_draft(house: str, month_date: str, total_utilities: float, vendor_breakdown: Dict = None, bills_data: List[Dict] = None) -> MIMEMultipart:
    """Create email draft for a specific house."""
    
    # Find utility images directly from fresh bill data instead of Excel
    if bills_data:
        image_paths = find_house_utility_images(house, month_date, bills_data)

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
        
        # Format as "Month 1" (first of the month)
        rent_date = next_month_dt.strftime("%B 1")
    except:
        rent_date = month_date
    
    # Get house policy to determine template type
    house_policy = _get_house_policy(house)
    template_type = house_policy.get("template", "single_vendor")
    
    # Generate email content
    email_content = get_email_template(
        tenant_name, rent_date, base_rent, utility_share_percent, 
        total_utilities, final_amount, template_type, vendor_breakdown
    )
    
    # Create email message
    msg = MIMEMultipart()
    msg['From'] = YAHOO_USER
    msg['To'] = f"tenant_{house}@example.com"  # Placeholder email
    # Subject should be one month before rent date (bill month, not rent month)
    try:
        from datetime import datetime, timedelta
        # Parse rent date to get previous month for subject
        rent_month_dt = datetime.strptime(rent_date.split()[0] + " 2025", "%B %Y")
        # Subtract one month for bill month
        bill_month_dt = rent_month_dt.replace(day=1) - timedelta(days=1)
        bill_month_dt = bill_month_dt.replace(day=1)  # First of previous month
        bill_month = bill_month_dt.strftime("%B")
        msg['Subject'] = f"{bill_month} utility bill"
    except:
        # Fallback: just use rent month if parsing fails
        month_only = rent_date.split()[0]
        msg['Subject'] = f"{month_only} utility bill"
    
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

def generate_email_drafts(processed_bills: List[Dict]) -> None:
    """
    Generate email drafts from freshly processed bill data.
    
    Args:
        processed_bills: List of bill dicts with keys: house_number, bill_date, bill_amount, vendor
    """
    # Check if we're in test/dry run mode
    try:
        test_mode = get_config('test_email_drafts')
        if test_mode:
            print("TEST MODE: Printing email drafts instead of creating them")
    except Exception:
        test_mode = False
    
    if not test_mode and (not YAHOO_USER or not YAHOO_APP_PASSWORD):
        print("Error: YAHOO_USER and YAHOO_APP_PASSWORD environment variables required")
        return
    
    if not processed_bills:
        print("No bills to process for email generation")
        return
    
    try:
        # Group bills by house and month
        from datetime import datetime
        from collections import defaultdict
        
        # Structure: {house: {month_key: [bills]}}
        bills_by_house_month = defaultdict(lambda: defaultdict(list))
        
        for bill in processed_bills:
            house = str(bill['house_number'])
            bill_date = bill['bill_date']
            amount = bill.get('bill_amount')
            vendor = bill.get('vendor')
            
            if not bill_date or amount is None or amount == '':
                continue
                
            # Convert to month-end key
            month_key = bill_date_to_month_end(bill_date)
            bills_by_house_month[house][month_key].append({
                'vendor': vendor,
                'amount': float(amount),
                'date': bill_date
            })
        
        print(f"Generating email drafts for {len(bills_by_house_month)} houses...")
        
        # Generate emails for each house's latest month
        for house, months_data in bills_by_house_month.items():
            # Get the latest month for this house
            latest_month = max(months_data.keys())
            bills_for_month = months_data[latest_month]
            
            # Calculate total utilities
            total_utilities = sum(b['amount'] for b in bills_for_month)
            
            print(f"Processing {house} for {latest_month} (utilities: ${total_utilities:.2f})")
            
            try:
                # Create vendor breakdown for dual vendor emails
                vendor_breakdown = {b['vendor']: b['amount'] for b in bills_for_month}
                
                # Create email draft (will enforce required vendors when applicable)
                msg = create_email_draft(house, latest_month, total_utilities, vendor_breakdown, bills_for_month)
                
                if test_mode:
                    # Test mode - just print what would be sent
                    print(f"TEST: Would create email draft for {house}")
                    print(f"  Subject: {msg['Subject']}")
                    print(f"  To: {msg['To']}")
                    
                    # List attachments
                    attach_names = list_attachments(msg)
                    print(f"  Attachments ({len(attach_names)}): {attach_names}")
                    print()
                else:
                    # Real mode - save draft to Yahoo via IMAP
                    saved = save_draft_via_imap(msg)
                    if not saved:
                        print(f"Failed to save draft for {house} via IMAP")
                        
            except Exception as e:
                print(f"Skipping {house} for {latest_month}: {e}")
                continue
        
        if test_mode:
            print("TEST MODE: Email drafts printed successfully!")
        else:
            print("Email drafts generated successfully!")
        
    except Exception as e:
        print(f"Error generating email drafts: {e}")


def main():
    """CLI entry point - generates emails from Excel data for testing."""
    print("Email generation now only works from processed bills data.")
    print("Run the main pipeline instead: python3 main.py")

if __name__ == "__main__":
    main()
