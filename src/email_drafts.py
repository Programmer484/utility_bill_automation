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
        
        return (
            f"Hi everyone\n"
            f"\n"
            f"Attached are last month's utilities bills.\n"
            f"\n"
            f"The {rent_date} rent & utilities\n"
            f"${base_rent:.0f} + {utility_share}%*(${enmax_amount:.2f} [Water&Waste] + ${atco_amount:.2f} [Atco]) = ${final_amount:.2f}\n"
            f"\n"
            f"Thanks,\n"
            f"Linda"
        )

    else:
        # Template for houses with single vendor (ENMAX only)
        return (
            f"Hi everyone\n"
            f"\n"
            f"Attached are last month's utilities bills.\n"
            f"\n"
            f"The {rent_date} rent & utilities\n"
            f"${base_rent:.0f} + {utility_share}%*(${total_utilities:.2f}) = ${final_amount:.2f}\n"
            f"\n"
            f"Thanks,\n"
            f"Linda"
        )


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
        
        # Create expected image filename (matches bill naming: {house} {YYYY-MM} {vendor}.png)
        date_formatted = bill_date[:7]  # Extract YYYY-MM from YYYY-MM-DD
        img_name = f"{house} {date_formatted} {vendor}.png"
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
    
    # Format rent date (month after bill month, with two-digit day)
    try:
        from datetime import datetime, timedelta
        import calendar
        
        # Parse the bill month and add one month
        bill_dt = datetime.strptime(month_date, "%Y-%m-%d")
        
        # Add approximately one month (handles year rollover automatically)
        next_month_dt = bill_dt.replace(day=1) + timedelta(days=32)
        next_month_dt = next_month_dt.replace(day=1)  # First of next month
        
        # Format as "Month 01" (first of the month, two-digit day)
        rent_date = next_month_dt.strftime("%B") + " 01"
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

def _group_bills_by_house(processed_bills: List[Dict], custom_month: int = None) -> Dict[str, tuple]:
    """
    Group bills by house and determine the month to use for each house.
    
    Args:
        processed_bills: List of bill dicts with keys: house_number, bill_date, bill_amount, vendor
        custom_month: If provided (1-12), use this month instead of deriving from bill dates
    
    Returns:
        Dict mapping house -> (month_date, bills_list)
        where month_date is ISO date string and bills_list is list of bill dicts
    """
    from datetime import datetime
    from collections import defaultdict
    import calendar as cal
    
    if custom_month is not None:
        # Custom month mode: group by house only, use custom month for all
        bills_by_house = defaultdict(list)
        
        for bill in processed_bills:
            house = str(bill['house_number'])
            bill_date = bill['bill_date']
            amount = bill.get('bill_amount')
            vendor = bill.get('vendor')
            
            if not bill_date or amount is None or amount == '':
                continue
            
            bills_by_house[house].append({
                'vendor': vendor,
                'amount': float(amount),
                'date': bill_date
            })
        
        # Create month-end date for custom month
        year = datetime.now().year
        last_day = cal.monthrange(year, custom_month)[1]
        custom_month_date = f"{year:04d}-{custom_month:02d}-{last_day:02d}"
        
        # Return dict with same month for all houses
        return {house: (custom_month_date, bills) for house, bills in bills_by_house.items()}
    
    else:
        # Normal mode: group by house and month, use latest month
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
        
        # Return dict with latest month for each house
        result = {}
        for house, months_data in bills_by_house_month.items():
            latest_month = max(months_data.keys())
            result[house] = (latest_month, months_data[latest_month])
        
        return result


def _send_or_print_draft(msg: MIMEMultipart, house: str, test_mode: bool) -> bool:
    """
    Send draft via IMAP or print details in test mode.
    
    Args:
        msg: Email message to send
        house: House number for logging
        test_mode: If True, print details instead of sending
    
    Returns:
        True if successful, False otherwise
    """
    if test_mode:
        # Test mode - just print what would be sent
        print(f"TEST: Would create email draft for {house}")
        print(f"  Subject: {msg['Subject']}")
        print(f"  To: {msg['To']}")
        
        # List attachments
        attach_names = list_attachments(msg)
        print(f"  Attachments ({len(attach_names)}): {attach_names}")
        print()
        return True
    else:
        # Real mode - save draft to Yahoo via IMAP
        saved = save_draft_via_imap(msg)
        if not saved:
            print(f"Failed to save draft for {house} via IMAP")
        return saved


def generate_email_drafts(processed_bills: List[Dict], test_mode: bool = None, custom_month: int = None) -> None:
    """
    Generate email drafts from freshly processed bill data.
    
    Args:
        processed_bills: List of bill dicts with keys: house_number, bill_date, bill_amount, vendor
        test_mode: If True, print email details instead of creating drafts. If None, reads from config.
        custom_month: If provided (1-12), use this month instead of deriving from bill dates
    """
    # Check if we're in test/dry run mode
    if test_mode is None:
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
        # Group bills by house and determine month for each
        bills_by_house = _group_bills_by_house(processed_bills, custom_month)
        
        # Print status message
        if custom_month is not None:
            import calendar as cal
            print(f"Generating email drafts for {len(bills_by_house)} houses using custom month {custom_month} ({cal.month_name[custom_month]})...")
        else:
            print(f"Generating email drafts for {len(bills_by_house)} houses...")
        
        # Generate emails for each house
        for house, (month_date, bills_for_month) in bills_by_house.items():
            # Calculate total utilities
            total_utilities = sum(b['amount'] for b in bills_for_month)
            
            # Create vendor breakdown for dual vendor emails
            vendor_breakdown = {b['vendor']: b['amount'] for b in bills_for_month}
            
            print(f"Processing {house} for {month_date} (utilities: ${total_utilities:.2f})")
            
            try:
                # Create email draft (will enforce required vendors when applicable)
                msg = create_email_draft(house, month_date, total_utilities, vendor_breakdown, bills_for_month)
                
                # Send or print the draft
                _send_or_print_draft(msg, house, test_mode)
                        
            except Exception as e:
                print(f"Skipping {house} for {month_date}: {e}")
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
