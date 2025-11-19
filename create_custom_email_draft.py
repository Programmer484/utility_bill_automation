"""
Create email drafts in dual or single format with custom bill attachments.

This script allows you to create email drafts and attach all files from the 
custom_bill folder. You can choose between dual vendor format (showing vendor 
breakdown) or single vendor format (showing total only).
"""

import os
import sys
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import Dict, List, Optional

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Add src to path for imports
sys.path.append(str(Path(__file__).parent / "src"))
from email_drafts import save_draft_via_imap, list_attachments, YAHOO_USER, YAHOO_APP_PASSWORD, get_email_template
from config import get_tenant_data

# Custom bill folder path
CUSTOM_BILL_FOLDER = Path(__file__).parent / "custom_bill"


def get_files_from_custom_bill_folder() -> List[Path]:
    """Get all files from the custom_bill folder."""
    # Create folder if it doesn't exist
    CUSTOM_BILL_FOLDER.mkdir(exist_ok=True)
    
    files = []
    for item in CUSTOM_BILL_FOLDER.iterdir():
        if item.is_file():
            files.append(item)
    
    return sorted(files)


def attach_file_to_email(msg: MIMEMultipart, file_path: Path) -> bool:
    """Attach a file to the email message."""
    try:
        with open(file_path, 'rb') as f:
            # Determine content type based on file extension
            ext = file_path.suffix.lower()
            if ext == '.pdf':
                maintype = 'application'
                subtype = 'pdf'
            elif ext in ['.jpg', '.jpeg', '.png', '.gif']:
                maintype = 'image'
                subtype = ext[1:] if ext != '.jpg' else 'jpeg'
            elif ext in ['.xlsx', '.xls']:
                maintype = 'application'
                subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if ext == '.xlsx' else 'vnd.ms-excel'
            else:
                # Generic binary attachment
                maintype = 'application'
                subtype = 'octet-stream'
            
            part = MIMEBase(maintype, subtype)
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename={file_path.name}'
            )
            msg.attach(part)
            return True
    except Exception as e:
        print(f"Warning: Could not attach file {file_path.name}: {e}")
        return False


def create_custom_email_draft(
    house_number: str,
    total_utilities: Optional[float] = None,
    vendor_breakdown: Optional[Dict[str, float]] = None,
    rent_date: Optional[str] = None,
    subject: Optional[str] = None,
    template_type: str = "dual"
) -> MIMEMultipart:
    """
    Create an email draft with custom bill attachments.
    
    Excel is the only source of truth for tenant data. All tenant information
    (name, base rent, utility share, email) must exist in Excel.
    
    Args:
        house_number: House number (required - used to look up all tenant data from Excel)
        total_utilities: Total utilities amount (required for single format)
        vendor_breakdown: Dict of vendor names to amounts (required for dual format)
        rent_date: Rent date (e.g., "November 1")
        subject: Email subject line
        template_type: "dual" or "single"
    
    Returns:
        MIMEMultipart message ready to be saved as draft
    
    Raises:
        KeyError: If house_number not found in Excel
        ValueError: If required data is missing
    """
    # Load tenant data from Excel - Excel is the only source of truth
    tenant_data = get_tenant_data(str(house_number))
    
    # Extract all required fields from Excel (no defaults)
    tenant_name = tenant_data["tenant_name"]
    base_rent = float(tenant_data["base_rent"])
    utility_share_percent = int(tenant_data["utility_share_percent"])
    recipient_email = tenant_data["email"]
    
    # Calculate utilities and final amount
    if template_type == "dual" and vendor_breakdown:
        total_utilities = sum(vendor_breakdown.values())
    elif total_utilities is None:
        raise ValueError("total_utilities is required for single format")
    
    utility_share_amount = (utility_share_percent / 100) * total_utilities
    final_amount = base_rent + utility_share_amount
    
    # Require rent_date to be provided
    if rent_date is None:
        raise ValueError("rent_date is required")
    
    # Map template_type to email_drafts.py format ("dual" -> "dual_vendor", "single" -> "single_vendor")
    email_template_type = "dual_vendor" if template_type == "dual" else "single_vendor"
    
    # Generate email content using template from email_drafts.py
    email_content = get_email_template(
        tenant_name=tenant_name,
        rent_date=rent_date,
        base_rent=base_rent,
        utility_share=utility_share_percent,
        total_utilities=total_utilities,
        final_amount=final_amount,
        template_type=email_template_type,
        vendor_breakdown=vendor_breakdown
    )
    
    # Create email message
    msg = MIMEMultipart()
    if not YAHOO_USER:
        raise ValueError("YAHOO_USER environment variable is required")
    msg['From'] = YAHOO_USER
    msg['To'] = recipient_email
    
    # Set subject
    if subject is None:
        # Extract month from rent_date for subject
        month = rent_date.split()[0]
        msg['Subject'] = f"{month} utility bill"
    else:
        msg['Subject'] = subject
    
    # Add email body
    msg.attach(MIMEText(email_content, 'plain'))
    
    # Attach all files from custom_bill folder
    custom_files = get_files_from_custom_bill_folder()
    if custom_files:
        print(f"Attaching {len(custom_files)} file(s) from custom_bill folder:")
        for file_path in custom_files:
            if attach_file_to_email(msg, file_path):
                print(f"  ✓ {file_path.name}")
            else:
                print(f"  ✗ {file_path.name} (failed)")
    else:
        print("Warning: No files found in custom_bill folder")
    
    return msg



def main():
    create_custom_email_draft(house_number="1705")
    

if __name__ == "__main__":
    main()

