#!/usr/bin/env python3
"""
Custom bill email generator - processes PDFs from custom_bill folder.

This script:
1. Processes all PDFs in the custom_bill folder
2. Extracts bill data (house, amount, vendor, date)
3. Crops and attaches images from PDFs
4. Uses a custom month (1-12) provided by user for the email
5. Always uses the "dual" email template

Usage:
    python3 custom_bill_email.py
    
    The script will prompt you to enter:
    - Month number (1-12): 1=January, 2=February, etc.
    - Test mode (y/n): Whether to print email details without sending
"""

import sys
import os
import logging
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime
import calendar

# Add project root to path for imports
sys.path.append(str(Path(__file__).parent))
from main import process_single_file
from src.email_drafts import generate_email_drafts

# Setup logging
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("custom-bill-email")

# Suppress INFO logs from main.py
logging.getLogger("bill-pipeline").setLevel(logging.WARNING)

# Custom bill folder path
CUSTOM_BILL_FOLDER = Path(__file__).parent / "custom_bill"


def get_custom_bill_pdfs() -> List[str]:
    """Get list of PDF files from custom_bill folder."""
    if not CUSTOM_BILL_FOLDER.exists():
        return []
    return sorted([
        p.name for p in CUSTOM_BILL_FOLDER.iterdir()
        if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith(".")
    ])


def validate_month(month: int) -> None:
    """Validate that month is between 1 and 12."""
    if not isinstance(month, int) or month < 1 or month > 12:
        raise ValueError(f"Month must be an integer from 1 to 12, got: {month}")


def process_custom_bills(custom_month: int) -> List[Dict]:
    """
    Process all PDFs in custom_bill folder.
    
    Args:
        custom_month: Month number (1-12) to use for email generation
    
    Returns:
        List of processed bill dicts with keys: house_number, bill_date, bill_amount, vendor, date
    """
    validate_month(custom_month)
    
    pdf_files = get_custom_bill_pdfs()
    if not pdf_files:
        print(f"No PDFs found in {CUSTOM_BILL_FOLDER}")
        return []
    
    print(f"Found {len(pdf_files)} PDFs in custom_bill folder")
    
    processed_bills = []
    
    for filename in pdf_files:
        # Process file (extract, validate, create image) - don't move file
        result = process_single_file(filename, source_folder=str(CUSTOM_BILL_FOLDER), move_file_after=False)
        
        if not result:
            log.warning(f"Skipping {filename} - processing failed")
            continue
        
        # Add to processed bills list
        processed_bills.append({
            'house_number': result['house_number'],
            'bill_date': result['bill_date'],
            'bill_amount': result['bill_amount'],
            'vendor': result['vendor'],
            'date': result['bill_date'],  # For image lookup
        })
    
    return processed_bills


def generate_custom_email(custom_month: int, test_mode: bool = False) -> None:
    """
    Generate email draft for custom bills using specified month.
    
    Args:
        custom_month: Month number (1-12) to use for email generation
        test_mode: If True, print email details instead of creating drafts
    """
    # Process bills
    processed_bills = process_custom_bills(custom_month)
    
    if not processed_bills:
        print("No bills processed, no email to generate")
        return
    
    # Use the existing generate_email_drafts function with custom month
    generate_email_drafts(processed_bills, test_mode=test_mode, custom_month=custom_month)


def main():
    """Main entry point for custom bill email generation."""
    print("Custom Bill Email Generator")
    print("=" * 50)
    
    # Prompt for month
    while True:
        try:
            month_input = input("\nEnter month number (1-12): ").strip()
            custom_month = int(month_input)
            validate_month(custom_month)
            break
        except ValueError as e:
            print(f"Error: {e}")
            print("Please enter a number from 1 to 12")
            continue
    
    # Prompt for test mode
    while True:
        test_input = input("Run in test mode? (y/n): ").strip().lower()
        if test_input in ['y', 'yes']:
            test_mode = True
            break
        elif test_input in ['n', 'no']:
            test_mode = False
            break
        else:
            print("Please enter 'y' or 'n'")
            continue
    
    mode_str = "TEST MODE" if test_mode else "LIVE MODE"
    print(f"\nStarting email generation [{mode_str}] for {calendar.month_name[custom_month]}...")
    print("=" * 50)
    
    try:
        generate_custom_email(custom_month, test_mode=test_mode)
    except Exception as e:
        log.error(f"Custom bill email generation failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

