"""
Bill extraction functionality - vendor-specific PDF data extraction.
"""

import os
import re
import sys
from pathlib import Path

# Add parent directory to path to import config
sys.path.append(str(Path(__file__).parent.parent))
from config import get_house_numbers
from pypdf import PdfReader

# -------------------- Common utilities --------------------

def _extract_first_page_text(path: str) -> str:
    """Extract text from the first page of a PDF."""
    reader = PdfReader(path)
    return reader.pages[0].extract_text() or ""

# -------------------- ENMAX extraction --------------------

MONTHS = {
    "january":"01","february":"02","march":"03","april":"04","may":"05","june":"06",
    "july":"07","august":"08","september":"09","october":"10","november":"11","december":"12"
}

def make_service_address_regex(houses):
    """Create regex to match service address with house numbers."""
    alts = "|".join(sorted((re.escape(h) for h in houses), key=len, reverse=True))
    pattern = rf"SERVICE\s*ADDRESS[^:\n]{{0,80}}:\s*({alts})"
    return re.compile(pattern, re.IGNORECASE)

def extract_enmax_from_pdf(pdf_file: str, folder: str = None) -> dict:
    """
    Extract bill data from ENMAX PDF.
    Returns: {'file', 'house_number', 'bill_amount', 'bill_date'}
    - bill_date is ISO YYYY-MM-DD from CurrentBillDate
    - house_number is matched from SERVICE ADDRESS or address lines using config numbers
    """
    path = os.path.join(folder, pdf_file) if folder else pdf_file
    text = _extract_first_page_text(path)

    # Foolproof: Look for house numbers near street indicators (AVE, ST, etc.)
    houses = get_house_numbers()
    house_number = None
    for house in sorted(houses, key=len, reverse=True):  # Try longer numbers first
        # Look for house number within 3-4 characters of street indicators
        street_pattern = rf'({re.escape(house)})\w{{0,4}}(?:AVE|ST|STREET|AVENUE|RD|ROAD|BLVD|BOULEVARD|DR|DRIVE|WAY|LANE|LN|CT|COURT|PL|PLACE)'
        street_match = re.search(street_pattern, text, re.IGNORECASE)
        if street_match:
            house_number = house
            break

    amount_match = re.search(
        r'(PreAuthorizedAmount.*?\$|TotalCurrentCharges.*?\$)\s*([\d]+\.\d{2})',
        text, re.IGNORECASE
    )
    bill_amount = amount_match.group(2) if amount_match else None

    date_match = re.search(
        r'CurrentBillDate:\s*(\d{4})(January|February|March|April|May|June|July|August|September|October|November|December)(\d{1,2})',
        text, re.IGNORECASE
    )
    bill_date = None
    if date_match:
        y, mon, d = date_match.groups()
        bill_date = f"{y}-{MONTHS[mon.lower()]}-{int(d):02d}"

    return {
        "file": os.path.basename(path),
        "house_number": house_number,
        "bill_amount": bill_amount,
        "bill_date": bill_date,
        "vendor": "ENMAX",
    }

# -------------------- ATCO extraction --------------------

MONTHS_ABBR = {
    "JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
    "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"
}

def make_house_line_regex(houses):
    """Create regex to match house numbers at start of address lines."""
    alts = "|".join(sorted((re.escape(h) for h in houses), key=len, reverse=True))
    return re.compile(rf'(?mi)^\s*({alts})\b')

def _get_previous_month(month: int, year: int) -> tuple:
    """
    Get the previous month and year, handling year rollover.
    
    Args:
        month: Current month (1-12)
        year: Current year
    
    Returns:
        Tuple of (previous_month, adjusted_year)
    """
    if month == 1:
        return (12, year - 1)
    return (month - 1, year)

def extract_atco_from_pdf(pdf_file: str, folder: str = None) -> dict:
    """
    Extract bill data from ATCO PDF.
    Returns: {'file', 'house_number', 'bill_amount', 'bill_date'}
    - bill_date is ISO YYYY-MM-DD derived from 'Total Amount Due By' month (minus 1 month)
    - house_number is matched at start of address line using config numbers
    """
    path = os.path.join(folder, pdf_file) if folder else pdf_file
    text = _extract_first_page_text(path)

    house_re = make_house_line_regex(get_house_numbers())
    house_match = house_re.search(text)
    house_number = house_match.group(1) if house_match else None

    amt_match = re.search(
        r'(?:TOTAL\s+AMOUNT\s+DUE|Amount\s+Due)\s*:?\s*\$?\s*([\d,]+\.\d{2})',
        text, re.IGNORECASE
    )
    bill_amount = amt_match.group(1).replace(",", "") if amt_match else None

    # Extract month from "Total Amount Due By: MON DD, YYYY" and use previous month
    bill_date = None
    due_by_match = re.search(
        r'Total\s+Amount\s+Due\s+By:\s*([A-Z]{3})\s+\d{1,2},\s*(\d{4})',
        text, re.IGNORECASE
    )
    if due_by_match:
        mon_abbr, y = due_by_match.groups()
        mm = MONTHS_ABBR.get(mon_abbr.upper())
        if mm:
            due_month = int(mm)
            due_year = int(y)
            # Use previous month as the bill month
            prev_month, prev_year = _get_previous_month(due_month, due_year)
            bill_date = f"{prev_year}-{prev_month:02d}-01"

    return {
        "file": os.path.basename(path),
        "house_number": house_number,
        "bill_amount": bill_amount,
        "bill_date": bill_date,
        "vendor": "ATCO",
    }

