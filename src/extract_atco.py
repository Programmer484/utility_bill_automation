import os
import re
import sys
from pathlib import Path

# Add parent directory to path to import config
sys.path.append(str(Path(__file__).parent.parent))
from config import get_house_numbers
from pypdf import PdfReader

MONTHS_ABBR = {
    "JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
    "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"
}

def make_house_line_regex(houses):
    # Match a configured house number at the start of an address line
    alts = "|".join(sorted((re.escape(h) for h in houses), key=len, reverse=True))
    return re.compile(rf'(?mi)^\s*({alts})\b')

def _extract_first_page_text(path: str) -> str:
    reader = PdfReader(path)
    return reader.pages[0].extract_text() or ""

def extract_atco_from_pdf(
    pdf_file: str,
    folder: str = None,
) -> dict:
    """
    Returns: {'file', 'house_number', 'bill_amount', 'bill_date'}
    - bill_date is ISO YYYY-MM-DD from 'Statement Date: AUG 20, 2025'
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

    date_match = re.search(
        r'Statement\s*Date:\s*([A-Z]{3})\s+(\d{1,2}),\s*(\d{4})',
        text, re.IGNORECASE
    )
    bill_date = None
    if date_match:
        mon_abbr, d, y = date_match.groups()
        mm = MONTHS_ABBR.get(mon_abbr.upper())
        if mm:
            bill_date = f"{y}-{mm}-{int(d):02d}"

    return {
        "file": os.path.basename(path),
        "house_number": house_number,
        "bill_amount": bill_amount,
        "bill_date": bill_date,
        "vendor": "ATCO",
    }
