import os
import re
import sys
from pathlib import Path

# Add parent directory to path to import config
sys.path.append(str(Path(__file__).parent.parent))
from config import get_house_numbers

from pypdf import PdfReader

MONTHS = {
    "january":"01","february":"02","march":"03","april":"04","may":"05","june":"06",
    "july":"07","august":"08","september":"09","october":"10","november":"11","december":"12"
}

def make_service_address_regex(houses):
    # Look for SERVICE ADDRESS ... : <HOUSE>
    alts = "|".join(sorted((re.escape(h) for h in houses), key=len, reverse=True))
    pattern = rf"SERVICE\s*ADDRESS[^:\n]{{0,80}}:\s*({alts})"
    return re.compile(pattern, re.IGNORECASE)

def _extract_first_page_text(path: str) -> str:
    reader = PdfReader(path)
    return reader.pages[0].extract_text() or ""

def extract_enmax_from_pdf(
    pdf_file: str,
    folder: str = None,
) -> dict:
    """
    Returns: {'file', 'house_number', 'bill_amount', 'bill_date'}
    - bill_date is ISO YYYY-MM-DD from CurrentBillDate
    - house_number is matched from SERVICE ADDRESS using config numbers
    """
    path = os.path.join(folder, pdf_file) if folder else pdf_file
    text = _extract_first_page_text(path)

    svc_addr_re = make_service_address_regex(get_house_numbers())
    house_match = svc_addr_re.search(text)
    house_number = house_match.group(1) if house_match else None

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
