"""
Data processing functionality - normalization, validation, and extraction routing.
"""

import re
import logging
from typing import Dict

from config import get_raw_bills_folder
from src.bill_extractors import extract_enmax_from_pdf, extract_atco_from_pdf
from src.vendor_detection import detect_vendor_from_pdf

log = logging.getLogger("bill-pipeline")


def normalize_row(row: Dict) -> Dict:
    """Normalize fields: vendor uppercase/trim, numeric amount, ISO date."""
    out = dict(row)
    # vendor
    v = (out.get("vendor") or "").strip().upper()
    if v in {"ENMAX", "ATCO"}:
        out["vendor"] = v
    else:
        # if unknown, keep whatever we got but uppercased
        out["vendor"] = v or "UNKNOWN"

    # amount
    amt = out.get("bill_amount")
    try:
        out["bill_amount"] = float(amt) if amt is not None and amt != "" else None
    except Exception:
        out["bill_amount"] = None

    # bill_date (must be YYYY-MM-DD)
    date_str = (out.get("bill_date") or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_str):
        out["bill_date"] = date_str
    else:
        # attempt to coerce simple patterns like "YYYY/M/D" etc.
        m = re.fullmatch(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", date_str)
        if m:
            y, mo, d = map(int, m.groups())
            out["bill_date"] = f"{y:04d}-{mo:02d}-{d:02d}"
        else:
            out["bill_date"] = None

    # house_number as int if possible (kept as string otherwise)
    house = out.get("house_number")
    try:
        out["house_number"] = int(house)
    except Exception:
        out["house_number"] = house  # leave as-is

    return out


def route_and_extract(filename: str) -> Dict:
    """Extract bill data from PDF and determine vendor by content analysis."""
    raw_bills_folder = get_raw_bills_folder()
    
    # Detect vendor from PDF content
    vendor = detect_vendor_from_pdf(filename, raw_bills_folder)
    extractor = extract_atco_from_pdf if vendor == "ATCO" else extract_enmax_from_pdf
    
    try:
        data = extractor(filename, folder=raw_bills_folder)
    except Exception as e:
        log.exception("Extractor failed: %s", filename)
        raise

    # Use the detected vendor (extractor should return the same vendor)
    data["vendor"] = vendor
    return normalize_row(data)
