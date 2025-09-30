"""
Vendor detection functionality - analyzes PDF content to identify bill vendors.
"""

import os
import logging
from pypdf import PdfReader

log = logging.getLogger("bill-pipeline")


def detect_vendor_from_pdf(filename: str, folder: str) -> str:
    """Analyze PDF content to determine vendor (ENMAX or ATCO)."""
    path = os.path.join(folder, filename)
    try:
        reader = PdfReader(path)
        text = reader.pages[0].extract_text() or ""
        if not text.strip():
            raise ValueError(f"No text could be extracted from PDF: {filename}")
        
        text_upper = text.upper()
        
        # Simple, stable heuristics - one per vendor
        if "ENMAX.COM" in text_upper:
            return "ENMAX"
        elif "STATEMENT DATE:" in text_upper:
            return "ATCO"
        else:
            raise ValueError(f"Could not determine vendor from PDF content: {filename}. "
                           f"Neither ENMAX.COM nor 'STATEMENT DATE:' found in text.")
                
    except Exception as e:
        if "Could not determine vendor" in str(e):
            raise  # Re-raise vendor detection errors
        else:
            raise ValueError(f"Error reading PDF for vendor detection: {filename} - {e}")
