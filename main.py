import logging
import re
from typing import Dict, List, Optional

from config import (
    get_excel_path,
    get_excel_data_sheet,
    get_raw_bills_folder,
    get_atco_indicator,
    get_move_processed_files,
)
from src.excel import append_rows_to_excel
from src.extract_enmax import extract_enmax_from_pdf
from src.extract_atco import extract_atco_from_pdf
from src.file_helpers import (
    setup_directories,
    get_pdf_files,
    create_pdf_image,
    move_processed_file,
)
from src.email_drafts import generate_email_drafts

# -------------------- logging --------------------
logging.basicConfig(    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("bill-pipeline")

# -------------------- helpers --------------------

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
    """Extract bill data from PDF and determine vendor by filename as fallback."""
    atco_indicator = get_atco_indicator()
    raw_bills_folder = get_raw_bills_folder()
    
    vendor = "ATCO" if atco_indicator in filename.lower() else "ENMAX"
    extractor = extract_atco_from_pdf if vendor == "ATCO" else extract_enmax_from_pdf
    try:
        data = extractor(filename, folder=raw_bills_folder)
    except Exception as e:
        log.exception("Extractor failed: %s", filename)
        raise

    # prefer extractor's vendor if present; else use filename heuristic
    extracted_vendor = (data.get("vendor") or "").strip().upper()
    data["vendor"] = extracted_vendor or vendor
    return normalize_row(data)

def process_single_file(filename: str) -> Optional[Dict]:
    """Extract, create image, move PDF; return row for Excel or None."""
    try:
        data = route_and_extract(filename)
    except Exception:
        log.error("Skipping due to extraction error: %s", filename)
        return None

    house = data.get("house_number")
    iso_date = data.get("bill_date")
    amount = data.get("bill_amount")
    vendor = data.get("vendor")

    if not (house and iso_date):
        log.warning("Skip (missing house/date): %s -> house:%s date:%s", filename, house, iso_date)
        return None
    if amount is None:
        log.warning("Amount missing/invalid; continuing: %s", filename)

    # Create image (non-fatal on failure)
    try:
        create_pdf_image(filename, house, iso_date, vendor)
    except Exception:
        log.exception("Image conversion failed (continuing): %s", filename)

    # Conditionally move PDF file
    if get_move_processed_files():
        final_filename = move_processed_file(filename, str(house), iso_date, vendor)
        log.info("File moved to processed folder")
    else:
        final_filename = filename  # Keep original filename, file stays in Downloads
        log.info("File left in Downloads folder (MOVE_PROCESSED_FILES=False)")

    return {
        "file": final_filename,
        "house_number": house,
        "bill_amount": amount if amount is not None else "",
        "bill_date": iso_date,
        "vendor": vendor,
    }


def output_results(rows: List[Dict]) -> None:
    """Save results to Excel and print to console."""
    append_rows_to_excel(get_excel_path(), rows, get_excel_data_sheet())
    for row in rows:
        print(f"{row['file']}\t{row['house_number']}\t{row['bill_amount']}\t{row['bill_date']}\t{row['vendor']}")


def main():
    setup_directories()

    pdf_files = get_pdf_files()
    if not pdf_files:
        log.info("No PDFs found in %s", get_raw_bills_folder())
        # Still generate emails even if no new PDFs to process
        log.info("Generating email drafts from existing data...")
        generate_email_drafts()
        return

    rows: List[Dict] = []

    for filename in pdf_files:
        result = process_single_file(filename)
        if not result:
            continue
        rows.append(result)

    if rows:
        output_results(rows)
        log.info("Processing complete. Generating email drafts...")
        generate_email_drafts()
    else:
        log.info("No new rows to append.")
        log.info("Generating email drafts from existing data...")
        generate_email_drafts()

if __name__ == "__main__":
    main()
