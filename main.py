import logging
from typing import Dict, List, Optional

from config import (
    get_excel_path,
    get_excel_data_sheet,
    get_raw_bills_folder,
    get_move_processed_files,
)
from src.excel import append_rows_to_excel
from src.data_processing import route_and_extract
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

# -------------------- processing functions --------------------

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
        log.info("No bills to process, no emails to generate.")
        return

    rows: List[Dict] = []

    for filename in pdf_files:
        result = process_single_file(filename)
        if not result:
            continue
        rows.append(result)

    if rows:
        # Generate emails from fresh data BEFORE saving to Excel
        log.info("Processing complete. Generating email drafts from fresh data...")
        generate_email_drafts(rows)
        
        # Save to Excel for record-keeping
        log.info("Saving processed bills to Excel...")
        output_results(rows)
    else:
        log.info("No valid bills processed, no emails to generate.")

if __name__ == "__main__":
    main()

