import logging
from typing import Dict, List, Optional

from config import (
    get_excel_path,
    get_excel_data_sheet,
    get_raw_bills_folder,
    get_processed_bills_folder,
    get_rename_files,
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

def validate_data(data: Dict, filename: str) -> None:
    """
    Validate extracted bill data for required fields.
    
    Raises ValueError if validation fails.
    Logs warning for missing optional fields.
    """
    house = data.get("house_number")
    iso_date = data.get("bill_date")
    amount = data.get("bill_amount")
    
    # Validate required fields
    if not house or not iso_date:
        error_msg = f"Missing required fields - house: {house}, date: {iso_date}"
        log.error("Validation failed for %s: %s", filename, error_msg)
        raise ValueError(error_msg)
    
    if amount is None:
        log.warning("Amount missing for %s; continuing with empty amount", filename)


def rename_file(filename: str, house: str, iso_date: str, vendor: str) -> str:
    """
    Determine target filename based on RENAME_FILES config.
    
    If config is True, returns standardized filename.
    If config is False, returns original filename.
    Does not actually move or rename the file - just determines the name.
    """
    if get_rename_files():
        # Use standardized naming from file_helpers
        import os
        from src.file_helpers import build_target_filename
        _, ext = os.path.splitext(filename)
        target_filename = build_target_filename(str(house), iso_date, vendor, ext)
        log.info("Rename enabled: %s -> %s", filename, target_filename)
        return target_filename
    else:
        log.info("Rename disabled: keeping original filename %s", filename)
        return filename


def move_file(filename: str, target_filename: str) -> None:
    """
    Move file from raw bills folder to processed bills folder.
    Always moves the file regardless of config.
    
    Args:
        filename: Original filename in raw folder
        target_filename: Target filename in processed folder
    """
    import os
    from pathlib import Path
    
    src_path = Path(get_raw_bills_folder()) / filename
    dst_path = Path(get_processed_bills_folder()) / target_filename
    
    try:
        os.replace(str(src_path), str(dst_path))
        log.info("File moved: %s -> %s", src_path.name, dst_path.name)
    except Exception as e:
        log.error("Failed to move file %s: %s", filename, str(e))
        raise


def process_single_file(filename: str, source_folder: str = None, move_file_after: bool = True) -> Optional[Dict]:
    """
    Process a single PDF bill file through the complete pipeline.
    
    Steps:
    1. Extract data from PDF
    2. Validate extracted data
    3. Create image from PDF (non-fatal)
    4. Rename file (conditional on config, only if move_file_after=True)
    5. Move file (only if move_file_after=True)
    
    Args:
        filename: Name of the PDF file
        source_folder: Source folder containing the PDF. If None, uses get_raw_bills_folder()
        move_file_after: If True, rename and move file after processing. If False, leave file in place.
    
    Returns dict with processed data for Excel, or None if processing failed.
    """
    if source_folder is None:
        source_folder = get_raw_bills_folder()
    
    # Step 1: Extract data (fail fast on error)
    try:
        data = route_and_extract(filename, source_folder)
    except Exception as e:
        log.error("Extraction failed for %s: %s", filename, str(e))
        return None
    
    # Step 2: Validate data (fail fast on error)
    try:
        validate_data(data, filename)
    except ValueError:
        log.error("Validation failed for %s", filename)
        return None
    
    house = data["house_number"]
    iso_date = data["bill_date"]
    amount = data["bill_amount"]
    vendor = data["vendor"]
    
    # Step 3: Create image (non-fatal, continue on failure)
    try:
        create_pdf_image(filename, house, iso_date, vendor, source_folder)
        log.info("Image created successfully for %s", filename)
    except Exception as e:
        log.error("Image creation failed for %s: %s (continuing)", filename, str(e))
    
    # Step 4: Determine target filename (rename if config enabled)
    target_filename = rename_file(filename, str(house), iso_date, vendor)
    
    # Step 5: Move file (only if requested)
    if move_file_after:
        try:
            move_file(filename, target_filename)
        except Exception:
            log.error("Failed to move file: %s", filename)
            return None
    
    return {
        "file": target_filename,
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

