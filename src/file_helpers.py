import os
import re
import calendar
import logging
import sys
from pathlib import Path
from typing import Dict, List, Optional

# Add parent directory to path to import config
sys.path.append(str(Path(__file__).parent.parent))
# Add current directory to path for local imports
sys.path.append(str(Path(__file__).parent))
from config import (
    get_raw_bills_folder,
    get_processed_bills_folder,
    get_images_folder,
    get_image_bottom_crop_px,
)
from pdf_utils import convert_pdf_to_image

log = logging.getLogger("bill-pipeline")

# -------------------- File & Directory Utilities --------------------

def setup_directories() -> None:
    """Create all necessary directories for the bill processing pipeline."""
    Path(get_raw_bills_folder()).mkdir(parents=True, exist_ok=True)
    Path(get_processed_bills_folder()).mkdir(parents=True, exist_ok=True)
    Path(get_images_folder()).mkdir(parents=True, exist_ok=True)

def get_pdf_files() -> List[str]:
    """Get sorted list of PDF files from raw bills folder (top-level only)."""
    raw = Path(get_raw_bills_folder())
    if not raw.exists():
        return []
    return sorted([p.name for p in raw.iterdir() if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith(".")])

def safe_filename(name: str) -> str:
    """Convert a string to a safe filename by replacing invalid characters."""
    return re.sub(r'[^A-Za-z0-9 _.\-]', '_', name).strip()

def ensure_unique_path(folder: Path, base: str, ext: str) -> Path:
    """Generate a unique file path by adding numbers if file already exists."""
    candidate = folder / f"{base}{ext}"
    i = 1
    while candidate.exists():
        candidate = folder / f"{base} ({i}){ext}"
        i += 1
    return candidate

# -------------------- Date & Filename Formatting --------------------

def iso_to_month_day_year(iso: str) -> str:
    """Convert ISO date (YYYY-MM-DD) to 'Month DD YYYY' format."""
    y, m, d = map(int, iso.split("-"))
    return f"{calendar.month_name[m]} {d} {y}"

def iso_to_year_month(iso: str) -> str:
    """Convert ISO date (YYYY-MM-DD) to 'YYYY-MM' format."""
    y, m, _ = map(int, iso.split("-"))
    return f"{y:04d}-{m:02d}"

def build_target_filename(house: str, iso_date: str, vendor: str, ext: str) -> str:
    """Build standardized filename: '{house} {YYYY-MM} {vendor}.ext'"""
    date_formatted = iso_to_year_month(iso_date)
    return safe_filename(f"{house} {date_formatted} {vendor}") + ext

# -------------------- PDF & Image Processing --------------------

def create_pdf_image(filename: str, house, iso_date: str, vendor: str, source_folder: str) -> None:
    """
    Convert first page of PDF to image with bottom crop (pixels).
    Avoid double-rendering; cropping is handled inside convert_pdf_to_image.
    
    Args:
        filename: Name of the PDF file
        house: House number
        iso_date: ISO date string (YYYY-MM-DD)
        vendor: Vendor name
        source_folder: Source folder path containing the PDF
    """
    src_path = Path(source_folder) / filename
    image_name = safe_filename(f"{house}_{iso_date}_{vendor}.png")
    image_path = Path(get_images_folder()) / image_name
    convert_pdf_to_image(
        str(src_path),
        str(image_path),
        page=1,
        dpi=300,
        bottom_crop_px=get_image_bottom_crop_px(),
    )

def move_processed_file(filename: str, house: str, iso_date: str, vendor: str) -> str:
    """
    Move a processed PDF file from raw to processed folder with standardized naming.
    Overwrites existing files with the same name.
    Returns the final filename used.
    """
    src_path = Path(get_raw_bills_folder()) / filename
    _, ext = os.path.splitext(filename)
    target_filename = build_target_filename(str(house), iso_date, vendor, ext)
    dst_path = Path(get_processed_bills_folder()) / target_filename
    
    try:
        os.replace(str(src_path), str(dst_path))  # atomic if same volume, overwrites existing
        log.info("Moved: %s -> %s", filename, dst_path.name)
        return dst_path.name
    except Exception:
        log.exception("Failed to move file: %s", filename)
        return filename  # fallback to original name
