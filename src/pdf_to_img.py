from pdf2image import convert_from_path
from PIL import Image
from typing import Optional, Tuple

def convert_pdf_to_image(
    pdf_path: str,
    image_path: str,
    page: int = 1,
    dpi: int = 300,
    crop_box: Optional[Tuple[int, int, int, int]] = None,
    bottom_crop_px: Optional[int] = None,
) -> None:
    """
    Render a single PDF page to an image and optionally crop it.

    - If crop_box is given, uses (left, top, right, bottom) pixels.
    - Else if bottom_crop_px is given, crops that many pixels off the bottom.
    - Saves as PNG inferred from image_path extension.

    This function renders the page only once (avoids double rendering).
    """
    pages = convert_from_path(pdf_path, dpi=dpi, first_page=page, last_page=page)
    if not pages:
        raise RuntimeError(f"No pages rendered from {pdf_path}")
    img = pages[0]

    if crop_box is not None:
        img = img.crop(crop_box)
    elif bottom_crop_px:
        w, h = img.size
        bottom = max(h - int(bottom_crop_px), 0)
        img = img.crop((0, 0, w, bottom))

    img.save(image_path)
