"""
PDF Table OCR Extractor

This script processes PDF files from an input directory, performs OCR on each page,
and extracts table data into CSV and Excel files.

Usage:
    python process_pdf_tables.py [--rotate ROTATE] [--crop CROP]

Example:
    python process_pdf_tables.py --rotate 0
"""
import argparse, io, os, re, sys
from pathlib import Path

import fitz                       # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd
from tqdm import tqdm

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# --------------------------- Parser & CLI --------------------------- #
def get_args() -> argparse.Namespace:
    """Parse command line arguments."""
    ap = argparse.ArgumentParser(description="Extract tables from PDFs using OCR")
    ap.add_argument("--rotate", type=int, choices=[0, 90, 180, 270], default=0,
                    help="Rotation angle in degrees")
    ap.add_argument("--crop", type=float, default=0.05,
                    help="Right margin crop ratio (0-1)")
    return ap.parse_args()


# --------------------------- Regex Patterns --------------------------- #
POSTAL_RE = re.compile(r"^\d{4,6}$")  # German postal code pattern
EMAIL_RE  = re.compile(r"[^@\s]+@[^@\s]+\.[A-Za-z]{2,}")  # Email pattern
PHONE_RE  = re.compile(r"\+\d[\d\s\-\(\)]+$")  # Phone number pattern


# --------------------------- Core Functions --------------------------- #
def rotate_page(page, rotate_deg: int, dpi: int = 200) -> Image.Image:
    """
    Rotate a PDF page and convert it to an image.
    
    Args:
        page: PDF page object
        rotate_deg: Rotation angle in degrees
        dpi: Image resolution
        
    Returns:
        PIL Image object
    """
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom).prerotate(rotate_deg)
    pix = page.get_pixmap(matrix=mat)
    return Image.open(io.BytesIO(pix.tobytes()))


def parse_lines(text: str, page_idx: int) -> list[dict]:
    """
    Parse OCR text into structured data.
    
    Args:
        text: OCR text from a page
        page_idx: Page number (0-based)
        
    Returns:
        List of dictionaries containing parsed data
    """
    rows = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.lower().startswith("plz"):
            continue

        tokens = line.split()
        email_idx = next((i for i, t in enumerate(tokens) if "@" in t), None)
        if email_idx is None:
            continue

        postal_idx = next((i for i, t in enumerate(tokens[:email_idx]) if POSTAL_RE.fullmatch(t)), None)
        if postal_idx is None or postal_idx < 3:
            continue

        company   = " ".join(tokens[: postal_idx - 2])
        city      = tokens[postal_idx - 2]
        street_no = tokens[postal_idx - 1]
        postal    = tokens[postal_idx]

        after = tokens[postal_idx + 1 : email_idx]
        if len(after) < 2:
            continue

        name  = " ".join(after[:2])
        title = " ".join(after[2:]) if len(after) > 2 else ""

        email = tokens[email_idx]
        phone_match = PHONE_RE.search(" ".join(tokens[email_idx + 1 :]))
        phone = phone_match.group(0) if phone_match else ""

        rows.append(
            dict(
                Page=page_idx + 1,
                Company=company,
                City=city,
                StreetNo=street_no,
                PostalCode=postal,
                Name=name,
                Title=title,
                Email=email,
                Phone=phone,
            )
        )
    return rows


def process_pdf(pdf_path: Path, args: argparse.Namespace) -> None:
    """
    Process a single PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        args: Command line arguments
    """
    print(f"\nProcessing: {pdf_path.name}")
    output_dir = Path("output")
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf = fitz.open(pdf_path)
    all_rows: list[dict] = []

    for page_idx in tqdm(range(pdf.page_count), desc="OCR Processing"):
        img = rotate_page(pdf.load_page(page_idx), args.rotate)
        # Crop right margin
        w, h = img.size
        img = img.crop((0, 0, int(w * (1 - args.crop)), h))

        ocr_text = pytesseract.image_to_string(img, lang="deu+eng", config="--psm 6")
        all_rows.extend(parse_lines(ocr_text, page_idx))

    pdf.close()

    # Create DataFrame and save files
    df = pd.DataFrame(all_rows)
    csv_file   = output_dir / (pdf_path.stem + "_tables.csv")
    excel_file = output_dir / (pdf_path.stem + "_tables.xlsx")

    df.to_csv(csv_file, index=False)
    df.drop(columns=["Page"]).to_excel(excel_file, index=False)

    print(f"Complete – {len(df)} rows written to {csv_file} and {excel_file}")


def main() -> None:
    """Main entry point."""
    args = get_args()
    input_dir = Path("input")
    
    if not input_dir.exists():
        print("Error: 'input' directory not found!")
        sys.exit(1)
        
    pdf_files = list(input_dir.glob("*.pdf"))
    if not pdf_files:
        print("No PDF files found in 'input' directory!")
        sys.exit(1)
        
    print(f"Found {len(pdf_files)} PDF files")
    for pdf_file in pdf_files:
        process_pdf(pdf_file, args)


if __name__ == "__main__":
    main()
