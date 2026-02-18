"""
PDF Extraction Module

Extracts pages from PDF files as PNG images for Claude analysis.
Mirrors the PPTX extraction workflow for consistent processing.
"""

import json
import re
import io
from pathlib import Path
from typing import Optional
import fitz  # PyMuPDF
from PIL import Image


def _correct_image_orientation(img: Image.Image) -> Image.Image:
    """
    Auto-correct image orientation for dashboard exports.

    Dashboard images (e.g., Power BI exports) are always landscape (16:9).
    Two common export artifacts are handled:
      1. EXIF rotation metadata embedded in the image file.
      2. 90-degree page rotation baked into the PDF stream — the rendered
         pixmap comes out portrait even though the content is landscape.

    In case (2) we rotate 90° clockwise (rotate(-90) in PIL), which is the
    standard correction for a landscape source that was embedded in portrait
    via a CCW rotation during PDF export.
    """
    from PIL import ImageOps

    # Apply EXIF orientation metadata first (safe no-op if no EXIF present)
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass

    width, height = img.size

    # If the image is taller than it is wide it is portrait — almost certainly
    # a 90-degree tilt artifact.  Rotate 90° CW (= rotate(-90)) to restore
    # landscape orientation.  Power BI embeds landscape content in portrait
    # PDF pages using a CCW rotation, so the inverse (CW) corrects it.
    if height > width:
        img = img.rotate(-90, expand=True)

    return img


def extract_pdf_page_as_image(pdf_document: fitz.Document, page_idx: int, output_path: str) -> bool:
    """
    Extract PDF page as PNG image (mirrors extract_slide_as_image for PPTX).

    Args:
        pdf_document: PyMuPDF document object
        page_idx: Page index (0-based)
        output_path: Path to save PNG image

    Returns:
        True if extraction successful, False otherwise
    """
    try:
        if page_idx >= len(pdf_document):
            return False

        page = pdf_document[page_idx]

        # Render page at 150 DPI (balance between quality and file size)
        # Power BI exports are typically high-resolution, 150 DPI is sufficient
        zoom = 150 / 72  # PDF default is 72 DPI
        mat = fitz.Matrix(zoom, zoom)

        # Render page to pixmap
        pix = page.get_pixmap(matrix=mat, alpha=False)

        # Convert to PIL Image
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))

        # Auto-correct orientation: handles EXIF metadata and 90° tilt
        # artifacts that can occur when exporting dashboards to PDF.
        img = _correct_image_orientation(img)

        # Save as PNG
        img.save(output_path)

        return True

    except Exception as e:
        print(f"  WARNING: Failed to extract page {page_idx + 1}: {e}")
        return False


def extract_pdf_page_title(page: fitz.Page) -> str:
    """
    Extract page title from PDF text layer (mirrors extract_slide_title for PPTX).

    Args:
        page: PyMuPDF page object

    Returns:
        Page title string (first non-empty line), or "Page N" if no text found
    """
    try:
        # Get text from page
        text = page.get_text("text")

        if not text or not text.strip():
            return f"Page {page.number + 1}"

        # Split into lines and find first non-empty line
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        if not lines:
            return f"Page {page.number + 1}"

        title = lines[0]

        # Remove emoji characters (same regex as PPTX version)
        title = re.sub(r'[\U00010000-\U0010ffff]', '', title).strip()

        # Fallback if empty after emoji removal
        if not title:
            return f"Page {page.number + 1}"

        return title

    except Exception:
        return f"Page {page.number + 1}"


def classify_slide_type(title: str) -> str:
    """
    Classify slide type for context (same logic as PPTX version).

    Args:
        title: Page/slide title

    Returns:
        Slide type classification
    """
    title_lower = title.lower()

    if 'trend' in title_lower or 'over time' in title_lower:
        return 'trends'
    elif 'leaderboard' in title_lower or 'top' in title_lower:
        return 'leaderboard'
    elif 'health' in title_lower or 'overview' in title_lower:
        return 'health_check'
    elif 'habit' in title_lower or 'frequency' in title_lower:
        return 'habit_formation'
    elif 'license' in title_lower or 'priority' in title_lower:
        return 'license_priority'
    else:
        return 'general'


def prepare_pdf_for_claude_analysis(source_path: str) -> str:
    """
    Prepare PDF pages for Claude analysis (mirrors prepare_for_claude_analysis for PPTX).

    Extracts each page as PNG image and creates analysis request JSON.
    NOTE: PDF exports typically have dashboard content on first page, so we include all pages.
    (Unlike PPTX exports from Power BI Online which have a cover page with metadata)

    Args:
        source_path: Path to source PDF file

    Returns:
        Path to analysis_request.json file
    """
    print("=" * 70)
    print("PREPARING PDF PAGES FOR CLAUDE ANALYSIS")
    print("=" * 70)

    # Open PDF document
    try:
        pdf_doc = fitz.open(source_path)
    except Exception as e:
        raise IOError(f"Failed to open PDF file '{source_path}': {e}")

    # Validate PDF
    if len(pdf_doc) == 0:
        raise ValueError(f"PDF file '{source_path}' is empty (0 pages)")

    # Create temp directory for images
    Path('temp').mkdir(exist_ok=True)

    pages_to_analyze = []

    print(f"\nExtracting all {len(pdf_doc)} pages from PDF...")
    print("  (PDF exports typically have dashboard content on page 1)")

    for page_idx in range(len(pdf_doc)):
        # Include all pages for PDF (unlike PPTX which skips cover page)
        page = pdf_doc[page_idx]
        title = extract_pdf_page_title(page)
        image_path = f"temp/page_{page_idx + 1}.png"

        # Extract page as image
        has_image = extract_pdf_page_as_image(pdf_doc, page_idx, image_path)

        if has_image:
            page_info = {
                'slide_number': page_idx + 1,  # Keep as 'slide_number' for compatibility
                'title': title,
                'image_path': image_path,
                'slide_type': classify_slide_type(title)
            }

            pages_to_analyze.append(page_info)
            print(f"  OK Page {page_idx + 1}: {title[:50]}...")

    # Close PDF document
    pdf_doc.close()

    # Save analysis request (identical structure to PPTX version)
    request_file = 'temp/analysis_request.json'
    with open(request_file, 'w', encoding='utf-8') as f:
        json.dump({
            'source_file': source_path,
            'total_slides': len(pages_to_analyze),
            'slides': pages_to_analyze
        }, f, indent=2)

    print(f"\nOK Prepared {len(pages_to_analyze)} pages for analysis")
    print(f"OK Analysis request saved to: {request_file}")

    return request_file
