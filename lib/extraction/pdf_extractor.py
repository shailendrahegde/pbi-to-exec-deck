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


def _apply_exif_correction(img: Image.Image) -> Image.Image:
    """Apply EXIF orientation metadata only (safe no-op if no EXIF present)."""
    from PIL import ImageOps
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    return img


def _compute_n_strips(img: Image.Image) -> int:
    """
    Compute how many 16:9 strips a portrait page should be split into.

    Some Power BI / Viva Insights PDF exports pack a vertically-scrollable
    dashboard into a standard portrait page (e.g. 8.5"×11").  Each 'screen'
    of the dashboard occupies one 16:9 band of the page height.  We split
    the page into that many bands so each band becomes its own slide.

    Returns 1 for landscape images (no split needed).
    """
    width, height = img.size
    if width >= height:
        return 1  # Already landscape — no split needed
    target_strip_height = width * 9 / 16
    return max(1, round(height / target_strip_height))


def _detect_background_brightness(img: Image.Image, sample_size: int = 10) -> int:
    """
    Detect the report background brightness by sampling the four corner patches.

    Report backgrounds vary: white, light grey, mid grey, dark grey, etc.
    Corner patches are reliably background-only areas (charts and tables never
    start in the extreme corners), so their median brightness is a robust proxy
    for the page background colour regardless of the specific shade used.
    """
    width, height = img.size
    sz = min(sample_size, width // 4, height // 4)
    gray = img.convert('L')
    corners = [
        (0,           0),
        (width - sz,  0),
        (0,           height - sz),
        (width - sz,  height - sz),
    ]
    samples = []
    for x0, y0 in corners:
        patch = gray.crop((x0, y0, x0 + sz, y0 + sz))
        samples.extend(list(patch.getdata()))
    samples.sort()
    return samples[len(samples) // 2]  # median corner brightness


def _trim_background_rows(img: Image.Image, bg: int, bg_tol: int = 15,
                          margin: int = 3) -> Image.Image:
    """
    Remove leading and trailing background-only rows from an image.

    PDF pages often have large blank areas above or below the actual content
    (e.g. a glossary page where all content is in the lower 35% of the page).
    Trimming these rows BEFORE computing n_strips prevents the page from being
    split into a mostly-blank strip and a content strip — the content height
    after trimming may be short enough to need only 1 strip.

    A margin of background rows is preserved so the visual padding around the
    outermost content element is not clipped.
    """
    width, height = img.size
    narrow = img.convert('L').resize((1, height), Image.LANCZOS)
    row_brightness = [narrow.getpixel((0, y)) for y in range(height)]

    top = 0
    for y in range(height):
        if abs(row_brightness[y] - bg) > bg_tol:
            top = max(0, y - margin)
            break

    bottom = height
    for y in range(height - 1, -1, -1):
        if abs(row_brightness[y] - bg) > bg_tol:
            bottom = min(height, y + margin + 1)
            break

    if top >= bottom:
        return img  # Entirely background — return as-is

    return img.crop((0, top, width, bottom))


def _find_split_point(img: Image.Image, target_y: int,
                      search_window_frac: float = 0.30,
                      cap_rows: int = 20,
                      bg_tol: int = 15,
                      bg_frac: float = 0.85) -> int:
    """
    Find the nearest clean horizontal split point to target_y.

    A split position is "clean" when both sides have a background cap: at least
    bg_frac of cap_rows immediately above AND below the cut match the detected
    page background colour (within bg_tol brightness units). This works for any
    background shade — white, light grey, mid grey, dark grey, etc.

    Algorithm:
      1. Detect background brightness from corner patches.
      2. Adjust target_y to the content midpoint (ignoring leading/trailing
         background rows) so pages with large blank areas at top/bottom are
         split through their content, not through dead space.
      3. If the adjusted target is already a clean split, return it immediately.
      4. Otherwise expand ±1 row at a time in BOTH directions simultaneously
         (inward toward page centre AND outward toward page edges) and return
         the first clean position found — whichever direction is closer.
      5. Fall back to the content midpoint if no clean split is found within
         the search window.
    """
    width, height = img.size

    # --- Step 1: Detect background colour ---
    bg = _detect_background_brightness(img)

    def is_bg(v: int) -> bool:
        return abs(v - bg) <= bg_tol

    # Compress width → 1px for efficient per-row brightness (no numpy needed)
    narrow = img.convert('L').resize((1, height), Image.LANCZOS)
    row_brightness = [narrow.getpixel((0, y)) for y in range(height)]

    # --- Step 2: Adjust target to content midpoint ---
    # Skip leading/trailing background rows so pages with large blank areas
    # (e.g. glossary page with background at top) are split through content.
    content_rows = [y for y, b in enumerate(row_brightness) if not is_bg(b)]
    if content_rows and (content_rows[-1] - content_rows[0]) >= 50:
        target_y = (content_rows[0] + content_rows[-1]) // 2

    # --- Step 3: Define search bounds ---
    window_px = max(20, int(height * search_window_frac))
    search_lo = max(cap_rows, target_y - window_px)
    search_hi = min(height - cap_rows, target_y + window_px)

    # --- Step 4: Clean-split predicate ---
    def cap_ok(y: int) -> bool:
        """True when both sides of a split at y have a background cap."""
        above = [row_brightness[y - i - 1] for i in range(cap_rows) if y - i - 1 >= 0]
        below = [row_brightness[y + i]     for i in range(cap_rows) if y + i < height]
        if not above or not below:
            return False
        return (sum(is_bg(v) for v in above) / len(above) >= bg_frac and
                sum(is_bg(v) for v in below) / len(below) >= bg_frac)

    # --- Step 5: Bidirectional search ---
    # Check target first; then expand ±1 row at a time in both directions,
    # returning the closest clean split found in either direction.
    if cap_ok(target_y):
        return target_y

    for delta in range(1, window_px + 1):
        lo, hi = target_y - delta, target_y + delta
        lo_ok = (search_lo <= lo) and cap_ok(lo)
        hi_ok = (hi <= search_hi) and cap_ok(hi)

        if lo_ok and hi_ok:
            # Both equidistant — prefer whichever is closer to page centre
            return lo if abs(lo - height // 2) < abs(hi - height // 2) else hi
        if lo_ok:
            return lo
        if hi_ok:
            return hi

    return target_y  # Fallback: content midpoint


def _split_image_into_strips(img: Image.Image, n_strips: int) -> list:
    """Split image into n strips using content-aware, bidirectional split detection.

    For each ideal split position, finds the nearest clean split (whitespace cap
    on both sides) searching both inward and outward, preventing cuts through
    chart cards, table rows, or definition boxes.
    """
    if n_strips <= 1:
        return [img]

    width, height = img.size

    # Compute ideal split positions, then find nearest clean split for each
    split_ys = []
    for i in range(1, n_strips):
        ideal_y = round(i * height / n_strips)
        split_ys.append(_find_split_point(img, ideal_y))

    # Crop strips using the (possibly adjusted) split points
    boundaries = [0] + split_ys + [height]
    strips = []
    for i in range(len(boundaries) - 1):
        top = boundaries[i]
        bottom = boundaries[i + 1]
        strips.append(img.crop((0, top, width, bottom)))

    return strips


def extract_pdf_page_as_image(pdf_document: fitz.Document, page_idx: int, output_path: str) -> Optional[Image.Image]:
    """
    Extract PDF page as PNG image (mirrors extract_slide_as_image for PPTX).

    Args:
        pdf_document: PyMuPDF document object
        page_idx: Page index (0-based)
        output_path: Path to save PNG image

    Returns:
        PIL Image if successful, None on failure
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

        # Apply EXIF correction only — geometric rotation is NOT applied here.
        # Portrait pages with scrollable dashboard content are split into strips
        # by the caller; rotating them would make the text unreadable.
        img = _apply_exif_correction(img)

        # Save as PNG
        img.save(output_path)

        return img  # Return image so caller can split if needed

    except Exception as e:
        print(f"  WARNING: Failed to extract page {page_idx + 1}: {e}")
        return None


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

    slide_counter = 0  # Sequential slide number across all pages and strips

    for page_idx in range(len(pdf_doc)):
        # Include all pages for PDF (unlike PPTX which skips cover page)
        page = pdf_doc[page_idx]
        title = extract_pdf_page_title(page)
        raw_path = f"temp/page_{page_idx + 1}_raw.png"

        # Extract page as image (EXIF-corrected, no rotation)
        img = extract_pdf_page_as_image(pdf_doc, page_idx, raw_path)

        if img is None:
            continue

        # Trim blank background rows before computing strip count.
        # Pages with sparse content (e.g. a glossary in the bottom 35% of the
        # page) would otherwise split into a mostly-blank strip + content strip.
        # After trimming, the content height may fit in a single strip.
        bg = _detect_background_brightness(img)
        img = _trim_background_rows(img, bg)

        # Determine how many 16:9 strips this page should be split into
        n_strips = _compute_n_strips(img)
        strips = _split_image_into_strips(img, n_strips)

        strip_label = f" ({n_strips} strips)" if n_strips > 1 else ""
        print(f"  OK Page {page_idx + 1}: {title[:40]}...{strip_label}")

        for strip_idx, strip_img in enumerate(strips):
            slide_counter += 1

            # Build strip title and image path
            if n_strips > 1:
                strip_title = f"{title} (Part {strip_idx + 1} of {n_strips})"
                image_path = f"temp/page_{page_idx + 1}_strip{strip_idx + 1}.png"
            else:
                strip_title = title
                image_path = f"temp/page_{page_idx + 1}.png"

            strip_img.save(image_path)

            page_info = {
                'slide_number': slide_counter,
                'title': strip_title,
                'image_path': image_path,
                'slide_type': classify_slide_type(title)
            }
            pages_to_analyze.append(page_info)

        # Clean up raw file if strips were saved separately
        import os
        if n_strips > 1 and Path(raw_path).exists():
            os.remove(raw_path)

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
