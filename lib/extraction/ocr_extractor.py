"""
OCR-based extraction fallback using EasyOCR.

When text-layer extraction (markitdown / python-pptx) returns insufficient
data — common with Power BI PPTX exports that embed dashboards as full-page
PNG screenshots — this module runs EasyOCR on the slide images to extract
real numbers, labels, and KPI values.

Usage hierarchy (per user requirement):
    1. Try embedded text first (text_layer_extractor)
    2. If text is mostly empty / boilerplate → fall back to OCR here
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, List, Tuple

# Lazy-load easyocr to avoid import overhead when not needed
_reader = None


def _get_reader():
    """Lazy-initialise the EasyOCR reader (English, CPU-only)."""
    global _reader
    if _reader is None:
        try:
            import easyocr
            _reader = easyocr.Reader(["en"], gpu=False, verbose=False)
        except ImportError:
            raise ImportError(
                "easyocr is required for OCR fallback. "
                "Install it with:  pip install easyocr"
            )
    return _reader


# ---------------------------------------------------------------------------
# Quality check: decide whether text-layer output is useful
# ---------------------------------------------------------------------------

# Tokens that indicate markitdown returned only boilerplate, not real data
_BOILERPLATE_TOKENS = {
    "no alt text provided",
    "lineclusteredcolumncombochart",
    "clusteredcolumnchart",
    "card",
    "slicer",
    "table",
    "image",
    "shape",
}


def text_layer_is_sufficient(text_layer: str, text_metrics: list) -> bool:
    """Return True if text-layer extraction produced usable data.

    Heuristics:
    - If text_metrics contains at least 3 items with numeric values → sufficient
    - If text_layer is mostly boilerplate ("No alt text provided", chart type
      names) → insufficient even if long
    - If text_layer has >50 non-boilerplate tokens with real content → sufficient
    """
    # Check metrics
    good_metrics = sum(
        1 for m in text_metrics
        if m.get("numeric_value") is not None and m["numeric_value"] != 0
    )
    if good_metrics >= 3:
        return True

    # Check text content
    if not text_layer:
        return False

    lower = text_layer.lower()

    # If "no alt text provided" appears repeatedly, it's markitdown boilerplate
    if lower.count("no alt text provided") >= 3:
        return False

    # Filter out boilerplate tokens
    words = lower.split()
    useful = [w for w in words if w not in _BOILERPLATE_TOKENS and len(w) > 1]

    # Also discount common markdown / boilerplate fragments
    useful = [
        w for w in useful
        if w not in {"alt", "text", "provided", "no", "###", "notes:", "![picture](picture.jpg)"}
    ]
    return len(useful) > 50


# ---------------------------------------------------------------------------
# OCR extraction — spatially-aware
# ---------------------------------------------------------------------------

# Type alias: each OCR result carries text, confidence, AND bounding box centre
# so downstream functions can reason about spatial proximity.
_OcrFragment = Tuple[str, float, Tuple[float, float]]   # (text, conf, (cx, cy))


def ocr_slide_image(image_path: str) -> List[Tuple[str, float]]:
    """Run EasyOCR on a single slide image.

    Returns:
        List of (text, confidence) tuples, sorted by *spatial reading order*
        (clustered into rows by Y-position, then left-to-right within each row).
    """
    reader = _get_reader()
    results = reader.readtext(image_path)

    if not results:
        return []

    # Compute centre of each bounding box
    enriched: List[_OcrFragment] = []
    for bbox, text, confidence in results:
        xs = [pt[0] for pt in bbox]
        ys = [pt[1] for pt in bbox]
        cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
        enriched.append((text, confidence, (cx, cy)))

    # ── Cluster into rows (fragments with similar Y are on the same line) ──
    enriched.sort(key=lambda r: r[2][1])  # sort by cy
    row_tolerance = _estimate_row_tolerance(enriched)
    rows: List[List[_OcrFragment]] = []
    current_row: List[_OcrFragment] = [enriched[0]]

    for frag in enriched[1:]:
        if abs(frag[2][1] - current_row[0][2][1]) <= row_tolerance:
            current_row.append(frag)
        else:
            rows.append(current_row)
            current_row = [frag]
    rows.append(current_row)

    # Sort each row left-to-right by cx
    for row in rows:
        row.sort(key=lambda r: r[2][0])

    # Flatten back to ordered list
    ordered = [frag for row in rows for frag in row]

    # Store enriched fragments for spatial context (used by metrics/text builder)
    _last_spatial_fragments.clear()
    _last_spatial_fragments.extend(ordered)

    return [(text, conf) for text, conf, _pos in ordered]


# Module-level cache so _parse_metrics_from_ocr can access bounding boxes
# without changing the public signature of ocr_slide_image.
_last_spatial_fragments: List[_OcrFragment] = []


def _estimate_row_tolerance(fragments: List[_OcrFragment]) -> float:
    """Estimate the Y-distance that separates distinct rows.

    Uses the median gap between consecutive fragments as a baseline.
    Fragments within 0.6× the median line-height are on the same row.
    Falls back to 20 px if the image has very few fragments.
    """
    if len(fragments) < 2:
        return 20.0

    ys = sorted(f[2][1] for f in fragments)
    gaps = [ys[i + 1] - ys[i] for i in range(len(ys) - 1) if ys[i + 1] - ys[i] > 0]
    if not gaps:
        return 20.0

    gaps.sort()
    median_gap = gaps[len(gaps) // 2]
    return max(median_gap * 0.6, 8.0)


def _spatial_context(fragments: List[_OcrFragment], idx: int) -> str:
    """Return context from the SAME ROW as the target fragment.

    This ensures a number is paired with labels that are actually beside it
    on the dashboard, not just vertically adjacent.
    """
    if not fragments or idx >= len(fragments):
        return ""

    target_cy = fragments[idx][2][1]
    row_tol = _estimate_row_tolerance(fragments) if len(fragments) > 1 else 20.0

    same_row = [
        f[0] for f in fragments
        if abs(f[2][1] - target_cy) <= row_tol
    ]

    return " | ".join(same_row)


def _parse_metrics_from_ocr(ocr_results: List[Tuple[str, float]]) -> List[Dict]:
    """Extract structured metrics from OCR text fragments.

    Uses spatial bounding-box data (if available) to build context from the
    same visual row rather than just vertical adjacency — preventing
    misattribution of numbers to wrong labels.
    """
    metrics = []
    pct_re = re.compile(r"(-?\d+(?:\.\d+)?)\s*%")
    num_re = re.compile(r"(-?\d+(?:,\d{3})*(?:\.\d+)?)")

    # Prefer spatial fragments; fall back to positional-only if unavailable
    spatial = _last_spatial_fragments if _last_spatial_fragments else None
    all_texts = [text for text, _conf in ocr_results]

    for i, (text, conf) in enumerate(ocr_results):
        if conf < 0.3:
            continue

        context = (
            _spatial_context(spatial, i) if spatial
            else _context_window(all_texts, i)
        )

        # Look for percentages
        for m in pct_re.finditer(text):
            metrics.append({
                "value": m.group(0),
                "numeric_value": float(m.group(1)),
                "context": context,
                "metric_type": "percentage",
            })

        # Look for standalone numbers (not already captured as %)
        stripped = pct_re.sub("", text)
        for m in num_re.finditer(stripped):
            raw = m.group(1).replace(",", "")
            try:
                val = float(raw)
            except ValueError:
                continue
            metrics.append({
                "value": m.group(0),
                "numeric_value": val,
                "context": context,
                "metric_type": "count",
            })

    return metrics


def _extract_key_phrases(ocr_results: List[Tuple[str, float]]) -> List[str]:
    """Build key phrases from high-confidence OCR fragments."""
    phrases = []
    for text, conf in ocr_results:
        if conf < 0.5:
            continue
        cleaned = text.strip()
        # Skip very short or purely numeric fragments
        if len(cleaned) < 3 or cleaned.replace(".", "").replace(",", "").isdigit():
            continue
        phrases.append(cleaned)
    return phrases


def _context_window(texts: List[str], idx: int, window: int = 2) -> str:
    """Return surrounding text fragments as context for a metric (fallback)."""
    start = max(0, idx - window)
    end = min(len(texts), idx + window + 1)
    return " | ".join(texts[start:end])


# ---------------------------------------------------------------------------
# Public API — enrich slides with OCR data
# ---------------------------------------------------------------------------

def _build_spatial_text(fragments: List[_OcrFragment]) -> str:
    """Build readable text from OCR fragments using spatial row grouping.

    Fragments on the same visual row are joined with " · " (middle dot)
    and rows are separated by newlines. This preserves the dashboard layout
    so the assistant can correctly associate numbers with their labels.
    """
    if not fragments:
        return ""

    row_tol = _estimate_row_tolerance(fragments)
    rows: List[List[_OcrFragment]] = []
    current_row: List[_OcrFragment] = [fragments[0]]

    for frag in fragments[1:]:
        if abs(frag[2][1] - current_row[0][2][1]) <= row_tol:
            current_row.append(frag)
        else:
            rows.append(current_row)
            current_row = [frag]
    rows.append(current_row)

    lines = []
    for row in rows:
        row.sort(key=lambda r: r[2][0])
        texts = [f[0] for f in row if f[1] >= 0.3]
        if texts:
            lines.append(" · ".join(texts))

    return "\n".join(lines)


def enrich_slides_with_ocr(slides: List[Dict], force: bool = False) -> int:
    """Run OCR on slides that lack sufficient text-layer data.

    Args:
        slides: List of slide dicts (mutated in-place).
        force:  If True, run OCR on every slide regardless of text quality.

    Returns:
        Number of slides that were OCR-enriched.
    """
    enriched = 0

    for slide in slides:
        image_path = slide.get("image_path", "")
        if not image_path or not os.path.isfile(image_path):
            continue

        text_layer = slide.get("text_layer", "")
        text_metrics = slide.get("text_metrics", [])

        needs_ocr = force or not text_layer_is_sufficient(text_layer, text_metrics)

        if not needs_ocr:
            continue

        slide_num = slide.get("slide_number", "?")
        print(f"  OCR  Slide {slide_num}: running EasyOCR...")

        ocr_results = ocr_slide_image(image_path)

        # Build full text from OCR — use spatial row grouping if available
        if _last_spatial_fragments:
            ocr_text = _build_spatial_text(_last_spatial_fragments)
        else:
            ocr_text = "\n".join(
                text for text, _conf in ocr_results if _conf >= 0.3
            )

        # Merge: prefer OCR text if existing text_layer was poor
        if not text_layer_is_sufficient(text_layer, text_metrics):
            slide["text_layer"] = ocr_text
            slide["text_metrics"] = _parse_metrics_from_ocr(ocr_results)
            slide["text_key_phrases"] = _extract_key_phrases(ocr_results)
        else:
            # Append OCR as supplementary
            slide["text_layer"] = text_layer + "\n--- OCR ---\n" + ocr_text

        slide["ocr_used"] = True
        enriched += 1
        print(f"  OK  Slide {slide_num}: {len(ocr_results)} text fragments extracted")

    return enriched
