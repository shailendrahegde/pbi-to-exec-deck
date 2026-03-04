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
# OCR extraction
# ---------------------------------------------------------------------------

def ocr_slide_image(image_path: str) -> List[Tuple[str, float]]:
    """Run EasyOCR on a single slide image.

    Returns:
        List of (text, confidence) tuples, sorted top-to-bottom by bounding box.
    """
    reader = _get_reader()
    results = reader.readtext(image_path)

    # Sort by vertical position (top of bounding box)
    results.sort(key=lambda r: r[0][0][1])

    return [(text, confidence) for (_bbox, text, confidence) in results]


def _parse_metrics_from_ocr(ocr_results: List[Tuple[str, float]]) -> List[Dict]:
    """Extract structured metrics from OCR text fragments."""
    metrics = []
    pct_re = re.compile(r"(-?\d+(?:\.\d+)?)\s*%")
    num_re = re.compile(r"(-?\d+(?:,\d{3})*(?:\.\d+)?)")

    all_texts = [text for text, _conf in ocr_results]

    for i, (text, conf) in enumerate(ocr_results):
        if conf < 0.3:
            continue

        # Look for percentages
        for m in pct_re.finditer(text):
            context = _context_window(all_texts, i)
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
            context = _context_window(all_texts, i)
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
    """Return surrounding text fragments as context for a metric."""
    start = max(0, idx - window)
    end = min(len(texts), idx + window + 1)
    return " | ".join(texts[start:end])


# ---------------------------------------------------------------------------
# Public API — enrich slides with OCR data
# ---------------------------------------------------------------------------

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

        # Build full text from OCR
        ocr_text = "\n".join(text for text, _conf in ocr_results if _conf >= 0.3)

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
