"""
Text-layer extraction helpers for Copilot mode.

Uses embedded text (PDF/PPTX) without OCR.
"""

from __future__ import annotations

from typing import Dict, List

from lib.extraction.extractor import DashboardExtractor


def sanitize_text(text: str) -> str:
    """Force ASCII output to avoid Unicode write errors."""
    if not text:
        return ""
    return text.encode("ascii", errors="replace").decode("ascii")


def enrich_slides_with_pptx_text(source_path: str, slides: List[Dict]) -> None:
    """Populate text_layer and metrics for PPTX slides using markitdown."""
    extractor = DashboardExtractor()
    slide_data = extractor.extract_from_file(source_path)

    by_number: Dict[int, object] = {}
    for data in slide_data.values():
        by_number[data.slide_number] = data

    for slide in slides:
        data = by_number.get(slide.get("slide_number"))
        if data is None:
            slide["text_layer"] = ""
            slide["text_metrics"] = []
            slide["text_key_phrases"] = []
            continue

        combined_text = f"{data.title}\n{data.text_content}".strip()
        combined_text = sanitize_text(combined_text)
        slide["text_layer"] = combined_text
        slide["text_metrics"] = [
            {
                "value": m.value,
                "numeric_value": m.numeric_value,
                "context": m.context,
                "metric_type": m.metric_type,
            }
            for m in data.metrics
        ]
        slide["text_key_phrases"] = data.key_phrases
