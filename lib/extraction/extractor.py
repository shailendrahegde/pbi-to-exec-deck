"""
Data Extraction Engine - Layer 1

Extracts structured data from Power BI dashboard exports using markitdown and regex.
Parses numbers, metrics, and context without requiring OCR.
"""

import re
from typing import Dict, List, Tuple, Any
from markitdown import MarkItDown
from dataclasses import dataclass


@dataclass
class ExtractedMetric:
    """Represents a metric extracted from the dashboard"""
    value: str  # Raw value as string (e.g., "1,275", "4%", "87.5K")
    numeric_value: float  # Parsed numeric value
    context: str  # Surrounding text context
    metric_type: str  # "percentage", "count", "ratio", etc.


@dataclass
class SlideData:
    """Represents extracted data from a single slide"""
    slide_number: int
    title: str
    metrics: List[ExtractedMetric]
    text_content: str
    key_phrases: List[str]


class DashboardExtractor:
    """Extracts data from Power BI dashboard PowerPoint exports"""

    # Regex patterns for number extraction
    PATTERNS = {
        'percentage': re.compile(r'(\d+(?:\.\d+)?)\s*%'),
        'large_number_suffix': re.compile(r'(\d+(?:\.\d+)?)\s*([KMB])\b', re.IGNORECASE),
        'comma_number': re.compile(r'\b(\d{1,3}(?:,\d{3})+)\b'),
        'decimal_number': re.compile(r'\b(\d+\.\d+)\b'),
        'plain_number': re.compile(r'\b(\d+)\b'),
    }

    MULTIPLIERS = {
        'K': 1_000,
        'M': 1_000_000,
        'B': 1_000_000_000,
    }

    def __init__(self):
        self.md_converter = MarkItDown()

    def extract_from_file(self, pptx_path: str) -> Dict[str, SlideData]:
        """
        Extract structured data from PowerPoint file.

        Args:
            pptx_path: Path to .pptx file

        Returns:
            Dictionary mapping slide titles to SlideData objects
        """
        # Extract text content using markitdown
        result = self.md_converter.convert(pptx_path)
        text_content = result.text_content

        # Parse slides from markdown output
        slides = self._parse_slides(text_content)

        # Extract metrics from each slide
        slide_data = {}
        for slide in slides:
            data = self._extract_slide_data(slide)
            if data and data.title:  # Only include slides with titles
                slide_data[data.title] = data

        return slide_data

    def _parse_slides(self, markdown_text: str) -> List[Dict[str, Any]]:
        """Parse markdown text into individual slides"""
        slides = []
        current_slide = None

        for line in markdown_text.split('\n'):
            # Detect slide boundaries (format: <!-- Slide number: 1 -->)
            if 'Slide number:' in line:
                if current_slide:
                    slides.append(current_slide)
                slide_num_match = re.search(r'Slide number:\s*(\d+)', line)
                if slide_num_match:
                    slide_num = int(slide_num_match.group(1))
                    current_slide = {
                        'number': slide_num,
                        'title': '',
                        'content': []
                    }
            elif current_slide is not None:
                # Look for markdown heading as title (# Title)
                if line.strip().startswith('#') and not current_slide['title']:
                    # Remove markdown heading markers and emoji
                    title = line.strip().lstrip('#').strip()
                    # Remove emoji characters
                    title = re.sub(r'[\U00010000-\U0010ffff]', '', title).strip()
                    current_slide['title'] = title
                else:
                    current_slide['content'].append(line)

        if current_slide:
            slides.append(current_slide)

        return slides

    def _extract_slide_data(self, slide_dict: Dict[str, Any]) -> SlideData:
        """Extract structured data from a single slide"""
        content = '\n'.join(slide_dict['content'])
        full_text = f"{slide_dict['title']}\n{content}"

        # Extract all metrics
        metrics = self._extract_metrics(full_text)

        # Extract key phrases (potential insight keywords)
        key_phrases = self._extract_key_phrases(full_text)

        return SlideData(
            slide_number=slide_dict['number'],
            title=slide_dict['title'],
            metrics=metrics,
            text_content=content,
            key_phrases=key_phrases
        )

    def _extract_metrics(self, text: str) -> List[ExtractedMetric]:
        """Extract all numeric metrics from text"""
        metrics = []

        # Extract percentages
        for match in self.PATTERNS['percentage'].finditer(text):
            value = match.group(1)
            context = self._get_context(text, match.span())
            metrics.append(ExtractedMetric(
                value=f"{value}%",
                numeric_value=float(value),
                context=context,
                metric_type='percentage'
            ))

        # Extract large numbers with K/M/B suffix
        for match in self.PATTERNS['large_number_suffix'].finditer(text):
            value = match.group(1)
            suffix = match.group(2).upper()
            multiplier = self.MULTIPLIERS[suffix]
            context = self._get_context(text, match.span())
            metrics.append(ExtractedMetric(
                value=match.group(0),
                numeric_value=float(value) * multiplier,
                context=context,
                metric_type='count'
            ))

        # Extract comma-formatted numbers
        for match in self.PATTERNS['comma_number'].finditer(text):
            value = match.group(1)
            numeric_value = float(value.replace(',', ''))
            context = self._get_context(text, match.span())

            # Skip if already captured as part of another metric
            if not any(m.numeric_value == numeric_value for m in metrics):
                metrics.append(ExtractedMetric(
                    value=value,
                    numeric_value=numeric_value,
                    context=context,
                    metric_type='count'
                ))

        # Extract decimal numbers
        for match in self.PATTERNS['decimal_number'].finditer(text):
            value = match.group(1)
            numeric_value = float(value)
            context = self._get_context(text, match.span())

            # Skip if already captured
            if not any(abs(m.numeric_value - numeric_value) < 0.01 for m in metrics):
                metrics.append(ExtractedMetric(
                    value=value,
                    numeric_value=numeric_value,
                    context=context,
                    metric_type='decimal'
                ))

        # Extract plain numbers (as last resort, for numbers not caught above)
        for match in self.PATTERNS['plain_number'].finditer(text):
            value = match.group(0)
            try:
                numeric_value = float(value)

                # Skip small numbers (likely noise like years, IDs, etc.) and already captured
                if numeric_value < 10:
                    continue

                if any(abs(m.numeric_value - numeric_value) < 0.01 for m in metrics):
                    continue

                context = self._get_context(text, match.span())

                # Only include if context suggests it's a meaningful metric
                context_lower = context.lower()
                if any(keyword in context_lower for keyword in ['users', 'license', 'value', 'count', 'total', 'active']):
                    metrics.append(ExtractedMetric(
                        value=value,
                        numeric_value=numeric_value,
                        context=context,
                        metric_type='count'
                    ))
            except ValueError:
                continue

        return metrics

    def _get_context(self, text: str, span: Tuple[int, int], window: int = 50) -> str:
        """Extract surrounding context for a matched pattern"""
        start, end = span
        context_start = max(0, start - window)
        context_end = min(len(text), end + window)

        context = text[context_start:context_end].strip()
        return context

    def _extract_key_phrases(self, text: str) -> List[str]:
        """Extract key phrases that might indicate insight opportunities"""
        key_terms = [
            'active users', 'adoption', 'growth', 'increase', 'decrease',
            'license', 'penetration', 'engagement', 'frequency', 'tier',
            'department', 'location', 'trend', 'leader', 'high-value',
            'opportunity', 'gap', 'training', 'awareness', 'habit',
            'workflow', 'integration', 'upgrade', 'intervention'
        ]

        phrases = []
        text_lower = text.lower()

        for term in key_terms:
            if term in text_lower:
                # Extract sentence containing the term
                sentences = text.split('.')
                for sentence in sentences:
                    if term in sentence.lower():
                        phrases.append(sentence.strip())
                        break

        return phrases


def extract_dashboard_data(source_path: str) -> Dict[str, SlideData]:
    """
    Main entry point for data extraction.

    Args:
        source_path: Path to source PowerPoint file

    Returns:
        Dictionary mapping slide titles to extracted data
    """
    extractor = DashboardExtractor()
    return extractor.extract_from_file(source_path)
