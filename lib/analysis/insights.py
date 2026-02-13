"""
Insight Composition - Layer 2C

Composes compelling insights following the formula:
[Number] + [Gap] + [Implication] + [Action]
"""

from typing import Dict, List, Tuple
from dataclasses import dataclass
from lib.extraction.extractor import SlideData
from lib.analysis.gaps import Gap, detect_all_gaps
from lib.analysis.implications import Implication, map_gaps_to_implications


@dataclass
class Insight:
    """A compelling, actionable insight"""
    headline: str  # Short, punchy headline
    bullet_points: List[str]  # 2-3 supporting insights
    source_numbers: List[str]  # Numbers used (for validation)


class InsightComposer:
    """Composes insights from implications following the formula"""

    def compose_slide_insights(
        self,
        slide_data: SlideData,
        implications: List[Implication]
    ) -> Insight:
        """
        Compose a complete insight for a slide.

        Args:
            slide_data: Extracted slide data
            implications: Business implications for the slide

        Returns:
            Composed insight with headline and bullets
        """
        # Generate headline from primary implication
        headline = self._generate_headline(slide_data, implications)

        # Generate 2-3 bullet points
        bullet_points = self._generate_bullets(slide_data, implications)

        # Track source numbers for validation
        source_numbers = self._extract_source_numbers(implications)

        return Insight(
            headline=headline,
            bullet_points=bullet_points,
            source_numbers=source_numbers
        )

    def _generate_headline(
        self,
        slide_data: SlideData,
        implications: List[Implication]
    ) -> str:
        """Generate insight-driven headline with specific number"""
        if not implications:
            # Fallback: Extract largest number from slide
            if slide_data.metrics:
                largest_metric = max(slide_data.metrics, key=lambda m: m.numeric_value)
                return f"{largest_metric.value} {slide_data.title}"
            return slide_data.title

        # Use primary (highest priority) implication
        primary = implications[0]
        gap = primary.gap

        # Format number
        number_str = self._format_number(gap.source_metric.value)

        # Compose headline: [Number] + [Gap insight]
        if gap.gap_type == 'adoption':
            return f"{number_str} {gap.description.split(' - ')[0]}"
        elif gap.gap_type == 'habit':
            return f"{number_str} {gap.description.split(' - ')[0]}"
        elif gap.gap_type == 'license':
            return f"{number_str} represent untapped upgrade opportunity"
        elif gap.gap_type == 'growth':
            return f"{number_str} {gap.description.split(' validates ')[0]}"
        else:
            return f"{number_str} signals strategic opportunity"

    def _generate_bullets(
        self,
        slide_data: SlideData,
        implications: List[Implication]
    ) -> List[str]:
        """Generate 2-3 supporting bullet points"""
        bullets = []

        if not implications:
            # Fallback: Use top metrics
            for metric in slide_data.metrics[:3]:
                bullets.append(f"{metric.value} - {metric.context[:60]}")
            return bullets

        # Bullet 1: Scale/scope with specific number
        primary = implications[0]
        gap = primary.gap
        bullets.append(
            f"{self._format_number(gap.source_metric.value)} {gap.description}"
        )

        # Bullet 2: Gap identification with business implication
        if len(implications) > 0:
            bullets.append(
                f"Reveals {primary.root_cause} - {primary.business_impact}"
            )

        # Bullet 3: Action recommendation
        if len(implications) > 0:
            bullets.append(
                f"Action: {primary.action_required}"
            )

        # If we have secondary implications, replace bullet 3
        if len(implications) > 1:
            secondary = implications[1]
            bullets[2] = (
                f"Additionally, {self._format_number(secondary.gap.source_metric.value)} "
                f"{secondary.gap.description.split(' - ')[0]} - {secondary.action_required}"
            )

        return bullets[:3]  # Max 3 bullets

    def _format_number(self, value_str: str) -> str:
        """Format number for readability"""
        # Already formatted from extraction (e.g., "4%", "1,275", "87K")
        return value_str

    def _extract_source_numbers(self, implications: List[Implication]) -> List[str]:
        """Extract all source numbers for validation"""
        return [impl.gap.source_metric.value for impl in implications]


def generate_insights(data: Dict[str, SlideData]) -> Dict[str, Insight]:
    """
    Main entry point: Generate insights for all slides.

    Args:
        data: Dictionary of slide titles to SlideData

    Returns:
        Dictionary of slide titles to Insights
    """
    # Step 1: Detect gaps
    gaps_by_slide = detect_all_gaps(data)

    # Step 2: Map to implications
    implications_by_slide = map_gaps_to_implications(gaps_by_slide)

    # Step 3: Compose insights
    composer = InsightComposer()
    insights = {}

    for title, slide_data in data.items():
        implications = implications_by_slide.get(title, [])
        insight = composer.compose_slide_insights(slide_data, implications)
        insights[title] = insight

    return insights
