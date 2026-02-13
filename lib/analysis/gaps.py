"""
Gap Detection - Layer 2A

Identifies adoption gaps, habit formation gaps, and license opportunities
from extracted metrics.
"""

from typing import Dict, List, Optional
from dataclasses import dataclass
from lib.extraction.extractor import ExtractedMetric, SlideData


@dataclass
class Gap:
    """Represents an identified gap with business impact"""
    gap_type: str  # "adoption", "habit", "license", "growth"
    magnitude: float  # Size of the gap (percentage or count)
    description: str  # Human-readable gap description
    source_metric: ExtractedMetric  # The metric that revealed the gap
    opportunity_size: float  # Quantified opportunity


class GapDetector:
    """Detects strategic gaps in dashboard data"""

    # Thresholds for gap detection
    LOW_ADOPTION_THRESHOLD = 10  # < 10% is low adoption
    LOW_FREQUENCY_THRESHOLD = 70  # > 70% infrequent is a habit gap
    HIGH_VALUE_UNLICENSED_THRESHOLD = 100  # > 100 unlicensed is opportunity

    def detect_gaps(self, slide_data: SlideData) -> List[Gap]:
        """
        Detect all gaps in a slide's data.

        Args:
            slide_data: Extracted slide data

        Returns:
            List of identified gaps
        """
        gaps = []

        # Try different gap detection strategies
        gaps.extend(self._detect_adoption_gaps(slide_data))
        gaps.extend(self._detect_habit_gaps(slide_data))
        gaps.extend(self._detect_license_gaps(slide_data))
        gaps.extend(self._detect_growth_gaps(slide_data))

        return gaps

    def _detect_adoption_gaps(self, slide_data: SlideData) -> List[Gap]:
        """Detect low adoption / penetration gaps"""
        gaps = []

        # Look for low penetration percentages
        for metric in slide_data.metrics:
            if metric.metric_type == 'percentage':
                # Low adoption indicators
                if metric.numeric_value < self.LOW_ADOPTION_THRESHOLD:
                    context_lower = metric.context.lower()

                    # Check if this is an adoption metric
                    adoption_keywords = ['adopt', 'penetration', 'usage', 'active', 'enabled']
                    if any(keyword in context_lower for keyword in adoption_keywords):
                        gap_size = 100 - metric.numeric_value

                        gaps.append(Gap(
                            gap_type='adoption',
                            magnitude=gap_size,
                            description=f"{gap_size:.0f}% of users have not adopted - massive awareness and adoption gap",
                            source_metric=metric,
                            opportunity_size=gap_size
                        ))

        return gaps

    def _detect_habit_gaps(self, slide_data: SlideData) -> List[Gap]:
        """Detect habit formation and workflow integration gaps"""
        gaps = []

        # Look for high percentages in "infrequent" or "light" tiers
        for metric in slide_data.metrics:
            if metric.metric_type == 'percentage':
                context_lower = metric.context.lower()

                # High infrequent usage is a habit gap
                habit_keywords = ['infrequent', 'light', 'occasional', 'low frequency']
                if any(keyword in context_lower for keyword in habit_keywords):
                    if metric.numeric_value > self.LOW_FREQUENCY_THRESHOLD:
                        gaps.append(Gap(
                            gap_type='habit',
                            magnitude=metric.numeric_value,
                            description=f"{metric.numeric_value:.0f}% remain in infrequent tier - workflows haven't embedded platform",
                            source_metric=metric,
                            opportunity_size=metric.numeric_value
                        ))

        # Look for low "heavy" or "frequent" usage
        for metric in slide_data.metrics:
            if metric.metric_type == 'percentage':
                context_lower = metric.context.lower()

                engagement_keywords = ['heavy', 'frequent', 'power', 'daily']
                if any(keyword in context_lower for keyword in engagement_keywords):
                    if metric.numeric_value < 20:  # Less than 20% heavy users
                        gap_size = 100 - metric.numeric_value
                        gaps.append(Gap(
                            gap_type='habit',
                            magnitude=gap_size,
                            description=f"Only {metric.numeric_value:.0f}% are frequent users - {gap_size:.0f}% need workflow integration",
                            source_metric=metric,
                            opportunity_size=gap_size
                        ))

        return gaps

    def _detect_license_gaps(self, slide_data: SlideData) -> List[Gap]:
        """Detect license upgrade opportunities"""
        gaps = []

        # Look for unlicensed or non-premium users with high value
        for metric in slide_data.metrics:
            context_lower = metric.context.lower()

            license_keywords = ['unlicensed', 'non-premium', 'free tier', 'no license', 'high-value']
            if any(keyword in context_lower for keyword in license_keywords):
                # High count of unlicensed valuable users
                if metric.metric_type == 'count' and metric.numeric_value > self.HIGH_VALUE_UNLICENSED_THRESHOLD:
                    gaps.append(Gap(
                        gap_type='license',
                        magnitude=metric.numeric_value,
                        description=f"{metric.numeric_value:.0f} high-value unlicensed users represent upgrade opportunity",
                        source_metric=metric,
                        opportunity_size=metric.numeric_value
                    ))

        return gaps

    def _detect_growth_gaps(self, slide_data: SlideData) -> List[Gap]:
        """Detect growth trends that mask underlying gaps"""
        gaps = []

        # Look for growth percentages that might hide gaps
        for metric in slide_data.metrics:
            if metric.metric_type == 'percentage':
                context_lower = metric.context.lower()

                growth_keywords = ['growth', 'increase', 'up', 'gain']
                if any(keyword in context_lower for keyword in growth_keywords):
                    # Moderate growth might indicate untapped potential
                    if 10 < metric.numeric_value < 50:
                        gaps.append(Gap(
                            gap_type='growth',
                            magnitude=100 - metric.numeric_value,
                            description=f"{metric.numeric_value:.0f}% growth validates capability but signals {100-metric.numeric_value:.0f}% untapped potential",
                            source_metric=metric,
                            opportunity_size=100 - metric.numeric_value
                        ))

        return gaps


def detect_all_gaps(data: Dict[str, SlideData]) -> Dict[str, List[Gap]]:
    """
    Detect gaps across all slides.

    Args:
        data: Dictionary of slide titles to SlideData

    Returns:
        Dictionary of slide titles to detected gaps
    """
    detector = GapDetector()
    gaps_by_slide = {}

    for title, slide_data in data.items():
        gaps = detector.detect_gaps(slide_data)
        if gaps:
            gaps_by_slide[title] = gaps

    return gaps_by_slide
