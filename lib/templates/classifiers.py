"""
Slide Type Classification - Layer 3A

Classifies dashboard slides by type to apply appropriate insight patterns.
"""

from typing import Optional
from enum import Enum


class SlideType(Enum):
    """Types of dashboard slides"""
    TREND = "trend"  # Time-series, growth over time
    LEADERBOARD = "leaderboard"  # Top users, departments, rankings
    HEALTH_CHECK = "health_check"  # Portfolio overview, KPI summary
    HABIT_FORMATION = "habit_formation"  # Engagement tiers, frequency distribution
    LICENSE_PRIORITY = "license_priority"  # License allocation, upgrade candidates
    GEOGRAPHIC = "geographic"  # Location-based analysis
    UNKNOWN = "unknown"  # Couldn't classify


class SlideClassifier:
    """Classifies slides by their content type"""

    # Keywords that indicate slide types
    TYPE_KEYWORDS = {
        SlideType.TREND: [
            'trend', 'over time', 'growth', 'timeline', 'history',
            'month', 'week', 'daily', 'progression', 'trajectory'
        ],
        SlideType.LEADERBOARD: [
            'top', 'leader', 'ranking', 'most active', 'power users',
            'champion', 'by user', 'by department', 'leaderboard'
        ],
        SlideType.HEALTH_CHECK: [
            'health', 'overview', 'summary', 'portfolio', 'dashboard',
            'kpi', 'metrics', 'snapshot', 'status'
        ],
        SlideType.HABIT_FORMATION: [
            'frequency', 'engagement', 'tier', 'habit', 'usage pattern',
            'infrequent', 'heavy', 'light', 'occasional', 'daily',
            'weekly', 'monthly', 'distribution'
        ],
        SlideType.LICENSE_PRIORITY: [
            'license', 'premium', 'unlicensed', 'subscription', 'upgrade',
            'priority', 'high-value', 'allocation', 'assignment'
        ],
        SlideType.GEOGRAPHIC: [
            'location', 'geography', 'region', 'country', 'city',
            'state', 'territory', 'by location'
        ]
    }

    def classify(self, slide_title: str, slide_content: str = "") -> SlideType:
        """
        Classify a slide based on title and content.

        Args:
            slide_title: The slide title
            slide_content: Optional slide content text

        Returns:
            Detected SlideType
        """
        text = f"{slide_title} {slide_content}".lower()

        # Count keyword matches for each type
        scores = {}
        for slide_type, keywords in self.TYPE_KEYWORDS.items():
            score = sum(1 for keyword in keywords if keyword in text)
            scores[slide_type] = score

        # Return type with highest score (if any matches)
        if scores:
            best_type = max(scores.items(), key=lambda x: x[1])
            if best_type[1] > 0:
                return best_type[0]

        return SlideType.UNKNOWN

    def get_slide_focus(self, slide_type: SlideType) -> str:
        """
        Get the analytical focus for a slide type.

        Args:
            slide_type: Detected slide type

        Returns:
            Description of what to focus on
        """
        focus_map = {
            SlideType.TREND: "trajectory, growth rate, and momentum gaps",
            SlideType.LEADERBOARD: "power user patterns and adoption disparities",
            SlideType.HEALTH_CHECK: "portfolio health and utilization gaps",
            SlideType.HABIT_FORMATION: "engagement tiers and workflow integration",
            SlideType.LICENSE_PRIORITY: "upgrade candidates and revenue opportunity",
            SlideType.GEOGRAPHIC: "regional disparities and location-based patterns",
            SlideType.UNKNOWN: "key metrics and strategic opportunities"
        }
        return focus_map.get(slide_type, focus_map[SlideType.UNKNOWN])
