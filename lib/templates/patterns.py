"""
Pattern Library - Layer 3B

Provides headline and insight patterns based on slide types and gap patterns.
Inspired by Example-Storyboard-Analytics.pptx.
"""

from typing import Dict, List
from lib.templates.classifiers import SlideType


class PatternLibrary:
    """Library of headline and insight patterns"""

    # Headline patterns by slide type and gap type
    HEADLINE_PATTERNS = {
        (SlideType.TREND, 'adoption'): "{pct}% penetration reveals untapped {opportunity} across {population}",
        (SlideType.TREND, 'growth'): "{pct}% growth validates {capability} but signals acceleration needed",
        (SlideType.HABIT_FORMATION, 'habit'): "{pct}% remain in {tier} tier - workflows haven't embedded platform",
        (SlideType.LEADERBOARD, 'adoption'): "Top {count} users drive {pct}% of activity - reveals adoption concentration",
        (SlideType.LICENSE_PRIORITY, 'license'): "{count} high-value users demonstrate upgrade readiness in {segment}",
        (SlideType.HEALTH_CHECK, 'adoption'): "{metric} across {population} indicates {gap_description}",
        (SlideType.GEOGRAPHIC, 'adoption'): "{location} shows {pct}% {metric} - reveals regional disparity",
    }

    # Insight bullet patterns
    INSIGHT_PATTERNS = {
        'scale_scope': [
            "{number} {metric} represents {interpretation}",
            "{population} demonstrates {pattern} with {number} {metric}",
            "Analysis reveals {number} {metric} across {population}",
        ],
        'gap_implication': [
            "Reveals {root_cause} - {business_impact}",
            "Indicates {root_cause} requiring {action_type}",
            "Signals {business_impact} from {root_cause}",
            "{gap_magnitude} gap demonstrates {business_impact}",
        ],
        'action_recommendation': [
            "Action: {action_required}",
            "Requires: {action_required}",
            "Recommendation: {action_required}",
            "Next step: {action_required}",
        ],
        'secondary_insight': [
            "Additionally, {number} {description} - {action}",
            "Further analysis shows {number} {description}",
            "{number} {description} validates {implication}",
        ]
    }

    # Context-specific vocabulary
    VOCABULARY = {
        'adoption_gaps': {
            'opportunity': ['platform value', 'user potential', 'capability', 'investment'],
            'population': ['organization', 'user base', 'departments', 'teams'],
            'gap_description': ['massive awareness gap', 'untapped potential', 'adoption opportunity'],
        },
        'habit_gaps': {
            'tier': ['infrequent', 'light usage', 'occasional'],
            'workflow_issues': ['not embedded', 'insufficient integration', 'limited workflow adoption'],
        },
        'license_gaps': {
            'segment': ['power users', 'high-value cohort', 'engaged population', 'priority segment'],
            'opportunity': ['revenue potential', 'upgrade candidates', 'license optimization'],
        }
    }

    def get_headline_pattern(self, slide_type: SlideType, gap_type: str) -> str:
        """
        Get headline pattern for slide type and gap type.

        Args:
            slide_type: Type of slide
            gap_type: Type of gap detected

        Returns:
            Headline pattern template
        """
        # Try to find specific pattern
        pattern = self.HEADLINE_PATTERNS.get((slide_type, gap_type))

        if pattern:
            return pattern

        # Fallback patterns
        fallback_patterns = {
            'adoption': "{number} reveals {gap_description} requiring intervention",
            'habit': "{number} indicates workflow integration gap",
            'license': "{number} represents upgrade opportunity",
            'growth': "{number} validates momentum but signals untapped potential",
        }

        return fallback_patterns.get(gap_type, "{number} signals strategic opportunity")

    def get_insight_bullets(
        self,
        pattern_type: str,
        context: Dict[str, str]
    ) -> str:
        """
        Get insight bullet pattern filled with context.

        Args:
            pattern_type: Type of insight ('scale_scope', 'gap_implication', 'action_recommendation', 'secondary_insight')
            context: Dictionary of context variables to fill in

        Returns:
            Formatted insight bullet
        """
        patterns = self.INSIGHT_PATTERNS.get(pattern_type, [])

        if not patterns:
            return ""

        # Select first pattern (could be randomized for variety)
        pattern = patterns[0]

        # Fill in context variables
        try:
            return pattern.format(**context)
        except KeyError:
            # Missing context variables, return pattern as-is
            return pattern

    def enrich_context(self, gap_type: str, base_context: Dict[str, str]) -> Dict[str, str]:
        """
        Enrich context with vocabulary appropriate to gap type.

        Args:
            gap_type: Type of gap
            base_context: Base context dictionary

        Returns:
            Enriched context
        """
        enriched = base_context.copy()

        # Add vocabulary based on gap type
        if gap_type == 'adoption':
            vocab = self.VOCABULARY['adoption_gaps']
            enriched.setdefault('opportunity', vocab['opportunity'][0])
            enriched.setdefault('population', vocab['population'][0])

        elif gap_type == 'habit':
            vocab = self.VOCABULARY['habit_gaps']
            enriched.setdefault('tier', vocab['tier'][0])

        elif gap_type == 'license':
            vocab = self.VOCABULARY['license_gaps']
            enriched.setdefault('segment', vocab['segment'][0])

        return enriched


class HeadlineGenerator:
    """Generates compelling headlines using pattern library"""

    def __init__(self):
        self.library = PatternLibrary()

    def generate(
        self,
        slide_type: SlideType,
        gap_type: str,
        context: Dict[str, str]
    ) -> str:
        """
        Generate headline for slide.

        Args:
            slide_type: Type of slide
            gap_type: Type of gap detected
            context: Context variables (numbers, metrics, etc.)

        Returns:
            Generated headline
        """
        pattern = self.library.get_headline_pattern(slide_type, gap_type)

        # Enrich context with vocabulary
        enriched_context = self.library.enrich_context(gap_type, context)

        # Fill in pattern
        try:
            return pattern.format(**enriched_context)
        except KeyError:
            # Missing variables, return simpler version
            number = context.get('number', context.get('pct', context.get('count', '')))
            return f"{number} signals strategic {gap_type} opportunity"


class InsightBulletGenerator:
    """Generates insight bullets using pattern library"""

    def __init__(self):
        self.library = PatternLibrary()

    def generate_bullets(
        self,
        gap_type: str,
        context: Dict[str, str]
    ) -> List[str]:
        """
        Generate 3 insight bullets.

        Args:
            gap_type: Type of gap
            context: Context variables

        Returns:
            List of 3 bullet points
        """
        enriched_context = self.library.enrich_context(gap_type, context)

        bullets = []

        # Bullet 1: Scale/scope
        bullets.append(
            self.library.get_insight_bullets('scale_scope', enriched_context)
        )

        # Bullet 2: Gap/implication
        bullets.append(
            self.library.get_insight_bullets('gap_implication', enriched_context)
        )

        # Bullet 3: Action recommendation
        bullets.append(
            self.library.get_insight_bullets('action_recommendation', enriched_context)
        )

        return bullets
