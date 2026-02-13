"""
Business Implication Mapping - Layer 2B

Maps detected gaps to business implications and required actions.
"""

from typing import Dict, List
from dataclasses import dataclass
from lib.analysis.gaps import Gap


@dataclass
class Implication:
    """Represents a business implication with recommended action"""
    gap: Gap
    business_impact: str  # What this means for the business
    root_cause: str  # Likely cause of the gap
    action_required: str  # What needs to be done
    urgency: str  # "immediate", "high", "moderate"


class ImplicationMapper:
    """Maps gaps to business implications and actions"""

    # Implication templates by gap type
    IMPLICATION_RULES = {
        'adoption': {
            'business_impact': 'unrealized value from underutilized platform investment',
            'root_cause': 'awareness or training gap',
            'action_template': 'requires immediate awareness campaign and targeted onboarding',
            'urgency': 'immediate'
        },
        'habit': {
            'business_impact': 'platform not embedded in daily workflows',
            'root_cause': 'workflow integration failure or insufficient use case alignment',
            'action_template': 'indicates need for workflow integration training and champion program',
            'urgency': 'high'
        },
        'license': {
            'business_impact': 'revenue opportunity and user satisfaction risk',
            'root_cause': 'license allocation not aligned with usage patterns',
            'action_template': 'validates license upgrade targeting for high-value segments',
            'urgency': 'moderate'
        },
        'growth': {
            'business_impact': 'momentum exists but ceiling not reached',
            'root_cause': 'organic growth without systematic acceleration',
            'action_template': 'signals opportunity to accelerate with targeted enablement',
            'urgency': 'moderate'
        }
    }

    def map_gap_to_implication(self, gap: Gap) -> Implication:
        """
        Map a single gap to business implication.

        Args:
            gap: Detected gap

        Returns:
            Implication with business context and action
        """
        rules = self.IMPLICATION_RULES.get(gap.gap_type, self.IMPLICATION_RULES['adoption'])

        # Customize action based on gap magnitude
        action = self._customize_action(gap, rules['action_template'])

        return Implication(
            gap=gap,
            business_impact=rules['business_impact'],
            root_cause=rules['root_cause'],
            action_required=action,
            urgency=rules['urgency']
        )

    def _customize_action(self, gap: Gap, action_template: str) -> str:
        """Customize action based on gap specifics"""
        # For very large gaps, emphasize urgency
        if gap.gap_type == 'adoption' and gap.magnitude > 90:
            return f"requires immediate intervention - {action_template}"

        # For license opportunities with large numbers, emphasize ROI
        if gap.gap_type == 'license' and gap.magnitude > 500:
            return f"represents significant revenue opportunity - {action_template}"

        # For habit gaps, emphasize workflow integration
        if gap.gap_type == 'habit' and gap.magnitude > 80:
            return f"critical workflow integration needed - {action_template}"

        return action_template

    def prioritize_implications(self, implications: List[Implication]) -> List[Implication]:
        """
        Prioritize implications by urgency and magnitude.

        Args:
            implications: List of implications

        Returns:
            Sorted list with highest priority first
        """
        urgency_order = {'immediate': 0, 'high': 1, 'moderate': 2}

        return sorted(
            implications,
            key=lambda i: (urgency_order.get(i.urgency, 3), -i.gap.magnitude)
        )


def map_gaps_to_implications(gaps_by_slide: Dict[str, List[Gap]]) -> Dict[str, List[Implication]]:
    """
    Map all gaps to implications.

    Args:
        gaps_by_slide: Dictionary of slide titles to gaps

    Returns:
        Dictionary of slide titles to implications
    """
    mapper = ImplicationMapper()
    implications_by_slide = {}

    for title, gaps in gaps_by_slide.items():
        implications = [mapper.map_gap_to_implication(gap) for gap in gaps]
        # Prioritize implications
        implications = mapper.prioritize_implications(implications)
        implications_by_slide[title] = implications

    return implications_by_slide
