"""
Unit tests for insight generation layer
"""

import unittest
from lib.extraction.extractor import ExtractedMetric, SlideData
from lib.analysis.gaps import GapDetector, Gap
from lib.analysis.implications import ImplicationMapper
from lib.analysis.insights import InsightComposer


class TestGapDetection(unittest.TestCase):
    """Test gap detection logic"""

    def setUp(self):
        self.detector = GapDetector()

    def test_adoption_gap_detection(self):
        """Test detection of low adoption gaps"""
        metric = ExtractedMetric(
            value="4%",
            numeric_value=4.0,
            context="4% of users have adopted agents",
            metric_type="percentage"
        )

        slide_data = SlideData(
            slide_number=1,
            title="Agent Adoption",
            metrics=[metric],
            text_content="4% adoption rate",
            key_phrases=["adoption"]
        )

        gaps = self.detector.detect_gaps(slide_data)

        # Should detect 96% adoption gap
        self.assertTrue(len(gaps) > 0)
        adoption_gaps = [g for g in gaps if g.gap_type == 'adoption']
        self.assertTrue(len(adoption_gaps) > 0)
        self.assertAlmostEqual(adoption_gaps[0].magnitude, 96.0)

    def test_habit_gap_detection(self):
        """Test detection of habit formation gaps"""
        metric = ExtractedMetric(
            value="87%",
            numeric_value=87.0,
            context="87% remain in infrequent tier",
            metric_type="percentage"
        )

        slide_data = SlideData(
            slide_number=2,
            title="Frequency Tiers",
            metrics=[metric],
            text_content="Most users are infrequent",
            key_phrases=["infrequent", "frequency"]
        )

        gaps = self.detector.detect_gaps(slide_data)

        # Should detect habit formation gap
        habit_gaps = [g for g in gaps if g.gap_type == 'habit']
        self.assertTrue(len(habit_gaps) > 0)

    def test_license_gap_detection(self):
        """Test detection of license opportunities"""
        metric = ExtractedMetric(
            value="929",
            numeric_value=929.0,
            context="929 high-value unlicensed users",
            metric_type="count"
        )

        slide_data = SlideData(
            slide_number=3,
            title="License Priority",
            metrics=[metric],
            text_content="High-value users without licenses",
            key_phrases=["unlicensed", "high-value"]
        )

        gaps = self.detector.detect_gaps(slide_data)

        # Should detect license opportunity
        license_gaps = [g for g in gaps if g.gap_type == 'license']
        self.assertTrue(len(license_gaps) > 0)
        self.assertGreater(license_gaps[0].magnitude, 500)


class TestImplicationMapping(unittest.TestCase):
    """Test business implication mapping"""

    def setUp(self):
        self.mapper = ImplicationMapper()

    def test_adoption_gap_implication(self):
        """Test mapping adoption gap to implication"""
        gap = Gap(
            gap_type='adoption',
            magnitude=96.0,
            description="96% of users have not adopted",
            source_metric=ExtractedMetric("4%", 4.0, "context", "percentage"),
            opportunity_size=96.0
        )

        implication = self.mapper.map_gap_to_implication(gap)

        # Should include awareness/training
        self.assertIn('awareness', implication.root_cause.lower())
        self.assertEqual(implication.urgency, 'immediate')

    def test_habit_gap_implication(self):
        """Test mapping habit gap to implication"""
        gap = Gap(
            gap_type='habit',
            magnitude=87.0,
            description="87% remain infrequent",
            source_metric=ExtractedMetric("87%", 87.0, "context", "percentage"),
            opportunity_size=87.0
        )

        implication = self.mapper.map_gap_to_implication(gap)

        # Should reference workflow integration
        self.assertIn('workflow', implication.root_cause.lower())


class TestInsightComposition(unittest.TestCase):
    """Test insight composition"""

    def setUp(self):
        self.composer = InsightComposer()

    def test_headline_generation(self):
        """Test headline contains number and gap insight"""
        metric = ExtractedMetric("4%", 4.0, "adoption rate", "percentage")
        gap = Gap('adoption', 96.0, "96% of users have not adopted - massive awareness gap",
                  metric, 96.0)

        from lib.analysis.implications import Implication
        impl = Implication(
            gap=gap,
            business_impact="unrealized value",
            root_cause="awareness gap",
            action_required="requires immediate awareness campaign",
            urgency="immediate"
        )

        slide_data = SlideData(1, "Adoption", [metric], "text", [])
        insight = self.composer.compose_slide_insights(slide_data, [impl])

        # Headline should contain the number
        self.assertIn("4%", insight.headline) or self.assertIn("96%", insight.headline)

        # Should not be generic
        self.assertNotIn("there are", insight.headline.lower())

    def test_bullet_generation(self):
        """Test generates 3 bullets following formula"""
        metric = ExtractedMetric("929", 929.0, "unlicensed users", "count")
        gap = Gap('license', 929.0, "929 high-value unlicensed users",
                  metric, 929.0)

        from lib.analysis.implications import Implication
        impl = Implication(
            gap=gap,
            business_impact="revenue opportunity",
            root_cause="license allocation gap",
            action_required="validate license upgrade targeting",
            urgency="moderate"
        )

        slide_data = SlideData(1, "License Priority", [metric], "text", [])
        insight = self.composer.compose_slide_insights(slide_data, [impl])

        # Should have 3 bullets
        self.assertEqual(len(insight.bullet_points), 3)

        # All bullets should contain numbers or action words
        combined = ' '.join(insight.bullet_points).lower()
        self.assertTrue(
            '929' in combined or
            any(word in combined for word in ['action', 'requires', 'reveals'])
        )


if __name__ == '__main__':
    unittest.main()
