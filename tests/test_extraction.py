"""
Unit tests for data extraction layer
"""

import unittest
from lib.extraction.extractor import DashboardExtractor, ExtractedMetric


class TestMetricExtraction(unittest.TestCase):
    """Test metric extraction patterns"""

    def setUp(self):
        self.extractor = DashboardExtractor()

    def test_percentage_extraction(self):
        """Test percentage pattern matching"""
        text = "4% of users have adopted the feature, while 87% remain infrequent"
        metrics = self.extractor._extract_metrics(text)

        percentages = [m for m in metrics if m.metric_type == 'percentage']
        self.assertEqual(len(percentages), 2)
        self.assertIn(4.0, [m.numeric_value for m in percentages])
        self.assertIn(87.0, [m.numeric_value for m in percentages])

    def test_large_number_suffix(self):
        """Test K/M/B suffix extraction"""
        text = "1.2K active users and 3.5M total interactions"
        metrics = self.extractor._extract_metrics(text)

        # Should extract 1.2K as 1200 and 3.5M as 3500000
        values = [m.numeric_value for m in metrics]
        self.assertIn(1200.0, values)
        self.assertIn(3500000.0, values)

    def test_comma_number_extraction(self):
        """Test comma-formatted number extraction"""
        text = "There are 1,275 active users and 929 unlicensed high-value users"
        metrics = self.extractor._extract_metrics(text)

        values = [m.numeric_value for m in metrics]
        self.assertIn(1275.0, values)
        self.assertIn(929.0, values)

    def test_context_extraction(self):
        """Test that context is captured around numbers"""
        text = "The analysis shows 324 users increased by 45% over the last month"
        metrics = self.extractor._extract_metrics(text)

        # Check that context contains relevant keywords
        for metric in metrics:
            self.assertTrue(len(metric.context) > 0)
            self.assertIn("users" if "324" in str(metric.value) else "analysis", metric.context.lower())


class TestSlideDataExtraction(unittest.TestCase):
    """Test full slide data extraction"""

    def setUp(self):
        self.extractor = DashboardExtractor()

    def test_key_phrase_extraction(self):
        """Test extraction of key phrases"""
        text = """
        Active users show strong adoption trends.
        Frequency tier distribution indicates habit formation gaps.
        License opportunities exist for high-value segments.
        """

        phrases = self.extractor._extract_key_phrases(text)

        # Should find terms like adoption, frequency, license
        phrase_text = ' '.join(phrases).lower()
        self.assertTrue(any(term in phrase_text for term in ['adoption', 'frequency', 'license']))


if __name__ == '__main__':
    unittest.main()
