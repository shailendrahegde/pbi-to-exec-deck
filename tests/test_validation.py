"""
Unit tests for Constitution validation
"""

import unittest
from lib.rendering.validator import ConstitutionValidator, ValidationResult
from lib.analysis.insights import Insight


class TestConstitutionValidation(unittest.TestCase):
    """Test Constitution compliance validation"""

    def setUp(self):
        self.validator = ConstitutionValidator()

    def test_reject_generic_headline(self):
        """Test rejection of generic headlines"""
        headline = "There are 1,275 active users"
        results = self.validator._validate_headline("Test", headline)

        # Should fail for generic pattern
        errors = [r for r in results if r.severity == "error"]
        self.assertTrue(len(errors) > 0)

    def test_accept_insight_headline(self):
        """Test acceptance of insight-driven headline"""
        headline = "96% of users have not adopted agents - massive awareness gap requiring intervention"
        results = self.validator._validate_headline("Test", headline)

        # Should pass (has number and not generic)
        errors = [r for r in results if r.severity == "error"]
        self.assertEqual(len(errors), 0)

    def test_reject_headline_without_number(self):
        """Test rejection of headline without specific number"""
        headline = "User adoption needs improvement"
        results = self.validator._validate_headline("Test", headline)

        # Should fail for missing number
        errors = [r for r in results if r.severity == "error"]
        self.assertTrue(len(errors) > 0)

    def test_actionability_check(self):
        """Test actionability validation"""
        # Good bullets with action words
        good_bullets = [
            "96% adoption gap reveals awareness issue",
            "Indicates need for targeted training program",
            "Action: Launch awareness campaign in Q1"
        ]

        results = self.validator._validate_actionability("Test", good_bullets)
        errors = [r for r in results if r.severity == "error"]
        self.assertEqual(len(errors), 0)

        # Bad bullets without action
        bad_bullets = [
            "There are 1,275 users",
            "Total of 324 users increased",
            "Shows 87% infrequent usage"
        ]

        results = self.validator._validate_actionability("Test", bad_bullets)
        # Should have warnings or errors
        issues = [r for r in results if not r.passed]
        self.assertTrue(len(issues) > 0)

    def test_source_number_tracking(self):
        """Test source number validation"""
        insight_with_numbers = Insight(
            headline="96% gap requires intervention",
            bullet_points=["Gap analysis", "Action needed"],
            source_numbers=["4%", "96%", "1,275"]
        )

        results = self.validator._validate_numbers("Test", insight_with_numbers)

        # Should track source numbers
        info_results = [r for r in results if r.severity == "info" and r.passed]
        self.assertTrue(len(info_results) > 0)

    def test_slide_count_validation(self):
        """Test 1:1 slide mapping validation"""
        # Correct mapping: 7 source slides + 1 title = 8 output
        result = self.validator.validate_slide_count(7, 8)
        self.assertTrue(result.passed)

        # Incorrect mapping
        result = self.validator.validate_slide_count(7, 10)
        self.assertFalse(result.passed)


class TestValidationReport(unittest.TestCase):
    """Test validation report generation"""

    def setUp(self):
        self.validator = ConstitutionValidator()

    def test_report_shows_errors(self):
        """Test report highlights errors"""
        results = [
            ValidationResult(False, "Test Rule", "Test error", "error"),
            ValidationResult(True, "Test Rule 2", "Test pass", "info")
        ]

        report = self.validator.generate_report(results)

        self.assertIn("ERRORS", report)
        self.assertIn("Test error", report)

    def test_report_shows_success(self):
        """Test report shows success when all pass"""
        results = [
            ValidationResult(True, "Test Rule 1", "Test pass 1", "info"),
            ValidationResult(True, "Test Rule 2", "Test pass 2", "info")
        ]

        report = self.validator.generate_report(results)

        self.assertIn("ALL CHECKS PASSED", report)


if __name__ == '__main__':
    unittest.main()
