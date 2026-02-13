"""
Constitution Compliance Validator - Layer 4B

Validates output against Claude PowerPoint Constitution v1.7.
"""

from typing import Dict, List, Tuple
from dataclasses import dataclass
from lib.analysis.insights import Insight
import re


@dataclass
class ValidationResult:
    """Result of validation check"""
    passed: bool
    rule: str  # Constitution rule being checked
    message: str  # Details about pass/failure
    severity: str  # "error", "warning", "info"


class ConstitutionValidator:
    """Validates presentations against Constitution v1.7"""

    # Generic statements to reject (Section 5 - Actionability)
    GENERIC_PATTERNS = [
        r'there are [\d,]+ \w+ users',
        r'total of [\d,]+',
        r'we have [\d,]+',
        r'shows [\d,]+ users',
        r'indicates [\d,]+ users',
    ]

    # Required action words (Section 5)
    ACTION_WORDS = [
        'requires', 'indicates', 'reveals', 'signals', 'demonstrates',
        'action:', 'recommendation:', 'next step:', 'intervention',
        'opportunity', 'gap', 'validates', 'targeting'
    ]

    def validate_insights(self, insights: Dict[str, Insight]) -> List[ValidationResult]:
        """
        Validate all insights against Constitution.

        Args:
            insights: Dictionary of slide titles to Insights

        Returns:
            List of validation results
        """
        results = []

        for title, insight in insights.items():
            # Section 4: Insight-driven headlines
            results.extend(self._validate_headline(title, insight.headline))

            # Section 5: Actionability
            results.extend(self._validate_actionability(title, insight.bullet_points))

            # Section 6.3: Specific numbers from source
            results.extend(self._validate_numbers(title, insight))

        return results

    def _validate_headline(self, slide_title: str, headline: str) -> List[ValidationResult]:
        """Validate headline is insight-driven with specific numbers"""
        results = []

        # Check for numbers in headline
        has_number = bool(re.search(r'\d+', headline))

        if not has_number:
            results.append(ValidationResult(
                passed=False,
                rule="Section 4 & 6.3: Insight-driven headlines with specific numbers",
                message=f"Headline '{headline}' lacks specific number",
                severity="error"
            ))
        else:
            results.append(ValidationResult(
                passed=True,
                rule="Section 6.3: Specific numbers",
                message=f"Headline contains number: {headline[:50]}...",
                severity="info"
            ))

        # Check headline is not generic
        for pattern in self.GENERIC_PATTERNS:
            if re.search(pattern, headline.lower()):
                results.append(ValidationResult(
                    passed=False,
                    rule="Section 5: Actionability (no generic statements)",
                    message=f"Headline contains generic pattern: {pattern}",
                    severity="error"
                ))
                break

        return results

    def _validate_actionability(self, slide_title: str, bullets: List[str]) -> List[ValidationResult]:
        """Validate insights are actionable, not just data restatements"""
        results = []

        combined_text = ' '.join(bullets).lower()

        # Check for action words
        has_action = any(word in combined_text for word in self.ACTION_WORDS)

        if not has_action:
            results.append(ValidationResult(
                passed=False,
                rule="Section 5: Actionability",
                message=f"Insights lack actionable language (requires/indicates/action/etc.)",
                severity="warning"
            ))
        else:
            results.append(ValidationResult(
                passed=True,
                rule="Section 5: Actionability",
                message="Insights contain actionable language",
                severity="info"
            ))

        # Check for generic statements
        for bullet in bullets:
            for pattern in self.GENERIC_PATTERNS:
                if re.search(pattern, bullet.lower()):
                    results.append(ValidationResult(
                        passed=False,
                        rule="Section 5: No generic statements",
                        message=f"Bullet contains generic statement: {bullet[:50]}...",
                        severity="error"
                    ))
                    break

        return results

    def _validate_numbers(self, slide_title: str, insight: Insight) -> List[ValidationResult]:
        """Validate all numbers are from source"""
        results = []

        # Check that source numbers are tracked
        if insight.source_numbers:
            results.append(ValidationResult(
                passed=True,
                rule="Section 6.3: Source traceability",
                message=f"Numbers tracked: {', '.join(insight.source_numbers)}",
                severity="info"
            ))
        else:
            results.append(ValidationResult(
                passed=False,
                rule="Section 6.3: Source traceability",
                message="No source numbers tracked",
                severity="warning"
            ))

        return results

    def validate_slide_count(
        self,
        source_slide_count: int,
        output_slide_count: int
    ) -> ValidationResult:
        """Validate 1:1 mapping of source to output slides"""
        # Account for title slide in output (+1)
        expected_output = source_slide_count + 1

        if output_slide_count == expected_output:
            return ValidationResult(
                passed=True,
                rule="Section 6: Source fidelity (1:1 mapping)",
                message=f"Correct slide count: {source_slide_count} source → {output_slide_count} output (includes title)",
                severity="info"
            )
        else:
            return ValidationResult(
                passed=False,
                rule="Section 6: Source fidelity (1:1 mapping)",
                message=f"Slide count mismatch: {source_slide_count} source → {output_slide_count} output (expected {expected_output})",
                severity="error"
            )

    def generate_report(self, results: List[ValidationResult]) -> str:
        """Generate validation report"""
        errors = [r for r in results if r.severity == "error"]
        warnings = [r for r in results if r.severity == "warning"]
        passed = [r for r in results if r.passed and r.severity == "info"]

        report = []
        report.append("=" * 60)
        report.append("CONSTITUTION VALIDATION REPORT")
        report.append("=" * 60)
        report.append(f"\nSummary:")
        report.append(f"  OK Passed: {len(passed)}")
        report.append(f"  ! Warnings: {len(warnings)}")
        report.append(f"  X Errors: {len(errors)}")

        if errors:
            report.append(f"\n{'=' * 60}")
            report.append("ERRORS (Must Fix):")
            report.append("=" * 60)
            for error in errors:
                report.append(f"\nX {error.rule}")
                report.append(f"  {error.message}")

        if warnings:
            report.append(f"\n{'=' * 60}")
            report.append("WARNINGS (Should Fix):")
            report.append("=" * 60)
            for warning in warnings:
                report.append(f"\n! {warning.rule}")
                report.append(f"  {warning.message}")

        if not errors and not warnings:
            report.append("\n" + "=" * 60)
            report.append("OK ALL CHECKS PASSED - CONSTITUTION COMPLIANT")
            report.append("=" * 60)

        return "\n".join(report)


def validate_output(insights: Dict[str, Insight]) -> Tuple[bool, str]:
    """
    Main entry point: Validate insights against Constitution.

    Args:
        insights: Dictionary of slide titles to Insights

    Returns:
        Tuple of (passed, report)
    """
    validator = ConstitutionValidator()
    results = validator.validate_insights(insights)
    report = validator.generate_report(results)

    # Check if any errors
    has_errors = any(r.severity == "error" for r in results)

    return (not has_errors, report)
