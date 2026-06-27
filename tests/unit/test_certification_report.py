"""Tests for certification reports."""

import json

from xlsliberator.certification_report import (
    CertificationReport,
    certification_from_conversion_report,
)
from xlsliberator.report import ConversionReport
from xlsliberator.validation_models import (
    ValidationCertification,
    ValidationGateResult,
    ValidationSeverity,
)


def test_certification_json_serialization() -> None:
    """Certification report JSON should be parseable."""
    report = CertificationReport(
        ValidationCertification(
            certified=True,
            gate_results=[
                ValidationGateResult(
                    gate_name="inventory",
                    passed=True,
                    message="ok",
                )
            ],
        )
    )

    data = json.loads(report.to_json())

    assert data["certified"] is True
    assert data["gate_results"][0]["gate_name"] == "inventory"


def test_certification_markdown_contains_expected_headings() -> None:
    """Markdown report should contain all required sections."""
    report = CertificationReport(ValidationCertification())
    markdown = report.to_markdown()

    for heading in [
        "## Summary",
        "## Parse Coverage",
        "## Formula Validation",
        "## Macro Validation",
        "## GUI/Control Validation",
        "## Runtime Targets",
        "## Unsupported Artifacts",
        "## Waivers",
        "## Repair History",
        "## Errors and Warnings",
    ]:
        assert heading in markdown


def test_failed_gate_keeps_certified_false() -> None:
    """Failed gates should be represented in non-certified reports."""
    report = CertificationReport(
        ValidationCertification(
            certified=False,
            gate_results=[
                ValidationGateResult(
                    gate_name="formula",
                    passed=False,
                    severity=ValidationSeverity.ERROR,
                    message="parse failed",
                )
            ],
        )
    )

    assert not report.certification.certified
    assert "formula: failed" in report.to_markdown()


def test_bridge_from_conversion_report() -> None:
    """Legacy conversion reports should bridge to certification reports."""
    conversion = ConversionReport(input_file="in.xlsx", output_file="out.ods", success=True)
    report = certification_from_conversion_report(conversion)

    assert report.certification.certified
    assert report.certification.gate_results[0].gate_name == "conversion"


def test_conversion_gate_rendered_in_markdown() -> None:
    """The legacy 'conversion' gate must be visible in the Markdown report."""
    conversion = ConversionReport(input_file="in.xlsx", output_file="out.ods", success=False)
    markdown = certification_from_conversion_report(conversion).to_markdown()

    assert "## Other Gates" in markdown
    assert "conversion: failed" in markdown
