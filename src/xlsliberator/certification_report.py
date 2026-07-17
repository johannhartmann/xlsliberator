"""Certification reporting for validated transformations."""

import json
import zipfile
from dataclasses import dataclass
from pathlib import Path

from xlsliberator.report import ConversionReport
from xlsliberator.validation_models import (
    GateExecutionStatus,
    ValidationCertification,
    ValidationGateResult,
    ValidationSeverity,
)


@dataclass
class CertificationReport:
    """Wrapper for validation certification output."""

    certification: ValidationCertification

    def to_json(self) -> str:
        """Serialize the certification report as JSON."""
        return self.certification.model_dump_json(indent=2)

    def to_markdown(self) -> str:
        """Serialize the certification report as Markdown."""
        status = "CERTIFIED" if self.certification.certified else "NOT CERTIFIED"
        lines = [
            "# Validation Certification Report",
            "",
            "## Summary",
            f"- Status: {status}",
            f"- Target profiles: {', '.join(self.certification.target_profiles) or 'None'}",
            "",
            "## Parse Coverage",
            self._gate_section("inventory"),
            "",
            "## Formula Validation",
            self._gate_section("formula"),
            "",
            "## Artifact Loss Accounting",
            self._gate_section("coverage"),
            "",
            "## Macro Validation",
            self._gate_section("macro"),
            "",
            "## GUI/Control Validation",
            self._gate_section("control"),
            "",
            "## Runtime Targets",
            self._gate_section("backend"),
            "",
            "## Unsupported Artifacts",
        ]
        if self.certification.unsupported_artifacts:
            for artifact in self.certification.unsupported_artifacts:
                lines.append(f"- {artifact.source_ref.artifact_id}: {artifact.reason}")
        else:
            lines.append("- None")

        lines.extend(
            [
                "",
                "## Waivers",
                "- None",
                "",
                "## Repair History",
                self._gate_section("repair"),
                "",
            ]
        )
        if self.certification.repair_provenance:
            lines.extend(
                f"- Run {item.agent_run_id}; patch {item.candidate_patch_id}; "
                f"deterministic gates: {', '.join(item.deterministic_gate_names)}"
                for item in self.certification.repair_provenance
            )
            lines.append("")

        # Render any gate not covered by a fixed section above (e.g. the legacy
        # "conversion" gate from certification_from_conversion_report) so no gate
        # result is silently dropped from the report.
        known_prefixes = (
            "inventory",
            "coverage",
            "formula",
            "macro",
            "control",
            "backend",
            "repair",
        )
        other_gates = [
            gate
            for gate in self.certification.gate_results
            if not any(
                gate.gate_name == prefix or gate.gate_name.startswith(f"{prefix}_")
                for prefix in known_prefixes
            )
        ]
        if other_gates:
            lines.append("## Other Gates")
            lines.extend(self._format_gate(gate) for gate in other_gates)
            lines.append("")

        lines.append("## Errors and Warnings")
        if self.certification.errors:
            lines.append("### Errors")
            lines.extend(f"- {error}" for error in self.certification.errors)
        if self.certification.warnings:
            lines.append("### Warnings")
            lines.extend(f"- {warning}" for warning in self.certification.warnings)
        if not self.certification.errors and not self.certification.warnings:
            lines.append("- None")

        return "\n".join(lines) + "\n"

    def save_json(self, path: str | Path) -> None:
        """Write JSON report to disk."""
        Path(path).write_text(self.to_json())

    def save_markdown(self, path: str | Path) -> None:
        """Write Markdown report to disk."""
        Path(path).write_text(self.to_markdown())

    def _gate_section(self, gate_prefix: str) -> str:
        matches = [
            gate
            for gate in self.certification.gate_results
            if gate.gate_name == gate_prefix or gate.gate_name.startswith(f"{gate_prefix}_")
        ]
        if not matches:
            return "- Not run"
        return "\n".join(self._format_gate(gate) for gate in matches)

    @staticmethod
    def _format_gate(gate: ValidationGateResult) -> str:
        references = (
            f"; evidence: {', '.join(gate.evidence_references)}" if gate.evidence_references else ""
        )
        return (
            f"- {gate.gate_name}: {gate.status.value} "
            f"({'required' if gate.required else 'optional'}) - {gate.message}{references}"
        )


def certification_from_conversion_report(report: ConversionReport) -> CertificationReport:
    """Build a minimal certification report from a legacy conversion report."""
    output_path = Path(report.output_file)
    valid_output = output_path.is_file() and zipfile.is_zipfile(output_path)
    passed = report.success and not report.errors and valid_output
    gate_results = [
        ValidationGateResult(
            gate_name="conversion",
            status=(GateExecutionStatus.PASSED if passed else GateExecutionStatus.FAILED),
            severity=ValidationSeverity.INFO if passed else ValidationSeverity.ERROR,
            message="Legacy conversion completed" if passed else "Legacy conversion failed",
            details={
                **json.loads(report.to_json()),
                "output_exists": output_path.is_file(),
                "valid_zip": valid_output,
            },
        )
    ]
    certification = ValidationCertification(
        certified=passed,
        gate_results=gate_results,
        warnings=list(report.warnings),
        errors=list(report.errors),
        metadata={"input_file": report.input_file, "output_file": report.output_file},
    )
    return CertificationReport(certification=certification)
