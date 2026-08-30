"""High-level validated transformation API."""

import json
from collections.abc import Mapping
from pathlib import Path

from xlsliberator.api import convert
from xlsliberator.certification_report import CertificationReport
from xlsliberator.libreoffice_scenario_runner import LibreOfficeScenarioRunner
from xlsliberator.scenarios.models import EnvironmentManifest, RuntimeTrace, Scenario
from xlsliberator.validation_runner import (
    ValidationPlan,
    ValidationRunner,
    parse_target_kind,
    required_validation_gates,
)


class ValidatedTransformationError(Exception):
    """Raised when strict validated transformation fails."""

    def __init__(self, message: str, report: CertificationReport) -> None:
        super().__init__(message)
        self.report = report


def transform_validated(
    input_path: str | Path,
    output_path: str | Path,
    targets: list[str] | None = None,
    strict: bool = True,
    max_repair_iterations: int = 0,
    embed_macros: bool = False,
    python_modules: Mapping[str, str] | None = None,
    scenario: Scenario | None = None,
    source_trace: RuntimeTrace | None = None,
    target_trace: RuntimeTrace | None = None,
) -> CertificationReport:
    """Convert a workbook and run validation gates."""
    input_file = Path(input_path)
    output_file = Path(output_path)

    conversion_report = convert(
        input_file,
        output_file,
        strict=False,
        embed_macros=embed_macros,
        python_modules=python_modules,
    )

    target_kinds = []
    for target in targets or ["libreoffice"]:
        target_kinds.extend(parse_target_kind(target))

    if target_trace is None and scenario is not None:
        target_trace = LibreOfficeScenarioRunner().run(
            output_file,
            source_trace.environment if source_trace is not None else EnvironmentManifest(),
            scenario,
        )

    report = ValidationRunner(
        ValidationPlan(
            input_path=input_file,
            output_path=output_file,
            target_kinds=target_kinds,
            strict=strict,
            repair=max_repair_iterations > 0,
            max_repair_iterations=max_repair_iterations,
            conversion_report=conversion_report,
            scenario=scenario,
            source_trace=source_trace,
            target_trace=target_trace,
            enabled_gates=required_validation_gates(
                include_conversion=True,
                include_target_scenario=target_trace is not None,
            ),
        )
    ).run_all()
    report.certification.metadata["conversion_report"] = json.loads(conversion_report.to_json())

    if strict and not report.certification.certified:
        raise ValidatedTransformationError("Validated transformation failed", report)

    return report
