"""High-level validated transformation API."""

from pathlib import Path

from xlsliberator.api import convert
from xlsliberator.certification_report import CertificationReport
from xlsliberator.validation_runner import ValidationPlan, ValidationRunner, parse_target_kind


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
    embed_macros: bool = True,
    use_agent: bool = True,
) -> CertificationReport:
    """Convert a workbook and run validation gates."""
    input_file = Path(input_path)
    output_file = Path(output_path)

    convert(
        input_file,
        output_file,
        strict=False,
        embed_macros=embed_macros,
        use_agent=use_agent,
    )

    target_kinds = []
    for target in targets or ["both"]:
        target_kinds.extend(parse_target_kind(target))

    report = ValidationRunner(
        ValidationPlan(
            input_path=input_file,
            output_path=output_file,
            target_kinds=target_kinds,
            strict=strict,
            repair=max_repair_iterations > 0,
            max_repair_iterations=max_repair_iterations,
        )
    ).run_all()

    if strict and not report.certification.certified:
        raise ValidatedTransformationError("Validated transformation failed", report)

    return report
