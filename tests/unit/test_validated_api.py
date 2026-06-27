"""Tests for high-level validated API."""

from pathlib import Path
from typing import Any

import pytest

from xlsliberator.certification_report import CertificationReport
from xlsliberator.validated_api import ValidatedTransformationError, transform_validated
from xlsliberator.validation_models import ValidationCertification


class _PassingRunner:
    def __init__(self, _plan: object) -> None:
        pass

    def run_all(self) -> CertificationReport:
        return CertificationReport(ValidationCertification(certified=True))


class _FailingRunner:
    def __init__(self, _plan: object) -> None:
        pass

    def run_all(self) -> CertificationReport:
        return CertificationReport(ValidationCertification(certified=False, errors=["failed gate"]))


def test_transform_validated_returns_report(tmp_path: Path, monkeypatch: Any) -> None:
    """Validated transform should compose convert and ValidationRunner."""
    calls: list[tuple[Path, Path]] = []

    import xlsliberator.validated_api as validated_module

    monkeypatch.setattr(
        validated_module,
        "convert",
        lambda input_path, output_path, **_kwargs: calls.append((input_path, output_path)),
    )
    monkeypatch.setattr(validated_module, "ValidationRunner", _PassingRunner)

    report = transform_validated(tmp_path / "in.xlsx", tmp_path / "out.ods")

    assert report.certification.certified
    assert calls == [(tmp_path / "in.xlsx", tmp_path / "out.ods")]


def test_transform_validated_strict_failure_raises(tmp_path: Path, monkeypatch: Any) -> None:
    """Strict validated transform should raise with evidence report."""
    import xlsliberator.validated_api as validated_module

    monkeypatch.setattr(validated_module, "convert", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(validated_module, "ValidationRunner", _FailingRunner)

    with pytest.raises(ValidatedTransformationError) as exc:
        transform_validated(tmp_path / "in.xlsx", tmp_path / "out.ods", strict=True)

    assert not exc.value.report.certification.certified


def test_transform_validated_non_strict_returns_failed_report(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Non-strict validated transform should return failed certification evidence."""
    import xlsliberator.validated_api as validated_module

    monkeypatch.setattr(validated_module, "convert", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(validated_module, "ValidationRunner", _FailingRunner)

    report = transform_validated(tmp_path / "in.xlsx", tmp_path / "out.ods", strict=False)

    assert not report.certification.certified
