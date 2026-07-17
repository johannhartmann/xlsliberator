"""Windows VBA micro-conformance trace generation tests."""

from pathlib import Path

from xlsliberator.excel_oracle import FakeExcelOracle, UnavailableExcelOracle
from xlsliberator.scenarios.models import EnvironmentManifest
from xlsliberator.validation_models import GateExecutionStatus
from xlsliberator.vba_conformance import (
    default_vba_micro_programs,
    generate_windows_micro_traces,
)


def test_default_micro_corpus_covers_required_semantics() -> None:
    programs = default_vba_micro_programs()

    assert {program.feature for program in programs} == {
        "type-coercion",
        "byref",
        "default-properties",
        "arrays",
        "error-handling",
        "classes",
        "events",
        "range-operations",
    }
    assert all(program.scenario.steps for program in programs)


def test_missing_windows_fixtures_are_explicitly_unavailable(tmp_path: Path) -> None:
    results = generate_windows_micro_traces(
        UnavailableExcelOracle(),
        tmp_path / "fixtures",
        tmp_path / "traces",
        EnvironmentManifest(),
    )

    assert all(result.status is GateExecutionStatus.UNAVAILABLE for result in results)
    assert all(result.error and result.error["type"] == "missing_fixture" for result in results)


def test_fake_oracle_trace_is_never_published_as_real_excel_evidence(tmp_path: Path) -> None:
    program = default_vba_micro_programs()[0]
    fixtures = tmp_path / "fixtures"
    fixtures.mkdir()
    (fixtures / program.workbook_fixture).write_bytes(b"fake workbook")

    result = generate_windows_micro_traces(
        FakeExcelOracle(),
        fixtures,
        tmp_path / "traces",
        EnvironmentManifest(),
        [program],
    )[0]

    assert result.status is GateExecutionStatus.UNAVAILABLE
    assert not result.real_excel_trace
    assert result.trace_path is None
    assert result.error == {"type": "non_excel_trace", "message": "fake_excel_oracle"}
