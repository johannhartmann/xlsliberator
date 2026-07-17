"""Versioned conformance-corpus metadata, generators, and result accounting.

The corpus intentionally separates a fixture recipe from execution evidence.  A
recipe proves that a case is tracked; only a signed-off result produced by a
pinned runtime may prove that the case passed.
"""

from __future__ import annotations

import hashlib
import json
import random
from collections.abc import Callable, Iterable
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

EvidenceStatus = Literal["passed", "failed", "skipped", "unavailable", "unsupported", "waived"]
FixtureFormat = Literal["xls", "xlsx", "xlsm", "xlsb", "ods", "recipe"]
FixtureOrigin = Literal["generated", "public", "synthetic", "malicious", "regression"]


class ExpectedOutcome(BaseModel):
    """Expected outcome for one scenario without claiming it was executed."""

    model_config = ConfigDict(extra="forbid")

    scenario: str
    expected_status: Literal["passed", "failed", "unsupported"]
    assertions: dict[str, str | int | float | bool | None] = Field(default_factory=dict)


class CorpusFixture(BaseModel):
    """A redistributable fixture or deterministic fixture recipe."""

    model_config = ConfigDict(extra="forbid")

    fixture_id: str = Field(pattern=r"^[a-z0-9][a-z0-9._-]+$")
    path: str
    format: FixtureFormat
    origin: FixtureOrigin
    license: str
    provenance: str
    generator: str | None = None
    materialization: Literal["checked-in", "generated", "target-required"]
    confidential: Literal[False] = False
    blocking: bool = False
    categories: list[str] = Field(min_length=1)
    features: list[str] = Field(min_length=1)
    expected: list[ExpectedOutcome] = Field(min_length=1)
    sha256: str | None = Field(default=None, pattern=r"^[0-9a-f]{64}$")
    fixed_defect: str | None = None

    @model_validator(mode="after")
    def require_regression_link(self) -> CorpusFixture:
        """Every regression must identify the defect that it fixes."""
        if self.origin == "regression" and not self.fixed_defect:
            raise ValueError("regression fixtures require fixed_defect")
        if self.materialization == "checked-in" and self.sha256 is None:
            raise ValueError("checked-in fixtures require sha256")
        return self


class CorpusManifest(BaseModel):
    """Machine-readable index for every non-confidential corpus entry."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    corpus_id: str
    fixtures: list[CorpusFixture]

    @model_validator(mode="after")
    def unique_fixture_ids(self) -> CorpusManifest:
        """Reject ambiguous fixture identities."""
        ids = [fixture.fixture_id for fixture in self.fixtures]
        if len(ids) != len(set(ids)):
            raise ValueError("duplicate corpus fixture id")
        return self

    @classmethod
    def load(cls, path: Path) -> CorpusManifest:
        """Load and validate the complete manifest."""
        return cls.model_validate_json(path.read_text(encoding="utf-8"))

    def verify_files(self, root: Path) -> list[str]:
        """Return integrity errors without treating absent generated files as passes."""
        errors: list[str] = []
        for fixture in self.fixtures:
            path = root / fixture.path
            if fixture.materialization != "checked-in":
                continue
            if not path.is_file():
                errors.append(f"{fixture.fixture_id}: fixture is absent")
                continue
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            if digest != fixture.sha256:
                errors.append(f"{fixture.fixture_id}: sha256 mismatch")
        return errors


class CorpusExecution(BaseModel):
    """One immutable execution disposition."""

    model_config = ConfigDict(extra="forbid")

    fixture_id: str
    scenario: str
    environment: str
    target: str
    target_version: str
    status: EvidenceStatus
    failure_signature: str | None = None
    evidence_path: str | None = None

    @model_validator(mode="after")
    def require_evidence_for_pass_or_failure(self) -> CorpusExecution:
        """A decisive result must have a durable evidence reference."""
        if self.status in {"passed", "failed"} and not self.evidence_path:
            raise ValueError("passed and failed executions require evidence_path")
        if self.status == "failed" and not self.failure_signature:
            raise ValueError("failed executions require failure_signature")
        return self


class CorpusStatistics(BaseModel):
    """Counts preserve every disposition instead of hiding it in a percentage."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    fixture_count: int
    format_counts: dict[str, int]
    origin_counts: dict[str, int]
    status_counts: dict[EvidenceStatus, int]
    unique_failure_signatures: int
    duplicate_failures: int


class CorpusTrendReport(BaseModel):
    """Machine-readable current corpus statistics and deltas from a baseline."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    corpus_id: str
    current: CorpusStatistics
    previous: CorpusStatistics | None = None
    fixture_delta: int
    status_deltas: dict[EvidenceStatus, int]
    unique_failure_signature_delta: int


def normalized_failure_signature(*, gate: str, error_type: str, trace_diff: Iterable[str]) -> str:
    """Hash a stable failure identity while removing volatile paths and numbers."""
    normalized: list[str] = []
    for raw in trace_diff:
        text = raw.replace("\\", "/")
        parts = [part for part in text.split("/") if not part.startswith("tmp")]
        compact = "/".join(parts)
        compact = "".join("#" if character.isdigit() else character for character in compact)
        normalized.append(" ".join(compact.split()))
    payload = json.dumps(
        {"gate": gate, "error_type": error_type, "trace_diff": sorted(normalized)},
        separators=(",", ":"),
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode()).hexdigest()


def corpus_statistics(
    manifest: CorpusManifest, executions: Iterable[CorpusExecution]
) -> CorpusStatistics:
    """Aggregate counts with explicit unavailable/skipped/unsupported buckets."""
    format_counts: dict[str, int] = {}
    origin_counts: dict[str, int] = {}
    for fixture in manifest.fixtures:
        format_counts[fixture.format] = format_counts.get(fixture.format, 0) + 1
        origin_counts[fixture.origin] = origin_counts.get(fixture.origin, 0) + 1
    status_counts: dict[EvidenceStatus, int] = dict.fromkeys(
        ("passed", "failed", "skipped", "unavailable", "unsupported", "waived"), 0
    )
    signatures: list[str] = []
    for execution in executions:
        status_counts[execution.status] += 1
        if execution.failure_signature:
            signatures.append(execution.failure_signature)
    return CorpusStatistics(
        fixture_count=len(manifest.fixtures),
        format_counts=dict(sorted(format_counts.items())),
        origin_counts=dict(sorted(origin_counts.items())),
        status_counts=status_counts,
        unique_failure_signatures=len(set(signatures)),
        duplicate_failures=len(signatures) - len(set(signatures)),
    )


def corpus_trend_report(
    manifest: CorpusManifest,
    executions: Iterable[CorpusExecution],
    *,
    previous: CorpusStatistics | None = None,
) -> CorpusTrendReport:
    """Generate a stable report without collapsing non-decisive dispositions."""
    current = corpus_statistics(manifest, executions)
    previous_status = previous.status_counts if previous else {}
    return CorpusTrendReport(
        corpus_id=manifest.corpus_id,
        current=current,
        previous=previous,
        fixture_delta=current.fixture_count - (previous.fixture_count if previous else 0),
        status_deltas={
            status: count - previous_status.get(status, 0)
            for status, count in current.status_counts.items()
        },
        unique_failure_signature_delta=(
            current.unique_failure_signatures
            - (previous.unique_failure_signatures if previous else 0)
        ),
    )


def load_corpus_executions(path: Path) -> list[CorpusExecution]:
    """Load a versioned list of corpus execution dispositions."""
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0.0":
        raise ValueError("unsupported corpus execution schema")
    return [CorpusExecution.model_validate(item) for item in payload.get("executions", [])]


class FormulaCase(BaseModel):
    """Deterministic formula-fuzz case."""

    model_config = ConfigDict(extra="forbid")

    case_id: str
    formula: str
    equivalent_formula: str
    locale: str
    date_system: Literal["1900", "1904"]
    calculation_mode: Literal["automatic", "manual"]


def generate_formula_cases(seed: int, count: int) -> list[FormulaCase]:
    """Generate stable formula combinations and equivalent rewrites."""
    # Corpus reproducibility requires a deterministic generator, not cryptographic randomness.
    rng = random.Random(seed)  # nosec B311
    cases: list[FormulaCase] = []
    for index in range(count):
        left = rng.randint(-1000, 1000)
        right = rng.randint(-1000, 1000)
        cases.append(
            FormulaCase(
                case_id=f"formula-{seed}-{index}",
                formula=f"=SUM({left};{right})",
                equivalent_formula=f"={left}+{right}",
                locale=rng.choice(["en-US", "de-DE"]),
                date_system=rng.choice(["1900", "1904"]),
                calculation_mode=rng.choice(["automatic", "manual"]),
            )
        )
    return cases


class WorkbookRecipe(BaseModel):
    """Portable deterministic recipe materialized by a format-capable runtime."""

    model_config = ConfigDict(extra="forbid")

    recipe_id: str
    seed: int
    format: FixtureFormat
    sheets: list[str]
    formulas: dict[str, str] = Field(default_factory=dict)
    names: dict[str, str] = Field(default_factory=dict)
    tables: list[str] = Field(default_factory=list)
    styles: dict[str, str] = Field(default_factory=dict)
    validations: dict[str, str] = Field(default_factory=dict)
    charts: list[str] = Field(default_factory=list)
    pivots: list[str] = Field(default_factory=list)
    controls: list[str] = Field(default_factory=list)
    vba_modules: dict[str, str] = Field(default_factory=dict)
    locale: str = "en-US"
    date_system: Literal["1900", "1904"] = "1900"
    calculation_mode: Literal["automatic", "manual"] = "automatic"

    @property
    def canonical_sha256(self) -> str:
        """Hash the normalized recipe rather than a tool-specific ZIP timestamp."""
        payload = self.model_dump_json(exclude_none=True)
        return hashlib.sha256(payload.encode()).hexdigest()


def generate_names_tables_recipe(seed: int, *, format: FixtureFormat = "xlsx") -> WorkbookRecipe:
    """Generate names, tables, structured references, and rename operations."""
    return WorkbookRecipe(
        recipe_id=f"names-tables-{seed}",
        seed=seed,
        format=format,
        sheets=["Input", "Summary"],
        formulas={"Summary.A1": "=SUM(Sales[Amount])"},
        names={"TaxRate": "Input.$B$1", "SalesData": "Input.$A$3:$B$8"},
        tables=["Sales"],
    )


def generate_styles_validations_recipe(
    seed: int, *, format: FixtureFormat = "xlsx"
) -> WorkbookRecipe:
    """Generate stable style and validation combinations."""
    return WorkbookRecipe(
        recipe_id=f"styles-validations-{seed}",
        seed=seed,
        format=format,
        sheets=["Styled"],
        styles={"Styled.A1": "currency-red-negative", "Styled.B1": "date-iso"},
        validations={"Styled.C1:C20": "list:Open,Closed", "Styled.D1": "whole:1:10"},
    )


def generate_charts_pivots_controls_recipe(
    seed: int, *, format: FixtureFormat = "xlsb"
) -> WorkbookRecipe:
    """Generate chart, pivot, form-control, and event-binding surfaces."""
    return WorkbookRecipe(
        recipe_id=f"charts-pivots-controls-{seed}",
        seed=seed,
        format=format,
        sheets=["Data", "Dashboard"],
        charts=["Dashboard.SalesChart"],
        pivots=["Dashboard.SalesPivot"],
        controls=["Dashboard.RefreshButton:onAction=RefreshReport"],
    )


def generate_vba_recipe(seed: int, *, format: FixtureFormat = "xlsm") -> WorkbookRecipe:
    """Generate a deterministic VBA micro-program with an observable result."""
    return WorkbookRecipe(
        recipe_id=f"vba-micro-{seed}",
        seed=seed,
        format=format,
        sheets=["MacroResult"],
        vba_modules={
            "Module1": (
                "Option Explicit\n"
                "Public Sub WriteResult()\n"
                '  Worksheets("MacroResult").Range("A1").Value = 42\n'
                "End Sub\n"
            )
        },
    )


def generate_environment_recipe(
    seed: int,
    *,
    locale: str,
    date_system: Literal["1900", "1904"],
    calculation_mode: Literal["automatic", "manual"],
) -> WorkbookRecipe:
    """Generate locale/date-system/calculation-mode combinations."""
    return WorkbookRecipe(
        recipe_id=f"environment-{seed}-{locale}-{date_system}-{calculation_mode}",
        seed=seed,
        format="xlsx",
        sheets=["Environment"],
        formulas={"Environment.A1": "=DATE(2024;2;29)", "Environment.A2": "=A1+1"},
        locale=locale,
        date_system=date_system,
        calculation_mode=calculation_mode,
    )


class MetamorphicResult(BaseModel):
    """Result of one invariant checked across related executions."""

    model_config = ConfigDict(extra="forbid")

    relation: str
    status: EvidenceStatus
    observations: list[str]


def check_metamorphic_relations(
    case: FormulaCase,
    *,
    execute: Callable[[str], str] | None,
    independent_target: Callable[[str], str] | None,
) -> list[MetamorphicResult]:
    """Check rewrite, repeated recalculation, and independent-target invariants."""
    if execute is None:
        return [
            MetamorphicResult(relation=relation, status="unavailable", observations=[])
            for relation in ("equivalent-rewrite", "repeated-recalculation", "independent-target")
        ]
    original = execute(case.formula)
    rewritten = execute(case.equivalent_formula)
    repeated = execute(case.formula)
    results = [
        MetamorphicResult(
            relation="equivalent-rewrite",
            status="passed" if original == rewritten else "failed",
            observations=[original, rewritten],
        ),
        MetamorphicResult(
            relation="repeated-recalculation",
            status="passed" if original == repeated else "failed",
            observations=[original, repeated],
        ),
    ]
    if independent_target is None:
        results.append(
            MetamorphicResult(
                relation="independent-target", status="unavailable", observations=[original]
            )
        )
    else:
        independent = independent_target(case.formula)
        results.append(
            MetamorphicResult(
                relation="independent-target",
                status="passed" if original == independent else "failed",
                observations=[original, independent],
            )
        )
    return results


def copy_move_rename_recipe(
    recipe: WorkbookRecipe, *, source_sheet: str, target_sheet: str
) -> WorkbookRecipe:
    """Apply a stable sheet rename and rewrite formula/name references."""
    if source_sheet not in recipe.sheets:
        raise ValueError(f"source sheet is absent: {source_sheet}")
    sheets = [target_sheet if name == source_sheet else name for name in recipe.sheets]
    formulas = {
        address.replace(f"{source_sheet}.", f"{target_sheet}."): formula.replace(
            f"{source_sheet}.", f"{target_sheet}."
        )
        for address, formula in recipe.formulas.items()
    }
    names = {
        name: reference.replace(f"{source_sheet}.", f"{target_sheet}.")
        for name, reference in recipe.names.items()
    }
    return recipe.model_copy(update={"sheets": sheets, "formulas": formulas, "names": names})


def recipe_save_reopen_stable(recipe: WorkbookRecipe) -> bool:
    """Prove canonical serialization survives a save/reopen round trip."""
    reopened = WorkbookRecipe.model_validate_json(recipe.model_dump_json())
    return reopened.canonical_sha256 == recipe.canonical_sha256


class DifferentialResult(BaseModel):
    """Truthful result of one source/target differential attempt."""

    model_config = ConfigDict(extra="forbid")

    case_id: str
    status: EvidenceStatus
    source_value: str | None = None
    target_value: str | None = None
    failure_signature: str | None = None


def differential_fuzz(
    cases: Iterable[FormulaCase],
    *,
    source: Callable[[str], str] | None,
    target: Callable[[str], str] | None,
) -> list[DifferentialResult]:
    """Compare runtimes, retaining unavailable as a first-class outcome."""
    results: list[DifferentialResult] = []
    for case in cases:
        if source is None or target is None:
            results.append(DifferentialResult(case_id=case.case_id, status="unavailable"))
            continue
        source_value = source(case.formula)
        target_value = target(case.formula)
        if source_value == target_value:
            results.append(
                DifferentialResult(
                    case_id=case.case_id,
                    status="passed",
                    source_value=source_value,
                    target_value=target_value,
                )
            )
        else:
            signature = normalized_failure_signature(
                gate="source-differential",
                error_type="value-mismatch",
                trace_diff=[source_value, target_value],
            )
            results.append(
                DifferentialResult(
                    case_id=case.case_id,
                    status="failed",
                    source_value=source_value,
                    target_value=target_value,
                    failure_signature=signature,
                )
            )
    return results
