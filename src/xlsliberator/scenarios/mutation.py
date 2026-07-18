"""Mutation testing against public migration acceptance scenarios."""

from __future__ import annotations

import ast
import hashlib
import io
import os
import shutil
import tempfile
import tokenize
import xml.etree.ElementTree as ET  # nosec B405 - serialization only
import zipfile
from pathlib import Path
from typing import Protocol

from defusedxml import ElementTree as DefusedET

from xlsliberator.odstool import (
    CONTENT_PATH,
    OdsToolError,
    inspect_scripts,
    transform_package_member,
)
from xlsliberator.scenarios.acceptance_evidence import write_acceptance_evidence
from xlsliberator.scenarios.assertions import evaluate_trace
from xlsliberator.scenarios.models import (
    AcceptanceDefinition,
    EnvironmentManifest,
    MutationCampaign,
    MutationCaseResult,
    MutationOutcome,
    RuntimeTrace,
    Scenario,
)
from xlsliberator.validation_models import GateExecutionStatus


class MutationTargetRunner(Protocol):
    """Runtime boundary shared by mutation campaigns and the Docker target."""

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> RuntimeTrace:
        """Execute one mutant working copy."""


def run_mutation_campaign(
    *,
    source_workbook: Path,
    acceptance: AcceptanceDefinition,
    directory: Path,
    runner: MutationTargetRunner,
) -> MutationCampaign:
    """Create isolated Python/formula mutants and run public acceptance for each."""
    source = source_workbook.resolve()
    output = directory.resolve()
    if output.exists():
        raise FileExistsError(f"mutation directory already exists: {output}")
    candidates = _mutation_candidates(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.mkdir()

    cases: list[MutationCaseResult] = []
    for index, candidate in enumerate(candidates, start=1):
        kind, target = candidate
        case_id = f"{kind}-{index:03d}"
        case_dir = output / case_id
        case_dir.mkdir()
        mutant = case_dir / "mutant.ods"
        shutil.copy2(source, mutant)
        _apply_mutation(mutant, kind, target)
        trace = runner.run(mutant, acceptance.environment, acceptance.scenario)
        evaluation = evaluate_trace(acceptance, trace)
        evidence_dir = case_dir / "evidence"
        write_acceptance_evidence(
            evidence_dir,
            workbook=mutant,
            acceptance=acceptance,
            trace=trace,
            evaluation=evaluation,
        )
        outcome, reason = _mutation_outcome(trace, evaluation.status)
        cases.append(
            MutationCaseResult(
                id=case_id,
                kind="python" if kind == "python" else "formula",
                target=target,
                mutant_workbook=str(mutant.relative_to(output)),
                mutant_sha256=_hash_file(mutant),
                trace=str((evidence_dir / "trace.json").relative_to(output)),
                evaluation=str((evidence_dir / "evaluation.json").relative_to(output)),
                outcome=outcome,
                reason=reason,
            )
        )

    if cases and all(case.outcome is MutationOutcome.KILLED for case in cases):
        status = GateExecutionStatus.PASSED
    elif any(case.outcome is MutationOutcome.SURVIVED for case in cases):
        status = GateExecutionStatus.FAILED
    else:
        status = GateExecutionStatus.UNAVAILABLE
    campaign = MutationCampaign(
        migration_id=acceptance.migration.id,
        source_workbook_sha256=_hash_file(source),
        status=status,
        cases=cases,
    )
    _atomic_write(output / "mutation-report.json", campaign.model_dump_json(indent=2) + "\n")
    _atomic_write(output / "mutation-report.md", render_mutation_report(campaign))
    return campaign


def render_mutation_report(campaign: MutationCampaign) -> str:
    """Render a deterministic Markdown summary for a mutation campaign."""
    lines = [
        f"# Mutation campaign: {campaign.migration_id}",
        "",
        f"- Result: **{campaign.status.value.upper()}**",
        f"- Source SHA-256: `{campaign.source_workbook_sha256}`",
        f"- Mutants: {len(campaign.cases)}",
        "",
        "| Mutant | Kind | Target | Outcome | Reason |",
        "| --- | --- | --- | --- | --- |",
    ]
    lines.extend(
        f"| {case.id} | {case.kind} | {_markdown(case.target)} | "
        f"{case.outcome.value} | {_markdown(case.reason)} |"
        for case in campaign.cases
    )
    if not campaign.cases:
        lines.extend(
            [
                "",
                "No embedded Python or ODF formulas were available to mutate; "
                "the campaign is UNAVAILABLE, never passed.",
            ]
        )
    return "\n".join(lines) + "\n"


def _mutation_candidates(source: Path) -> list[tuple[str, str]]:
    verification = inspect_scripts(source)
    if not verification.valid:
        raise OdsToolError("; ".join(verification.errors))
    candidates: list[tuple[str, str]] = []
    with zipfile.ZipFile(source) as archive:
        for script in verification.scripts:
            payload = archive.read(script.package_path)
            try:
                _mutate_python(payload)
            except OdsToolError:
                continue
            candidates.append(("python", script.package_path))
            break
        try:
            _mutate_formula_content(archive.read(CONTENT_PATH))
        except OdsToolError:
            pass
        else:
            candidates.append(("formula", CONTENT_PATH))
    return candidates


def _apply_mutation(workbook: Path, kind: str, target: str) -> None:
    transform_package_member(
        workbook,
        target,
        _mutate_python if kind == "python" else _mutate_formula_content,
    )


def _mutate_python(payload: bytes) -> bytes:
    try:
        source = payload.decode("utf-8")
        tokens = list(tokenize.generate_tokens(io.StringIO(source).readline))
    except (UnicodeDecodeError, tokenize.TokenError) as exc:
        raise OdsToolError(f"cannot tokenize embedded Python: {exc}") from exc
    mutated = False
    output: list[tokenize.TokenInfo] = []
    for token in tokens:
        replacement = token.string
        if not mutated and token.type == tokenize.NAME and token.string in {"True", "False"}:
            replacement = "False" if token.string == "True" else "True"
            mutated = True
        elif not mutated and token.type == tokenize.NUMBER:
            replacement = "1" if token.string != "1" else "2"
            mutated = True
        elif not mutated and token.type == tokenize.OP and token.string in {"+", "-", "*"}:
            replacement = {"+": "-", "-": "+", "*": "+"}[token.string]
            mutated = True
        output.append(token._replace(string=replacement))
    if not mutated:
        raise OdsToolError("embedded Python has no supported mutation point")
    rewritten = tokenize.untokenize(output)
    try:
        ast.parse(rewritten)
    except SyntaxError as exc:
        raise OdsToolError(f"Python mutation is not syntactically valid: {exc}") from exc
    return rewritten.encode("utf-8")


def _mutate_formula_content(payload: bytes) -> bytes:
    try:
        root = DefusedET.fromstring(payload)
    except ET.ParseError as exc:
        raise OdsToolError(f"cannot parse formula content: {exc}") from exc
    for element in root.iter():
        for attribute, formula in element.attrib.items():
            if attribute.rsplit("}", 1)[-1] != "formula":
                continue
            mutated = _mutate_formula(formula)
            if mutated is not None:
                element.set(attribute, mutated)
                return bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    raise OdsToolError("ODS contains no supported formula mutation point")


def _mutate_formula(formula: str) -> str | None:
    for original, replacement in (("+", "-"), ("-", "+"), ("*", "+"), ("/", "*")):
        index = formula.find(original)
        if index >= 0:
            return formula[:index] + replacement + formula[index + 1 :]
    for index, character in enumerate(formula):
        if character.isdigit():
            replacement = "1" if character != "1" else "2"
            return formula[:index] + replacement + formula[index + 1 :]
    return None


def _mutation_outcome(
    trace: RuntimeTrace,
    evaluation_status: GateExecutionStatus,
) -> tuple[MutationOutcome, str]:
    if trace.status in {
        GateExecutionStatus.UNAVAILABLE,
        GateExecutionStatus.NOT_RUN,
        GateExecutionStatus.SKIPPED,
    }:
        return MutationOutcome.INCONCLUSIVE, f"runtime was {trace.status.value}"
    if trace.status is GateExecutionStatus.FAILED and not trace.steps:
        return MutationOutcome.INCONCLUSIVE, "runtime failed before acceptance actions ran"
    if evaluation_status is GateExecutionStatus.PASSED:
        return MutationOutcome.SURVIVED, "public acceptance still passed"
    return MutationOutcome.KILLED, "public acceptance detected the mutation"


def _atomic_write(path: Path, payload: str) -> None:
    descriptor, temporary = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    except Exception:
        Path(temporary).unlink(missing_ok=True)
        raise


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _markdown(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ")
