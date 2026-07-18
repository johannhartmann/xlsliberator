"""Public migration acceptance, evidence, diff, mutation, and report CLI."""

from __future__ import annotations

import json
import os
import re
import tempfile
from pathlib import Path
from typing import Any

import click
import yaml

from xlsliberator.container_boundary import require_application_container
from xlsliberator.libreoffice_session_scenario_runner import (
    LibreOfficeSessionScenarioRunner,
)
from xlsliberator.scenarios.acceptance_evidence import (
    inspect_acceptance_evidence,
    write_acceptance_evidence,
)
from xlsliberator.scenarios.assertions import evaluate_trace
from xlsliberator.scenarios.diff import diff_traces
from xlsliberator.scenarios.models import (
    AcceptanceDefinition,
    AcceptanceEvidenceManifest,
    MutationCampaign,
    RuntimeTrace,
)
from xlsliberator.scenarios.mutation import MutationTargetRunner, run_mutation_campaign
from xlsliberator.validation_models import GateExecutionStatus


@click.group()
def cli() -> None:
    """Execute independently reviewed ODS migration acceptance scenarios."""
    require_application_container()


@cli.command("run")
@click.argument(
    "acceptance_file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--output", type=click.Path(file_okay=False, path_type=Path))
@click.option("--timeout", type=click.IntRange(min=1), default=120, show_default=True)
def run_command(
    acceptance_file: Path,
    ods_file: Path,
    output: Path | None,
    timeout: int,
) -> None:
    """Run ACCEPTANCE_FILE against one ODS_FILE in the pinned Docker runtime."""
    acceptance = load_acceptance(acceptance_file)
    destination = output or (
        ods_file.parent / f"{ods_file.stem}-{_safe_slug(acceptance.migration.id)}-evidence"
    )
    manifest = run_acceptance(
        acceptance=acceptance,
        workbook=ods_file,
        evidence_dir=destination,
        runner=LibreOfficeSessionScenarioRunner(timeout_seconds=timeout),
    )
    click.echo(manifest.model_dump_json(indent=2))
    if manifest.status is not GateExecutionStatus.PASSED:
        raise click.exceptions.Exit(1)


@cli.command("inspect")
@click.argument(
    "evidence_dir",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
def inspect_command(evidence_dir: Path) -> None:
    """Verify an acceptance evidence directory without executing LibreOffice."""
    manifest = inspect_acceptance_evidence(evidence_dir)
    click.echo(manifest.model_dump_json(indent=2))
    if manifest.status is not GateExecutionStatus.PASSED:
        raise click.exceptions.Exit(1)


@cli.command("diff")
@click.argument("trace_a", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("trace_b", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def diff_command(trace_a: Path, trace_b: Path) -> None:
    """Compare two typed traces with exact default comparison rules."""
    source = RuntimeTrace.model_validate_json(trace_a.read_text(encoding="utf-8"))
    target = RuntimeTrace.model_validate_json(trace_b.read_text(encoding="utf-8"))
    difference = diff_traces(source, target)
    click.echo(difference.model_dump_json(indent=2))
    if difference.status is not GateExecutionStatus.PASSED:
        raise click.exceptions.Exit(1)


@cli.command("mutate")
@click.argument(
    "migration_dir",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
@click.option("--timeout", type=click.IntRange(min=1), default=120, show_default=True)
def mutate_command(migration_dir: Path, timeout: int) -> None:
    """Mutate copied generated Python/formulas and run public acceptance."""
    acceptance_path = find_acceptance_file(migration_dir)
    acceptance = load_acceptance(acceptance_path)
    workbook = find_target_workbook(migration_dir, acceptance)
    campaign = run_mutation_campaign(
        source_workbook=workbook,
        acceptance=acceptance,
        directory=migration_dir / "mutations",
        runner=LibreOfficeSessionScenarioRunner(timeout_seconds=timeout),
    )
    click.echo(campaign.model_dump_json(indent=2))
    if campaign.status is not GateExecutionStatus.PASSED:
        raise click.exceptions.Exit(1)


@cli.command("report")
@click.argument(
    "migration_dir",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
def report_command(migration_dir: Path) -> None:
    """Aggregate verified acceptance and mutation evidence."""
    report = build_migration_report(migration_dir)
    _atomic_write(
        migration_dir / "migration-report.json",
        json.dumps(report, indent=2, sort_keys=True) + "\n",
    )
    _atomic_write(migration_dir / "migration-report.md", _render_migration_report(report))
    click.echo(json.dumps(report, indent=2, sort_keys=True))
    if report["status"] != GateExecutionStatus.PASSED.value:
        raise click.exceptions.Exit(1)


def load_acceptance(path: Path) -> AcceptanceDefinition:
    """Load one strict versioned YAML or JSON acceptance definition."""
    payload = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("acceptance definition must be a YAML/JSON object")
    return AcceptanceDefinition.model_validate(payload)


def run_acceptance(
    *,
    acceptance: AcceptanceDefinition,
    workbook: Path,
    evidence_dir: Path,
    runner: MutationTargetRunner,
) -> AcceptanceEvidenceManifest:
    """Execute and persist one fail-closed public acceptance run."""
    target = workbook.resolve()
    if target.suffix.lower() != ".ods":
        raise ValueError("migration-check accepts only LibreOffice .ods targets")
    trace = runner.run(target, acceptance.environment, acceptance.scenario)
    evaluation = evaluate_trace(acceptance, trace)
    return write_acceptance_evidence(
        evidence_dir,
        workbook=target,
        acceptance=acceptance,
        trace=trace,
        evaluation=evaluation,
    )


def find_acceptance_file(migration_dir: Path) -> Path:
    """Resolve exactly one public acceptance file from a migration directory."""
    root = migration_dir.resolve()
    preferred = [
        "public-acceptance.yaml",
        "acceptance.yaml",
        "public-acceptance.yml",
        "acceptance.yml",
        "public-acceptance.json",
        "acceptance.json",
    ]
    matches = [root / name for name in preferred if (root / name).is_file()]
    if not matches:
        raise FileNotFoundError("migration directory has no public acceptance YAML/JSON")
    if len(matches) > 1:
        raise ValueError(f"migration directory has ambiguous acceptance files: {matches}")
    return matches[0]


def find_target_workbook(
    migration_dir: Path,
    acceptance: AcceptanceDefinition,
) -> Path:
    """Resolve a declared or unambiguous ODS target without leaving the migration root."""
    root = migration_dir.resolve()
    declared = acceptance.migration.target_workbook
    if declared:
        target = (root / declared).resolve()
        if not target.is_relative_to(root):
            raise ValueError("declared target workbook escapes the migration directory")
        if not target.is_file() or target.suffix.lower() != ".ods":
            raise FileNotFoundError(f"declared ODS target does not exist: {declared}")
        return target
    matches = sorted(path for path in root.glob("*.ods") if path.is_file())
    if len(matches) != 1:
        raise ValueError("migration directory must contain exactly one ODS target")
    return matches[0]


def build_migration_report(migration_dir: Path) -> dict[str, Any]:
    """Aggregate only verified evidence and typed mutation reports."""
    root = migration_dir.resolve()
    evidence: list[dict[str, str]] = []
    for path in sorted(root.rglob("manifest.json")):
        if any(
            (ancestor / "mutation-report.json").is_file()
            for ancestor in path.parents
            if ancestor.is_relative_to(root)
        ):
            continue
        raw = json.loads(path.read_text(encoding="utf-8"))
        if "evidence_id" not in raw:
            continue
        manifest = inspect_acceptance_evidence(path.parent)
        evidence.append(
            {
                "path": str(path.parent.relative_to(root)),
                "migration_id": manifest.migration_id,
                "status": manifest.status.value,
                "trace": manifest.execution_trace,
            }
        )
    campaigns = [
        MutationCampaign.model_validate_json(path.read_text(encoding="utf-8"))
        for path in sorted(root.rglob("mutation-report.json"))
    ]
    statuses = [
        *(item["status"] for item in evidence),
        *(campaign.status.value for campaign in campaigns),
    ]
    status = (
        GateExecutionStatus.PASSED
        if evidence and all(item == GateExecutionStatus.PASSED.value for item in statuses)
        else GateExecutionStatus.FAILED
    )
    return {
        "schema_version": "1.0.0",
        "status": status.value,
        "evidence": evidence,
        "mutation_campaigns": [
            {
                "migration_id": campaign.migration_id,
                "status": campaign.status.value,
                "mutants": len(campaign.cases),
                "killed": sum(case.outcome.value == "killed" for case in campaign.cases),
            }
            for campaign in campaigns
        ],
    }


def _render_migration_report(report: dict[str, Any]) -> str:
    lines = [
        "# Migration evidence report",
        "",
        f"- Result: **{str(report['status']).upper()}**",
        f"- Acceptance runs: {len(report['evidence'])}",
        f"- Mutation campaigns: {len(report['mutation_campaigns'])}",
        "",
        "## Acceptance runs",
        "",
        "| Path | Migration | Status |",
        "| --- | --- | --- |",
    ]
    lines.extend(
        f"| {_markdown(item['path'])} | {_markdown(item['migration_id'])} | {item['status']} |"
        for item in report["evidence"]
    )
    lines.extend(
        [
            "",
            "## Mutation campaigns",
            "",
            "| Migration | Status | Mutants | Killed |",
            "| --- | --- | ---: | ---: |",
        ]
    )
    lines.extend(
        f"| {_markdown(item['migration_id'])} | {item['status']} | "
        f"{item['mutants']} | {item['killed']} |"
        for item in report["mutation_campaigns"]
    )
    return "\n".join(lines) + "\n"


def _safe_slug(value: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", value).strip("-")
    return slug or "migration"


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


def _markdown(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ")


if __name__ == "__main__":
    cli()
