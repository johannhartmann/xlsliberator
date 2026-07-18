"""CLI interface for xlsliberator (Phase F12)."""

import json
import sys
from datetime import timedelta
from pathlib import Path
from typing import Literal

import click

from xlsliberator.api import convert
from xlsliberator.container_boundary import ContainerBoundaryError, require_application_container
from xlsliberator.scenarios.models import RuntimeTrace, Scenario


@click.group()
@click.version_option(version="0.1.0")
def cli() -> None:
    """XLSLiberator - Excel to LibreOffice Calc converter."""
    try:
        require_application_container()
    except ContainerBoundaryError as exc:
        raise click.ClickException(str(exc)) from exc


def _load_formula_evidence(
    scenario_file: Path | None,
    source_trace_file: Path | None,
    target_trace_file: Path | None,
) -> tuple[Scenario | None, RuntimeTrace | None, RuntimeTrace | None]:
    """Load an exact scenario and its typed runtime traces for certification."""
    supplied = (scenario_file, source_trace_file, target_trace_file)
    if not any(supplied):
        return None, None, None
    if scenario_file is None or source_trace_file is None:
        raise click.ClickException(
            "--scenario-file and --source-trace-file must be supplied together"
        )
    scenario = Scenario.model_validate_json(scenario_file.read_text(encoding="utf-8"))
    source_trace = RuntimeTrace.model_validate_json(source_trace_file.read_text(encoding="utf-8"))
    target_trace = (
        RuntimeTrace.model_validate_json(target_trace_file.read_text(encoding="utf-8"))
        if target_trace_file is not None
        else None
    )
    return scenario, source_trace, target_trace


@cli.command()
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.argument("output_file", type=click.Path(path_type=Path))
@click.option("--locale", default="en-US", help="Target locale (en-US or de-DE)")
@click.option("--strict", is_flag=True, help="Fail on any errors")
@click.option(
    "--embed-macros/--no-macros",
    default=False,
    help="Embed separately supplied target-native modules (disabled by default)",
)
@click.option("--report", type=click.Path(path_type=Path), help="Save report to file")
def convert_cmd(
    input_file: Path,
    output_file: Path,
    locale: str,
    strict: bool,
    embed_macros: bool,
    report: Path | None,
) -> None:
    """Convert Excel file to ODS format.

    \b
    Examples:
        xlsliberator convert input.xlsx output.ods
        xlsliberator convert input.xlsm output.ods --locale de-DE
        xlsliberator convert input.xlsx output.ods --report report.json
    """
    click.echo(f"Converting {input_file} → {output_file}")

    try:
        result = convert(
            input_file,
            output_file,
            locale=locale,
            strict=strict,
            embed_macros=embed_macros,
        )

        # Display results
        click.echo("\n" + "=" * 60)
        if result.success:
            click.secho("✓ Conversion successful", fg="green", bold=True)
        else:
            click.secho("✗ Conversion failed", fg="red", bold=True)

        click.echo(f"Duration: {result.duration_seconds:.2f}s")
        click.echo(f"Cells: {result.total_cells:,}")
        click.echo(f"Formulas: {result.total_formulas:,}")
        click.echo(f"Sheets: {result.sheet_count}")

        if result.warnings:
            click.echo(f"\nWarnings: {len(result.warnings)}")
            for warning in result.warnings[:5]:  # Show first 5
                click.secho(f"  - {warning}", fg="yellow")

        if result.errors:
            click.echo(f"\nErrors: {len(result.errors)}")
            for error in result.errors:
                click.secho(f"  - {error}", fg="red")

        # Save report if requested
        if report:
            if report.suffix == ".md":
                result.save_markdown(report)
            else:
                result.save_json(report)
            click.echo(f"\nReport saved to: {report}")

        sys.exit(0 if result.success else 1)

    except Exception as e:
        click.secho(f"Error: {e}", fg="red")
        sys.exit(1)


@cli.command(name="inspect")
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.option("--json", "json_output", is_flag=True, help="Print structured JSON")
@click.option("--output", type=click.Path(path_type=Path), help="Write inventory JSON to file")
@click.option("--role", type=click.Choice(["source", "target"]), default="source")
def inspect_cmd(
    input_file: Path,
    json_output: bool,
    output: Path | None,
    role: Literal["source", "target"],
) -> None:
    """Inspect workbook parse inventory."""
    from xlsliberator.inspect_workbook import inspect_workbook, inventory_to_dict

    try:
        inventory = inspect_workbook(input_file, role=role)
        inventory_dict = inventory_to_dict(inventory)
        rendered = json.dumps(inventory_dict, indent=2)

        if output:
            output.write_text(rendered)

        if json_output or not output:
            click.echo(rendered)
        else:
            click.echo(f"Inventory saved to: {output}")

        sys.exit(0)
    except Exception as e:
        click.secho(f"Error: {e}", fg="red")
        sys.exit(1)


@cli.command(name="interactive-game-build")
@click.argument("source", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("output", type=click.Path(dir_okay=False, path_type=Path))
@click.option("--timeout", type=click.IntRange(min=1, max=600), default=120, show_default=True)
def interactive_game_build_cmd(source: Path, output: Path, timeout: int) -> None:
    """Build the public interactive-game ODS in pinned Docker LibreOffice."""
    from xlsliberator.interactive_game_showcase import build_target

    try:
        result = build_target(source, output, timeout_seconds=timeout)
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(json.dumps(result, indent=2, sort_keys=True))


@cli.command(name="interactive-game-run")
@click.argument("target", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("scenario", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("evidence_archive", type=click.Path(dir_okay=False, path_type=Path))
@click.option("--timeout", type=click.IntRange(min=1, max=600), default=180, show_default=True)
def interactive_game_run_cmd(
    target: Path,
    scenario: Path,
    evidence_archive: Path,
    timeout: int,
) -> None:
    """Run declared real-GUI actions and write a replay evidence ZIP."""
    from xlsliberator.interactive_game_showcase import run_gui_scenario

    try:
        payload = json.loads(scenario.read_text(encoding="utf-8"))
        if not isinstance(payload, dict) or not isinstance(payload.get("actions"), list):
            raise ValueError("scenario JSON must contain an actions array")
        result = run_gui_scenario(
            target,
            evidence_archive,
            list(payload["actions"]),
            timer_enabled=bool(payload.get("timer_enabled", True)),
            timeout_seconds=timeout,
        )
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(json.dumps(result, indent=2, sort_keys=True))


@cli.command(name="interactive-game-bundle")
@click.argument(
    "evidence_archives",
    nargs=5,
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.argument("replay_archive", type=click.Path(dir_okay=False, path_type=Path))
@click.option("--timeout", type=click.IntRange(min=1, max=600), default=180, show_default=True)
def interactive_game_bundle_cmd(
    evidence_archives: tuple[Path, ...],
    replay_archive: Path,
    timeout: int,
) -> None:
    """Bundle the five canonical GUI recordings into one public replay."""
    from xlsliberator.interactive_game_showcase import (
        PUBLIC_SCENARIOS,
        bundle_gui_replays,
    )

    try:
        mapped = dict(zip(PUBLIC_SCENARIOS, evidence_archives, strict=True))
        result = bundle_gui_replays(mapped, replay_archive, timeout_seconds=timeout)
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(json.dumps(result, indent=2, sort_keys=True))


@cli.command(name="inventory-diff")
@click.argument("source_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("target_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--output", type=click.Path(dir_okay=False, path_type=Path))
def inventory_diff_cmd(source_file: Path, target_file: Path, output: Path | None) -> None:
    """Generate canonical source/target inventories and explicit dispositions."""
    from xlsliberator.artifact_inventory import (
        diff_inventories,
        disposition_coverage_errors,
    )
    from xlsliberator.inspect_workbook import inspect_workbook

    try:
        source = inspect_workbook(source_file, role="source")
        target = inspect_workbook(target_file, role="target")
        difference = diff_inventories(source, target)
        source.dispositions = list(difference.dispositions)
        errors = disposition_coverage_errors(source)
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    rendered = json.dumps(
        {
            "source_inventory": source.model_dump(mode="json"),
            "target_inventory": target.model_dump(mode="json"),
            "diff": difference.model_dump(mode="json"),
            "coverage_errors": errors,
        },
        indent=2,
    )
    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(rendered + "\n", encoding="utf-8")
    click.echo(rendered)
    if errors:
        raise click.exceptions.Exit(1)


@cli.command(name="scenario-validate")
@click.argument("scenario_file", type=click.Path(exists=True, path_type=Path))
def scenario_validate_cmd(scenario_file: Path) -> None:
    """Validate and normalize a versioned scenario JSON file."""
    from xlsliberator.scenarios.models import Scenario

    try:
        scenario = Scenario.model_validate_json(scenario_file.read_text(encoding="utf-8"))
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(scenario.model_dump_json(indent=2))


@cli.command(name="trace-diff")
@click.argument("scenario_file", type=click.Path(exists=True, path_type=Path))
@click.argument("source_trace", type=click.Path(exists=True, path_type=Path))
@click.argument("target_trace", type=click.Path(exists=True, path_type=Path))
@click.option("--output", type=click.Path(path_type=Path))
def trace_diff_cmd(
    scenario_file: Path,
    source_trace: Path,
    target_trace: Path,
    output: Path | None,
) -> None:
    """Compare two runtime traces using the scenario's declared rules."""
    from xlsliberator.scenarios.diff import diff_traces
    from xlsliberator.scenarios.models import RuntimeTrace, Scenario

    try:
        scenario = Scenario.model_validate_json(scenario_file.read_text(encoding="utf-8"))
        source = RuntimeTrace.model_validate_json(source_trace.read_text(encoding="utf-8"))
        target = RuntimeTrace.model_validate_json(target_trace.read_text(encoding="utf-8"))
        result = diff_traces(source, target, scenario)
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    rendered = result.model_dump_json(indent=2)
    if output:
        output.write_text(rendered + "\n", encoding="utf-8")
    click.echo(rendered)
    if not result.equivalent:
        raise click.exceptions.Exit(1)


@cli.command(name="evidence-inspect")
@click.argument("bundle", type=click.Path(exists=True, file_okay=False, path_type=Path))
def evidence_inspect_cmd(bundle: Path) -> None:
    """Validate and print an evidence-bundle manifest."""
    from xlsliberator.scenarios.evidence import inspect_evidence_bundle

    try:
        manifest = inspect_evidence_bundle(bundle)
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(manifest.model_dump_json(indent=2))


@cli.command(name="scenario-run-target")
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("scenario_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option(
    "--environment",
    "environment_file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option("--output", type=click.Path(dir_okay=False, path_type=Path))
@click.option("--timeout", type=click.IntRange(min=1), default=120, show_default=True)
def scenario_run_target_cmd(
    workbook: Path,
    scenario_file: Path,
    environment_file: Path | None,
    output: Path | None,
    timeout: int,
) -> None:
    """Run a scenario in the authoritative Docker-backed LibreOffice target."""
    from xlsliberator.libreoffice_session_scenario_runner import (
        LibreOfficeSessionScenarioRunner,
    )
    from xlsliberator.scenarios.models import EnvironmentManifest, Scenario
    from xlsliberator.validation_models import GateExecutionStatus

    try:
        scenario = Scenario.model_validate_json(scenario_file.read_text(encoding="utf-8"))
        environment = (
            EnvironmentManifest.model_validate_json(environment_file.read_text(encoding="utf-8"))
            if environment_file
            else EnvironmentManifest()
        )
        trace = LibreOfficeSessionScenarioRunner(timeout_seconds=timeout).run(
            workbook,
            environment,
            scenario,
        )
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    rendered = trace.model_dump_json(indent=2)
    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(rendered + "\n", encoding="utf-8")
    click.echo(rendered)
    if trace.status is not GateExecutionStatus.PASSED:
        raise click.exceptions.Exit(1)


@cli.command(name="validate")
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.argument("output_ods", required=False, type=click.Path(path_type=Path))
@click.option(
    "--target",
    type=click.Choice(["libreoffice"]),
    default="libreoffice",
    help="Runtime target",
)
@click.option("--json", "json_output", is_flag=True, help="Print structured JSON")
@click.option(
    "--non-strict", is_flag=True, help="Report errors without strict certification failure"
)
@click.option(
    "--scenario-file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Exact scenario used for source and target execution",
)
@click.option(
    "--source-trace-file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Microsoft Excel source-oracle trace",
)
@click.option(
    "--target-trace-file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Existing Docker LibreOffice target trace; otherwise run it now",
)
def validate_cmd(
    input_file: Path,
    output_ods: Path | None,
    target: str,
    json_output: bool,
    non_strict: bool,
    scenario_file: Path | None,
    source_trace_file: Path | None,
    target_trace_file: Path | None,
) -> None:
    """Validate workbook transformation gates."""
    from xlsliberator.validation_runner import (
        ValidationPlan,
        ValidationRunner,
        parse_target_kind,
    )

    try:
        scenario, source_trace, target_trace = _load_formula_evidence(
            scenario_file,
            source_trace_file,
            target_trace_file,
        )
        if target_trace is None and scenario is not None and source_trace is not None:
            if output_ods is None:
                raise click.ClickException(
                    "an output ODS is required to execute the target scenario"
                )
            from xlsliberator.libreoffice_session_scenario_runner import (
                LibreOfficeSessionScenarioRunner,
            )

            target_trace = LibreOfficeSessionScenarioRunner().run(
                output_ods,
                source_trace.environment,
                scenario,
            )
        report = ValidationRunner(
            ValidationPlan(
                input_path=input_file,
                output_path=output_ods,
                target_kinds=parse_target_kind(target),
                strict=not non_strict,
                scenario=scenario,
                source_trace=source_trace,
                target_trace=target_trace,
            )
        ).run_all()

        if json_output:
            click.echo(report.to_json())
        else:
            click.echo(report.to_markdown())

        sys.exit(0 if report.certification.certified else 1)
    except Exception as e:
        click.secho(f"Error: {e}", fg="red")
        sys.exit(1)


@cli.command(name="transform-validated")
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.argument("output_file", type=click.Path(path_type=Path))
@click.option(
    "--target",
    multiple=True,
    type=click.Choice(["libreoffice"]),
    default=["libreoffice"],
    help="Runtime target; may be supplied multiple times",
)
@click.option("--strict/--non-strict", default=True, help="Fail on validation errors")
@click.option(
    "--max-repair-iterations", default=0, type=int, help="Deterministic repair iterations"
)
@click.option(
    "--embed-macros/--no-macros",
    default=False,
    help="Embed separately supplied target-native modules (disabled by default)",
)
@click.option("--no-agent", is_flag=True, hidden=True)
@click.option("--json", "json_output", is_flag=True, help="Print structured JSON")
@click.option(
    "--scenario-file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Exact externally supplied acceptance scenario",
)
@click.option(
    "--source-trace-file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Declared source-evidence trace from an independent test authority",
)
@click.option(
    "--target-trace-file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Existing Docker LibreOffice target trace; otherwise run it after conversion",
)
def transform_validated_cmd(
    input_file: Path,
    output_file: Path,
    target: tuple[str, ...],
    strict: bool,
    max_repair_iterations: int,
    embed_macros: bool,
    no_agent: bool,
    json_output: bool,
    scenario_file: Path | None,
    source_trace_file: Path | None,
    target_trace_file: Path | None,
) -> None:
    """Convert a workbook and run deterministic validation gates."""
    del no_agent
    from xlsliberator.validated_api import (
        ValidatedTransformationError,
        transform_validated,
    )

    try:
        scenario, source_trace, target_trace = _load_formula_evidence(
            scenario_file,
            source_trace_file,
            target_trace_file,
        )
        report = transform_validated(
            input_file,
            output_file,
            targets=list(target),
            strict=strict,
            max_repair_iterations=max_repair_iterations,
            embed_macros=embed_macros,
            use_agent=False,
            scenario=scenario,
            source_trace=source_trace,
            target_trace=target_trace,
        )
    except ValidatedTransformationError as e:
        report = e.report
    except Exception as e:
        click.secho(f"Error: {e}", fg="red")
        sys.exit(1)

    click.echo(report.to_json() if json_output else report.to_markdown())
    sys.exit(0 if report.certification.certified else 1)


@cli.command(name="libreoffice-mcp-serve")
@click.option("--host", default="127.0.0.1", help="Loopback address to bind to")
@click.option("--port", default=8000, help="Port number")
@click.option(
    "--workspace-root",
    "workspace_roots",
    multiple=True,
    type=click.Path(exists=True, file_okay=False, path_type=Path),
    help="Host root accessible to MCP tools; repeat for multiple roots",
)
def libreoffice_mcp_serve_cmd(
    host: str,
    port: int,
    workspace_roots: tuple[Path, ...],
) -> None:
    """Start the stateful LibreOffice session MCP service.

    \b
    Every document operation requires an explicit session ID.

    \b
    Examples:
        xlsliberator libreoffice-mcp-serve
        xlsliberator libreoffice-mcp-serve --port 9000

    \b
    Client endpoint: http://<host>:<port>/mcp
    """
    _serve_libreoffice_mcp(host, port, workspace_roots)


@cli.command(name="mcp-serve", deprecated=True)
@click.option("--host", default="127.0.0.1", help="Loopback address to bind to")
@click.option("--port", default=8000, help="Port number")
@click.option(
    "--workspace-root",
    "workspace_roots",
    multiple=True,
    type=click.Path(exists=True, file_okay=False, path_type=Path),
    help="Host root accessible to MCP tools; repeat for multiple roots",
)
def mcp_serve_cmd(host: str, port: int, workspace_roots: tuple[Path, ...]) -> None:
    """Deprecated alias for ``libreoffice-mcp-serve``."""
    _serve_libreoffice_mcp(host, port, workspace_roots)


def _serve_libreoffice_mcp(
    host: str,
    port: int,
    workspace_roots: tuple[Path, ...],
) -> None:
    from xlsliberator.mcp_server import serve

    if workspace_roots:
        import os

        os.environ["XLSLIBERATOR_WORKSPACE_ROOTS"] = os.pathsep.join(
            str(path.resolve()) for path in workspace_roots
        )

    click.echo(f"Starting stateful LibreOffice runtime MCP server on {host}:{port}")
    click.echo(f"Client endpoint: http://{host}:{port}/mcp")
    click.echo("\nPress Ctrl+C to stop\n")

    try:
        serve(host=host, port=port)
    except KeyboardInterrupt:
        click.echo("\nShutting down...")
        sys.exit(0)
    except Exception as e:
        click.secho(f"Error: {e}", fg="red")
        sys.exit(1)


@cli.command(name="corpus-mcp-serve")
@click.option("--host", default="127.0.0.1", help="Loopback address to bind to")
@click.option("--port", default=8010, help="Port number")
def corpus_mcp_serve_cmd(host: str, port: int) -> None:
    """Start the public corpus and reviewer-gated hidden-suite MCP service."""
    from xlsliberator.migration_services_mcp import serve_corpus

    click.echo(f"Starting migration corpus MCP server on {host}:{port}")
    serve_corpus(host, port)


@cli.command(name="buildfarm-mcp-serve")
@click.option("--host", default="127.0.0.1", help="Loopback address to bind to")
@click.option("--port", default=8020, help="Port number")
def buildfarm_mcp_serve_cmd(host: str, port: int) -> None:
    """Start the authorized LibreOffice build-farm contract MCP service."""
    from xlsliberator.migration_services_mcp import serve_buildfarm

    click.echo(f"Starting LibreOffice build-farm MCP server on {host}:{port}")
    serve_buildfarm(host, port)


@cli.command(name="web-serve")
@click.option("--host", default="0.0.0.0", help="Host address to bind to")  # nosec B104
@click.option("--port", default=8080, help="Port number")
@click.option("--reload", is_flag=True, help="Reload on code changes")
def web_serve_cmd(host: str, port: int, reload: bool) -> None:
    """Start the browser web app."""
    try:
        import uvicorn
    except ImportError as e:
        raise click.ClickException('Install web extras first: pip install -e ".[web]"') from e

    uvicorn.run(
        "xlsliberator.web.app:create_app",
        host=host,
        port=port,
        reload=reload,
        factory=True,
    )


@cli.command(name="cleanup-jobs")
@click.option(
    "--data-dir",
    type=click.Path(path_type=Path),
    default=Path("/data"),
    help="Web data directory",
)
@click.option("--older-than-hours", type=int, default=24, help="Delete jobs older than this")
def cleanup_jobs_cmd(data_dir: Path, older_than_hours: int) -> None:
    """Delete old web job artifacts."""
    from xlsliberator.web.cleanup import cleanup_old_jobs

    try:
        deleted = cleanup_old_jobs(data_dir, timedelta(hours=older_than_hours))
    except Exception as e:
        raise click.ClickException(str(e)) from e
    click.echo(f"Deleted {len(deleted)} job director{'y' if len(deleted) == 1 else 'ies'}")


@cli.command(name="capability-report")
@click.option(
    "--corpus",
    "corpus_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=Path("corpus/manifest.json"),
    show_default=True,
)
@click.option(
    "--evidence",
    "evidence_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    required=True,
)
@click.option(
    "--release-inputs",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    required=True,
)
@click.option(
    "--previous",
    "previous_report",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option("--report-id", default="current", show_default=True)
@click.option("--json-output", type=click.Path(dir_okay=False, path_type=Path))
@click.option("--markdown-output", type=click.Path(dir_okay=False, path_type=Path))
@click.option("--release-notes-output", type=click.Path(dir_okay=False, path_type=Path))
@click.option(
    "--check-release", is_flag=True, help="Exit non-zero unless every release gate passes"
)
def capability_report_cmd(
    corpus_path: Path,
    evidence_path: Path,
    release_inputs: Path,
    previous_report: Path | None,
    report_id: str,
    json_output: Path | None,
    markdown_output: Path | None,
    release_notes_output: Path | None,
    check_release: bool,
) -> None:
    """Generate capability claims from validated corpus/runtime evidence."""
    from xlsliberator.capability_matrix import (
        CapabilityReport,
        ReleaseInputs,
        generate_capability_report,
        load_measurements,
    )
    from xlsliberator.conformance_corpus import CorpusManifest
    from xlsliberator.formula_corpus import collect_formula_corpus_statistics

    try:
        corpus = CorpusManifest.load(corpus_path)
        if integrity_errors := corpus.verify_files(corpus_path.parent.parent):
            raise ValueError("; ".join(integrity_errors))
        measurements = load_measurements(evidence_path)
        inputs = ReleaseInputs.model_validate_json(release_inputs.read_text(encoding="utf-8"))
        previous = (
            CapabilityReport.model_validate_json(previous_report.read_text(encoding="utf-8"))
            if previous_report is not None
            else None
        )
        generated = generate_capability_report(
            corpus=corpus,
            measurements=measurements,
            release_inputs=inputs,
            formula_corpus=collect_formula_corpus_statistics(
                corpus_path.resolve().parent.parent / "tests" / "fixtures" / "formulas",
                display_path="tests/fixtures/formulas",
            ),
            previous=previous,
            report_id=report_id,
        )
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    rendered_json = generated.model_dump_json(indent=2) + "\n"
    rendered_markdown = generated.to_markdown()
    for path, content in (
        (json_output, rendered_json),
        (markdown_output, rendered_markdown),
        (release_notes_output, generated.to_release_notes()),
    ):
        if path is not None:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content, encoding="utf-8")
    if json_output is None and markdown_output is None:
        click.echo(rendered_markdown, nl=False)
    if check_release and not generated.release_ready:
        raise click.exceptions.Exit(1)


@cli.command(name="corpus-report")
@click.option(
    "--corpus",
    "corpus_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=Path("corpus/manifest.json"),
    show_default=True,
)
@click.option(
    "--executions",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    required=True,
)
@click.option(
    "--previous",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option("--output", type=click.Path(dir_okay=False, path_type=Path))
def corpus_report_cmd(
    corpus_path: Path,
    executions: Path,
    previous: Path | None,
    output: Path | None,
) -> None:
    """Generate corpus statistics and trends from versioned dispositions."""
    from xlsliberator.conformance_corpus import (
        CorpusManifest,
        CorpusStatistics,
        corpus_trend_report,
        load_corpus_executions,
    )

    try:
        manifest = CorpusManifest.load(corpus_path)
        if integrity_errors := manifest.verify_files(corpus_path.parent.parent):
            raise ValueError("; ".join(integrity_errors))
        previous_statistics = None
        if previous is not None:
            previous_payload = json.loads(previous.read_text(encoding="utf-8"))
            previous_statistics = CorpusStatistics.model_validate(previous_payload["current"])
        report = corpus_trend_report(
            manifest,
            load_corpus_executions(executions),
            previous=previous_statistics,
        )
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    rendered = report.model_dump_json(indent=2) + "\n"
    if output is None:
        click.echo(rendered, nl=False)
    else:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(rendered, encoding="utf-8")


@cli.command(name="demo-corpus-validate")
@click.option(
    "--manifest",
    "manifest_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=Path("tests/corpus/manifests/episodes.json"),
    show_default=True,
)
@click.option(
    "--search-index",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=Path("tests/corpus/manifests/search-index.json"),
    show_default=True,
)
def demo_corpus_validate_cmd(manifest_path: Path, search_index: Path) -> None:
    """Validate serious episode layout, integrity, and search metadata."""
    from xlsliberator.demo_corpus import DemoCorpusManifest

    try:
        manifest = DemoCorpusManifest.load(manifest_path)
        repository_root = manifest_path.resolve().parents[3]
        errors = manifest.verify(repository_root)
        checked_index = json.loads(search_index.read_text(encoding="utf-8"))
        if checked_index != manifest.search_index():
            errors.append("checked-in search index does not match episode manifest")
        if errors:
            raise ValueError("; ".join(errors))
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(f"Validated {len(manifest.episodes)} serious migration episodes")


@cli.command(name="demo-corpus-search")
@click.option(
    "--manifest",
    "manifest_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=Path("tests/corpus/manifests/episodes.json"),
    show_default=True,
)
@click.option("--query", required=True)
@click.option("--subset", type=click.Choice(["pr", "nightly", "security"]))
def demo_corpus_search_cmd(
    manifest_path: Path,
    query: str,
    subset: str | None,
) -> None:
    """Search public corpus metadata for migration candidates."""
    from typing import cast

    from xlsliberator.demo_corpus import (
        DemoCorpusManifest,
        SubsetName,
        search_demo_corpus,
    )

    try:
        manifest = DemoCorpusManifest.load(manifest_path)
        matches = search_demo_corpus(
            manifest,
            query=query,
            subset=cast(SubsetName | None, subset),
        )
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(json.dumps({"matches": matches}, indent=2, sort_keys=True))


@cli.command(name="demo-corpus-report")
@click.option(
    "--manifest",
    "manifest_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=Path("tests/corpus/manifests/episodes.json"),
    show_default=True,
)
@click.option(
    "--results",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    required=True,
)
@click.option("--output", type=click.Path(dir_okay=False, path_type=Path))
def demo_corpus_report_cmd(
    manifest_path: Path,
    results: Path,
    output: Path | None,
) -> None:
    """Generate feature and format dispositions from serious episode results."""
    from xlsliberator.demo_corpus import (
        DemoCorpusManifest,
        generate_demo_corpus_report,
        load_demo_results,
    )

    try:
        manifest = DemoCorpusManifest.load(manifest_path)
        report = generate_demo_corpus_report(manifest, load_demo_results(results))
    except Exception as exc:
        raise click.ClickException(str(exc)) from exc
    rendered = report.model_dump_json(indent=2) + "\n"
    if output is None:
        click.echo(rendered, nl=False)
    else:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(rendered, encoding="utf-8")


if __name__ == "__main__":
    cli()
