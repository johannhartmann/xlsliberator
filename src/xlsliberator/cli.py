"""CLI interface for xlsliberator (Phase F12)."""

import json
import sys
from datetime import timedelta
from pathlib import Path

import click

from xlsliberator.api import convert


@click.group()
@click.version_option(version="0.1.0")
def cli() -> None:
    """XLSLiberator - Excel to LibreOffice Calc converter."""
    pass


@cli.command()
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.argument("output_file", type=click.Path(path_type=Path))
@click.option("--locale", default="en-US", help="Target locale (en-US or de-DE)")
@click.option("--strict", is_flag=True, help="Fail on any errors")
@click.option("--no-macros", is_flag=True, help="Skip VBA macro translation")
@click.option("--report", type=click.Path(path_type=Path), help="Save report to file")
def convert_cmd(
    input_file: Path,
    output_file: Path,
    locale: str,
    strict: bool,
    no_macros: bool,
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
            embed_macros=not no_macros,
            use_agent=True,
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
def inspect_cmd(input_file: Path, json_output: bool, output: Path | None) -> None:
    """Inspect workbook parse inventory."""
    from xlsliberator.inspect_workbook import inspect_workbook, inventory_to_dict

    try:
        inventory = inspect_workbook(input_file)
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


@cli.command(name="validate")
@click.argument("input_file", type=click.Path(exists=True, path_type=Path))
@click.argument("output_ods", required=False, type=click.Path(path_type=Path))
@click.option(
    "--target",
    type=click.Choice(["libreoffice", "openoffice", "both"]),
    default="both",
    help="Runtime target",
)
@click.option("--json", "json_output", is_flag=True, help="Print structured JSON")
@click.option(
    "--non-strict", is_flag=True, help="Report errors without strict certification failure"
)
def validate_cmd(
    input_file: Path,
    output_ods: Path | None,
    target: str,
    json_output: bool,
    non_strict: bool,
) -> None:
    """Validate workbook transformation gates."""
    from xlsliberator.validation_runner import (
        ValidationPlan,
        ValidationRunner,
        parse_target_kind,
    )

    try:
        report = ValidationRunner(
            ValidationPlan(
                input_path=input_file,
                output_path=output_ods,
                target_kinds=parse_target_kind(target),
                strict=not non_strict,
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
    type=click.Choice(["libreoffice", "openoffice", "both"]),
    default=["both"],
    help="Runtime target; may be supplied multiple times",
)
@click.option("--strict/--non-strict", default=True, help="Fail on validation errors")
@click.option(
    "--max-repair-iterations", default=0, type=int, help="Deterministic repair iterations"
)
@click.option("--no-macros", is_flag=True, help="Skip VBA macro translation")
@click.option("--no-agent", is_flag=True, help="Disable agent-based VBA rewriting")
@click.option("--json", "json_output", is_flag=True, help="Print structured JSON")
def transform_validated_cmd(
    input_file: Path,
    output_file: Path,
    target: tuple[str, ...],
    strict: bool,
    max_repair_iterations: int,
    no_macros: bool,
    no_agent: bool,
    json_output: bool,
) -> None:
    """Convert a workbook and run validation gates."""
    from xlsliberator.validated_api import (
        ValidatedTransformationError,
        transform_validated,
    )

    try:
        report = transform_validated(
            input_file,
            output_file,
            targets=list(target),
            strict=strict,
            max_repair_iterations=max_repair_iterations,
            embed_macros=not no_macros,
            use_agent=not no_agent,
        )
    except ValidatedTransformationError as e:
        report = e.report
    except Exception as e:
        click.secho(f"Error: {e}", fg="red")
        sys.exit(1)

    click.echo(report.to_json() if json_output else report.to_markdown())
    sys.exit(0 if report.certification.certified else 1)


@cli.command(name="mcp-serve")
@click.option("--host", default="0.0.0.0", help="Host address to bind to")  # nosec B104
@click.option("--port", default=8000, help="Port number")
def mcp_serve_cmd(host: str, port: int) -> None:
    """Start MCP server with HTTP streaming transport.

    \b
    Exposes LibreOffice UNO operations as MCP tools for Claude Agent SDK integration.

    \b
    Examples:
        xlsliberator mcp-serve
        xlsliberator mcp-serve --port 9000
        xlsliberator mcp-serve --host 127.0.0.1 --port 8080

    \b
    Client endpoint: http://<host>:<port>/mcp
    """
    from xlsliberator.mcp_server import serve

    click.echo(f"Starting LibreOffice UNO MCP server on {host}:{port}")
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


if __name__ == "__main__":
    cli()
