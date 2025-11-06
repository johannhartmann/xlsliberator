"""CLI interface for xlsliberator (Phase F12)."""

import sys
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


if __name__ == "__main__":
    cli()
