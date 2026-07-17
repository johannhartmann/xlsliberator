"""Integration coverage for backend FormulaParser integration."""

import pytest

from xlsliberator.calc_backend import discover_backends


@pytest.mark.integration
@pytest.mark.docker
def test_backend_formula_parser_hook(skip_if_no_lo: None) -> None:
    """Backend formula parser should use FormulaParser only through Docker."""
    backends = discover_backends()
    if not backends:
        pytest.skip("No office backend discovered")

    result = backends[0].parse_formula_text("=SUM(1;2;3)", sheet_name="Sheet1")
    if result.details.get("target_parser") != "docker_uno_formula_parser":
        pytest.skip(
            "UNO FormulaParser unavailable: "
            f"{result.details.get('target_parser_unavailable', result.error)}"
        )

    assert result.success
    assert result.details["target_parser"] == "docker_uno_formula_parser"
    assert result.tokens

    invalid_result = backends[0].parse_formula_text("=SUM(1;2;3", sheet_name="Sheet1")

    assert not invalid_result.success
    assert invalid_result.details["target_parser"] == "docker_uno_formula_parser"
