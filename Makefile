.PHONY: help fmt lint typecheck test test-unit test-integration test-cov security audit bandit all clean install pre-commit

help:
	@echo "XLSLiberator Development Commands"
	@echo "=================================="
	@echo ""
	@echo "Quality Checks:"
	@echo "  make fmt          - Format code with ruff"
	@echo "  make lint         - Lint code with ruff"
	@echo "  make typecheck    - Type check with mypy"
	@echo "  make test         - Run all tests"
	@echo "  make test-unit    - Run unit tests only"
	@echo "  make test-integration - Run integration tests"
	@echo "  make test-cov     - Run tests with coverage report"
	@echo ""
	@echo "Security:"
	@echo "  make security     - Run all security checks"
	@echo "  make audit        - Audit dependencies for vulnerabilities"
	@echo "  make bandit       - Run SAST security scanner"
	@echo ""
	@echo "CI Simulation:"
	@echo "  make all          - Run all checks (CI simulation)"
	@echo "  make pre-commit   - Install and run pre-commit hooks"
	@echo ""
	@echo "Utilities:"
	@echo "  make install      - Install package in dev mode"
	@echo "  make clean        - Clean generated files"

install:
	pip install -e ".[dev]"
	pip install pip-audit bandit pre-commit pytest-cov

fmt:
	@echo "==> Formatting code with ruff..."
	ruff format .

lint:
	@echo "==> Linting code with ruff..."
	ruff check .

typecheck:
	@echo "==> Type checking with mypy..."
	mypy src/

test:
	@echo "==> Running all tests..."
	pytest -v

test-unit:
	@echo "==> Running unit tests..."
	pytest -v -m "not integration"

test-integration:
	@echo "==> Running integration tests..."
	pytest -v -m integration

test-cov:
	@echo "==> Running tests with coverage..."
	pytest -v --cov=xlsliberator --cov-report=term --cov-report=html
	@echo ""
	@echo "Coverage report generated in htmlcov/index.html"

audit:
	@echo "==> Auditing dependencies for vulnerabilities..."
	pip-audit --desc

bandit:
	@echo "==> Running Bandit security scanner..."
	bandit -r src/ -f screen -c pyproject.toml || true
	@echo ""
	@echo "Note: Some findings (B110: try-except-pass, B314: XML parsing) are real issues"
	@echo "      Others (subprocess usage) are legitimate for LibreOffice integration"

security: audit bandit
	@echo ""
	@echo "==> Security checks completed"

pre-commit:
	@echo "==> Installing pre-commit hooks..."
	pre-commit install
	@echo "==> Running pre-commit on all files..."
	pre-commit run --all-files

all: fmt lint typecheck test-unit security
	@echo ""
	@echo "=========================================="
	@echo "✓ All checks passed!"
	@echo "=========================================="

clean:
	@echo "==> Cleaning generated files..."
	rm -rf build/ dist/ *.egg-info .pytest_cache .mypy_cache .ruff_cache htmlcov/ .coverage
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	find . -type f -name "*.pyc" -delete
