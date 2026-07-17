.PHONY: help fmt lint typecheck test test-unit test-integration test-cov security audit bandit all clean install pre-commit

DOCKER_TEST := docker compose run --rm test
DOCKER_ORCHESTRATOR := docker compose --profile ci-orchestrator run --rm test-orchestrator
DOCKER_SECURITY := docker compose --profile ci-orchestrator run --rm security-audit

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
	docker compose build test

fmt:
	@echo "==> Formatting code with ruff..."
	$(DOCKER_TEST) ruff format .

lint:
	@echo "==> Linting code with ruff..."
	$(DOCKER_TEST) ruff check .

typecheck:
	@echo "==> Type checking with mypy..."
	$(DOCKER_TEST) mypy src/

test:
	@echo "==> Running all tests..."
	$(DOCKER_TEST) pytest -v

test-unit:
	@echo "==> Running unit tests..."
	$(DOCKER_TEST) pytest -v -m "not integration"

test-integration:
	@echo "==> Running integration tests..."
	@echo "Integration tests are orchestrated from Docker against disposable office containers."
	mkdir -p artifacts/runtime-tmp artifacts/pytest-tmp artifacts/ci
	$(DOCKER_ORCHESTRATOR) python tools/ci_check.py office

test-cov:
	@echo "==> Running tests with coverage..."
	$(DOCKER_TEST) pytest -v --cov=xlsliberator --cov-report=term --cov-report=html
	@echo ""
	@echo "Coverage report generated in htmlcov/index.html"

audit:
	@echo "==> Auditing dependencies for vulnerabilities..."
	$(DOCKER_SECURITY) pip-audit --desc

bandit:
	@echo "==> Running Bandit security scanner..."
	$(DOCKER_TEST) bandit -r src/ -f screen -c pyproject.toml

security: audit bandit
	@echo ""
	@echo "==> Security checks completed"

pre-commit:
	@echo "Pre-commit execution is containerized; host hook installation is intentionally disabled."
	$(DOCKER_TEST) pre-commit run --all-files

all: fmt lint typecheck test-unit security
	@echo ""
	@echo "=========================================="
	@echo "✓ All checks passed!"
	@echo "=========================================="

clean:
	@echo "==> Cleaning generated files..."
	rm -rf build/ dist/ *.egg-info .pytest_cache .mypy_cache .ruff_cache htmlcov/ .coverage
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
