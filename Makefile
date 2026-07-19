.PHONY: help fmt lint skill-lint typecheck test test-unit test-integration test-docker-web test-package test-cov security audit bandit all clean install pre-commit

DOCKER_TEST := docker compose run --rm test
DOCKER_RUNNER := docker compose --profile ci-runner run --rm test-runner
DOCKER_SECURITY := docker compose --profile ci-runner run --rm security-audit

help:
	@echo "XLSLiberator Development Commands"
	@echo "=================================="
	@echo ""
	@echo "Quality Checks:"
	@echo "  make fmt          - Format code with ruff"
	@echo "  make lint         - Lint code with ruff"
	@echo "  make skill-lint   - Validate Deep Agents migration skills"
	@echo "  make typecheck    - Type check with mypy"
	@echo "  make test         - Run all tests"
	@echo "  make test-unit    - Run unit tests only"
	@echo "  make test-integration - Run integration tests"
	@echo "  make test-docker-web - Run the blocking Docker web smoke"
	@echo "  make test-package - Build and validate distributions"
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

skill-lint:
	@echo "==> Validating Deep Agents migration skills..."
	$(DOCKER_TEST) python -m xlsliberator.skill_validation skills

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
	@echo "Integration tests run from Docker against disposable office containers."
	mkdir -p artifacts/runtime-tmp artifacts/pytest-tmp artifacts/ci
	$(DOCKER_RUNNER) python tools/ci_check.py office

test-docker-web:
	@echo "==> Running blocking Docker web smoke..."
	mkdir -p artifacts/runtime-tmp artifacts/pytest-tmp artifacts/ci
	docker compose --profile ci-runner run --rm \
		-e DOCKER_TESTS=1 -e XLSLIBERATOR_FAIL_ON_SKIP=1 \
		test-runner python tools/ci_check.py docker-web

test-package:
	@echo "==> Building and validating distributions..."
	mkdir -p dist
	$(DOCKER_TEST) python tools/ci_check.py package

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

all: fmt lint skill-lint typecheck test-unit test-integration test-docker-web test-package security
	@echo ""
	@echo "=========================================="
	@echo "✓ All checks passed!"
	@echo "=========================================="

clean:
	@echo "==> Cleaning generated files..."
	rm -rf build/ dist/ *.egg-info .pytest_cache .mypy_cache .ruff_cache htmlcov/ .coverage
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
