.PHONY: help fmt lint typecheck test check clean install

help:
	@echo "XLSLiberator Development Commands"
	@echo "=================================="
	@echo "make install    - Install dependencies with uv"
	@echo "make fmt        - Format code with ruff"
	@echo "make lint       - Lint code with ruff"
	@echo "make typecheck  - Type check with mypy"
	@echo "make test       - Run tests with pytest"
	@echo "make check      - Run all quality checks (fmt + lint + typecheck + test)"
	@echo "make clean      - Clean build artifacts"

install:
	uv pip install -e ".[dev]"

fmt:
	ruff format src/ tests/

lint:
	ruff check src/ tests/

typecheck:
	mypy src/

test:
	pytest -q

check: fmt lint typecheck test
	@echo "✅ All quality checks passed!"

clean:
	rm -rf build/ dist/ *.egg-info
	rm -rf .pytest_cache/ .mypy_cache/ .ruff_cache/
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
