"""Deterministic mutation campaign for the interactive-game migration."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Final


@dataclass(frozen=True, slots=True)
class MutationSpec:
    """One exact source mutation expected to be killed by public acceptance."""

    mutation_id: str
    behavior: str
    original: str
    replacement: str


MUTATIONS: Final = (
    MutationSpec(
        mutation_id="reverse-left-control",
        behavior="left movement",
        original="return _move(state, row_delta=0, column_delta=-1)",
        replacement="return _move(state, row_delta=0, column_delta=1)",
    ),
    MutationSpec(
        mutation_id="disable-clockwise-rotation",
        behavior="clockwise rotation",
        original="rotation=(state.active.rotation + 1) % orientation_count",
        replacement="rotation=state.active.rotation % orientation_count",
    ),
    MutationSpec(
        mutation_id="shorten-soft-drop",
        behavior="two-row quick drop",
        original="candidate = replace(state.active, row=state.active.row + 2)",
        replacement="candidate = replace(state.active, row=state.active.row + 1)",
    ),
    MutationSpec(
        mutation_id="double-timer-fall",
        behavior="bounded timer tick",
        original="candidate = replace(state.active, row=state.active.row + 1)",
        replacement="candidate = replace(state.active, row=state.active.row + 2)",
    ),
    MutationSpec(
        mutation_id="remove-line-score",
        behavior="completed-line scoring",
        original="LINE_POINTS: Final = 100",
        replacement="LINE_POINTS: Final = 0",
    ),
    MutationSpec(
        mutation_id="break-pause-state",
        behavior="pause control",
        original="return replace(state, phase=GamePhase.PAUSED, event_index=state.event_index + 1)",
        replacement=(
            "return replace(state, phase=GamePhase.STOPPED, event_index=state.event_index + 1)"
        ),
    ),
    MutationSpec(
        mutation_id="slow-initial-timer",
        behavior="source-derived timer progression",
        original="if lines <= 9:\n        return 160",
        replacement="if lines <= 9:\n        return 999",
    ),
)

_FAILED_TESTS = re.compile(r"(?:^|\s)[1-9][0-9]* failed(?:,|\s|$)")


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _apply_mutation(source: str, mutation: MutationSpec) -> str:
    occurrences = source.count(mutation.original)
    if occurrences != 1:
        raise ValueError(f"{mutation.mutation_id} expected one source match, found {occurrences}")
    return source.replace(mutation.original, mutation.replacement, 1)


def run_campaign(repository_root: Path, output_dir: Path) -> dict[str, object]:
    """Run every mutant against an isolated copy of the public engine tests."""
    root = repository_root.resolve()
    package = root / "demos/interactive-game/candidate/candidate_interactive_game"
    engine = package / "engine.py"
    validator = root / "tests/unit/test_interactive_game_engine.py"
    if not engine.is_file() or not validator.is_file() or not package.is_dir():
        raise FileNotFoundError("interactive-game mutation inputs are incomplete")

    output = output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    source_text = engine.read_text(encoding="utf-8")
    source_sha256 = _sha256(engine)
    validator_sha256_before = _sha256(validator)
    cases: list[dict[str, object]] = []

    for mutation in MUTATIONS:
        case_root = output / "work" / mutation.mutation_id
        mutant_package = case_root / "candidate/candidate_interactive_game"
        tests_dir = case_root / "tests"
        shutil.copytree(package, mutant_package)
        tests_dir.mkdir(parents=True)
        shutil.copy2(validator, tests_dir / validator.name)
        mutant_engine = mutant_package / engine.name
        mutant_engine.write_text(
            _apply_mutation(source_text, mutation),
            encoding="utf-8",
        )

        command = [
            sys.executable,
            "-m",
            "pytest",
            "-p",
            "no:cacheprovider",
            "-q",
            f"tests/{validator.name}",
        ]
        environment = dict(os.environ)
        environment["XLSLIBERATOR_INTERACTIVE_CANDIDATE_ROOT"] = str(case_root / "candidate")
        completed = subprocess.run(
            command,
            cwd=case_root,
            env=environment,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            check=False,
            timeout=120,
        )
        log = completed.stdout
        log_path = output / f"{mutation.mutation_id}.log"
        log_path.write_text(log, encoding="utf-8")
        if completed.returncode == 0:
            outcome = "survived"
        elif completed.returncode == 1 and _FAILED_TESTS.search(log):
            outcome = "killed"
        else:
            outcome = "invalid"
        cases.append(
            {
                "mutation_id": mutation.mutation_id,
                "behavior": mutation.behavior,
                "outcome": outcome,
                "return_code": completed.returncode,
                "mutant_sha256": _sha256(mutant_engine),
                "log": log_path.name,
                "log_sha256": _sha256(log_path),
                "command": command,
            }
        )

    shutil.rmtree(output / "work")
    killed = sum(case["outcome"] == "killed" for case in cases)
    validator_sha256_after = _sha256(validator)
    passed = (
        killed == len(MUTATIONS)
        and validator_sha256_before == validator_sha256_after
        and _sha256(engine) == source_sha256
    )
    report: dict[str, object] = {
        "schema_version": "1.0.0",
        "campaign_id": "interactive-game-public-acceptance",
        "status": "PASSED" if passed else "FAILED",
        "source": str(engine.relative_to(root)),
        "source_sha256_before": source_sha256,
        "source_sha256_after": _sha256(engine),
        "validator": str(validator.relative_to(root)),
        "validator_sha256_before": validator_sha256_before,
        "validator_sha256_after": validator_sha256_after,
        "required_kill_rate": 1.0,
        "total": len(MUTATIONS),
        "killed": killed,
        "kill_rate": killed / len(MUTATIONS),
        "cases": cases,
    }
    report_path = output / "mutation-report.json"
    report_path.write_text(
        json.dumps(report, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return report


def main(argv: list[str] | None = None) -> int:
    """Run the campaign and fail unless every source-derived mutant is killed."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--repository-root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    report = run_campaign(args.repository_root, args.output)
    print(json.dumps(report, sort_keys=True))
    return 0 if report["status"] == "PASSED" else 1


if __name__ == "__main__":
    raise SystemExit(main())
