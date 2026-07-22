#!/usr/bin/env python3
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # Python < 3.11
    import tomli as tomllib


ROOT = Path(__file__).resolve().parents[1]
TASKS_DIR = ROOT / "tasks"
EXPECTED_TASKS = 210


def level_from_metadata(metadata: dict) -> str:
    if metadata.get("final_difficulty"):
        return str(metadata["final_difficulty"])
    raw = str(metadata.get("difficulty_level", ""))
    for level in ("L1", "L2", "L3", "L4"):
        if raw.startswith(level):
            return level
    return "unknown"


def main() -> None:
    task_dirs = sorted(path for path in TASKS_DIR.iterdir() if (path / "task.toml").is_file())
    errors: list[str] = []
    levels: Counter[str] = Counter()
    sources: Counter[str] = Counter()

    if len(task_dirs) != EXPECTED_TASKS:
        errors.append(f"expected {EXPECTED_TASKS} tasks, found {len(task_dirs)}")

    for task_dir in task_dirs:
        required = [
            task_dir / "instruction.md",
            task_dir / "task.toml",
            task_dir / "environment" / "Dockerfile",
            task_dir / "tests" / "test.sh",
            task_dir / "tests" / "test_outputs.py",
        ]
        for path in required:
            if not path.exists():
                errors.append(f"missing {path.relative_to(ROOT)}")

        data = tomllib.loads((task_dir / "task.toml").read_text())
        metadata = data.get("metadata", {})
        levels[level_from_metadata(metadata)] += 1
        source = "v2_final_80" if metadata.get("final_dataset") == "docops_v2_final_80" else "v1_final_130"
        sources[source] += 1

        if (task_dir / "solution").exists():
            errors.append(f"public package should not include solution/: {task_dir.relative_to(ROOT)}")

    print(f"tasks: {len(task_dirs)}")
    print("sources:", dict(sorted(sources.items())))
    print("levels:", dict(sorted(levels.items())))

    if errors:
        print("errors:")
        for error in errors:
            print(f"- {error}")
        sys.exit(1)

    print("package verification passed")


if __name__ == "__main__":
    main()
