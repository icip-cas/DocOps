#!/usr/bin/env python3
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path


IGNORE_NAMES = {
    ".DS_Store",
    "__pycache__",
    ".pytest_cache",
    "solution",
}


def ignore_runtime(_src: str, names: list[str]) -> set[str]:
    ignored: set[str] = set()
    for name in names:
        if name in IGNORE_NAMES:
            ignored.add(name)
        if name.endswith((".pyc", ".pyo")):
            ignored.add(name)
        if name.startswith("harbor_results") or name.startswith("results"):
            ignored.add(name)
        if name.startswith(".harbor_") and name.endswith(".env"):
            ignored.add(name)
    return ignored


def strip_task_toml(path: Path) -> None:
    text = path.read_text()
    text = re.sub(r'(?m)^skills_dir\s*=\s*".*"\n?', "", text)
    path.write_text(text)


def strip_instruction(path: Path) -> None:
    if not path.is_file():
        return
    text = path.read_text()
    text = re.sub(r"\n?<available_skills>.*?</available_skills>\n?", "\n", text, flags=re.S)
    path.write_text(text)


def strip_dockerfile(path: Path) -> None:
    if not path.is_file():
        return
    lines = []
    for line in path.read_text().splitlines():
        if "COPY skills/" in line or "COPY skills " in line:
            continue
        lines.append(line)
    path.write_text("\n".join(lines) + "\n")


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("usage: materialize_no_skill.py SOURCE_TASKS_DIR TARGET_TASKS_DIR")

    source = Path(sys.argv[1]).resolve()
    target = Path(sys.argv[2]).resolve()
    if not source.is_dir():
        raise SystemExit(f"source tasks directory not found: {source}")

    if target.exists():
        shutil.rmtree(target)
    shutil.copytree(source, target, ignore=ignore_runtime)

    for task_toml in target.glob("*/task.toml"):
        task_dir = task_toml.parent
        strip_task_toml(task_toml)
        strip_instruction(task_dir / "instruction.md")
        strip_dockerfile(task_dir / "environment" / "Dockerfile")
        shutil.rmtree(task_dir / "environment" / "skills", ignore_errors=True)

    print(f"materialized no-skill tasks: {target}")


if __name__ == "__main__":
    main()

