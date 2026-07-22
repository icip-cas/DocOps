#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # Python < 3.11
    import tomli as tomllib


ROOT = Path(__file__).resolve().parents[1]
TASKS_DIR = ROOT / "tasks"
OUT_DIR = ROOT / "metadata"


def norm_level(metadata: dict) -> str:
    if metadata.get("final_difficulty"):
        return str(metadata["final_difficulty"])
    raw = str(metadata.get("difficulty_level", ""))
    for level in ("L1", "L2", "L3", "L4"):
        if raw.startswith(level):
            return level
    return ""


def read_task(task_dir: Path) -> dict:
    task_toml = task_dir / "task.toml"
    data = tomllib.loads(task_toml.read_text())
    metadata = data.get("metadata", {})
    env = data.get("environment", {}).get("env", {})
    labels = metadata.get("atomic_operations", metadata.get("atomic_operation", ""))
    if isinstance(labels, list):
        labels_text = "; ".join(str(x) for x in labels)
    else:
        labels_text = str(labels)
    source = "v2_final_80" if str(metadata.get("final_dataset", "")) == "docops_v2_final_80" else "v1_final_130"
    return {
        "task_dir": task_dir.name,
        "harbor_task_name": data.get("task", {}).get("name", ""),
        "source": source,
        "task_id": metadata.get("task_id", ""),
        "level": norm_level(metadata),
        "doc_type": metadata.get("doc_type", ""),
        "benchmark_split": metadata.get("benchmark_split", ""),
        "operation_labels": labels_text,
        "input_path": env.get("INPUT_PATH", ""),
        "output_path": env.get("OUTPUT_PATH", ""),
    }


def main() -> None:
    rows = [read_task(path) for path in sorted(TASKS_DIR.iterdir()) if (path / "task.toml").is_file()]
    OUT_DIR.mkdir(exist_ok=True)
    fields = [
        "task_dir",
        "harbor_task_name",
        "source",
        "task_id",
        "level",
        "doc_type",
        "benchmark_split",
        "operation_labels",
        "input_path",
        "output_path",
    ]
    with (OUT_DIR / "tasks_manifest.csv").open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    (OUT_DIR / "tasks_manifest.json").write_text(json.dumps(rows, indent=2, ensure_ascii=False) + "\n")
    print(f"wrote {len(rows)} tasks to {OUT_DIR}")


if __name__ == "__main__":
    main()
