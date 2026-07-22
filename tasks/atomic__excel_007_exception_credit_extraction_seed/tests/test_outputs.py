import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from verifier_utils import *  # noqa: F401,F403

META = json.loads(Path('/tests/task_metadata.json').read_text(encoding='utf-8'))
INPUT_PATH = Path(META['input_path'])
OUTPUT_PATH = Path(META['output_path'])
TASK_ID = META['task_id']
DOC_TYPE = META['doc_type']


def verify_task() -> None:
    text = answer_text(OUTPUT_PATH)
    assert len(text) >= 10, 'Final answer is too short.'

    require_regex(text, r'^request id\s*:\s*cr-1184\s*$', TASK_ID)
    require_regex(text, r'^condition\s*:\s*.*after\s+(apr(?:il)?\.?\s*18|18\s+apr(?:il)?)\s*$', TASK_ID)
    require_regex(text, r'^final approved credit\s*:\s*\$?\s*6,950\s*$', TASK_ID)
    assert_single_expected_money(text, '6,950', TASK_ID)


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
