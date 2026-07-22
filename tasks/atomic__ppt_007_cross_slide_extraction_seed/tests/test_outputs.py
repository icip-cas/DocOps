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

    require_any_group(text, [['22 July'], ['2026-07-22']], TASK_ID)
    require_all(text, ['Ava Singh'], TASK_ID)
    require_any_group(text, [['launch date'], ['launch target'], ['launch readiness']], TASK_ID)
    forbid_any(text, ['15 July', 'Karl Meyer'], TASK_ID)



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
