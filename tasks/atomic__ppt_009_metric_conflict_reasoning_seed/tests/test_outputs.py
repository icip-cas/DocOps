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

    require_any_group(text, [['5 days'], ['5-day']], TASK_ID)
    require_any_group(text, [['2 days'], ['2-day']], TASK_ID)
    require_any_group(text, [['rebaselined'], ['re-baselined']], TASK_ID)
    require_any_group(text, [['original baseline']], TASK_ID)
    require_any_group(text, [['different baselines'], ['not directly comparable'], ['slides do not fully agree'], ['inconsistent'], ['conflict']], TASK_ID)



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
