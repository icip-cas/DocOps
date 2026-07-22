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

    require_any_group(text, [['clause 4.1']], TASK_ID)
    require_any_group(text, [['30 calendar days'], ['30 days']], TASK_ID)
    require_any_group(text, [['schedule b']], TASK_ID)
    require_any_group(text, [['45 calendar days'], ['45 days']], TASK_ID)
    require_any_group(text, [['day 46'], ['46']], TASK_ID)
    require_any_group(text, [['late payment fee'], ['late charge'], ['late-charge'], ['late fee'], ['late payment fee should apply']], TASK_ID)
    require_any_group(text, [['conflict'], ['inconsistent'], ['contradict']], TASK_ID)



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
