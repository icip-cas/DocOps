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

    require_regex(text, r'^statement a\s*:\s*.*before finance approval.*temporary waiver.*$', TASK_ID)
    require_regex(text, r'^statement b\s*:\s*.*temporary waiver.*removed from the final process.*override finance approval.*$', TASK_ID)
    require_regex(text, r'^final rule\s*:\s*.*finance approval.*(required|recorded|must still be recorded).*$', TASK_ID)
    require_regex(
        text,
        r'^why\s*:\s*.*((conflict|contradict|cannot all be true|inconsistent)|cancel|remove|no longer override).*$',
        TASK_ID,
    )
    forbid_any(text, ['temporary waiver still overrides finance approval', 'temporary waiver can override finance approval'], TASK_ID)


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
