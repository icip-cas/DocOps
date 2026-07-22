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

    require_regex(text, r'^note\s*:\s*1\s*$', TASK_ID)
    require_regex(text, r'^north\s*:\s*.*93.*95.*$', TASK_ID)
    require_regex(text, r'^east\s*:\s*.*83.*88.*$', TASK_ID)
    require_regex(text, r'^why\s*:\s*.*north.*not.*only.*east.*also.*below target.*$', TASK_ID)


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
