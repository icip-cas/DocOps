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
    input_doc = Document(INPUT_PATH)
    output_doc = Document(OUTPUT_PATH)
    input_texts = docx_texts(input_doc)
    output_texts = docx_texts(output_doc)

    old_open = 'Hi team, just checking again because this thing is dragging and we still do not have the paperwork we were told would be ready last Thursday.'
    old_ask = 'Can you please get this sorted fast and send us the corrected files today, because operations is getting annoyed and we are running out of patience.'
    out_norm = normalize_text('\n'.join(output_texts))
    assert normalize_text(old_open) not in out_norm and normalize_text(old_ask) not in out_norm, 'Blunt draft text is still present.'
    require_any_group('\n'.join(output_texts), [['following up', 'documentation'], ['completed package', 'corrected files']], TASK_ID)
    assert input_texts[0:5] == output_texts[0:5], 'Unexpected edits occurred before the target opening paragraph.'
    assert input_texts[6:9] == output_texts[6:9], 'Unexpected edits occurred around the key-facts section.'
    assert input_texts[10:] == output_texts[10:], 'Unexpected edits occurred after the target ask paragraph.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
