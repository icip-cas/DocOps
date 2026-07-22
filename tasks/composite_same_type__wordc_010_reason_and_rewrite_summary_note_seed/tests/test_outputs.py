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
    doc = Document(OUTPUT_PATH)
    input_texts = docx_texts(input_doc)
    output_texts = docx_texts(doc)
    assert len(output_texts) == len(input_texts), 'Summary note should be rewritten in place, not duplicated.'
    assert output_texts[:4] == input_texts[:4], 'Policy statements and heading should remain unchanged.'
    note = output_texts[4]
    norm = normalize_text(note)
    require_group_hits(norm, [
        ['cfo'],
        ['coo'],
        ['50,000', '$50,000', '50000'],
        ['exception', 'conflict', 'approval rule'],
    ], 'Summary note')
    assert 'mostly about timing and unclear routing' not in norm, 'Vague original note should be replaced.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
