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
    doc = Document(OUTPUT_PATH)
    opening = docx_texts(doc)[1]
    ask = docx_texts(doc)[3]
    norm = normalize_text(opening + ' ' + ask)
    for bad in ['hey team', 'kind of', 'just approve']:
        assert bad not in norm
    require_group_hits(norm, [
        ['vendor'],
        ['approve', 'approval', 'authorize', 'authorization'],
        ['extra shift', 'additional shift', 'shift coverage'],
        ['catch up', 'recover lost time', 'recover the schedule', 'make up lost time'],
    ], 'Rewritten letter')
    sigs = [docx_run_signature(doc.paragraphs[1]), docx_run_signature(doc.paragraphs[3]), docx_run_signature(doc.paragraphs[5])]
    assert len({json.dumps(sig, sort_keys=True) for sig in sigs}) == 1, 'Rewritten paragraphs should match the reference body style.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
