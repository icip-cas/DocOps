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
    wb = load_workbook(OUTPUT_PATH)
    ws = wb['Status Draft']
    text = normalize_text(str(ws['B8'].value))
    for bad in ['kinda', 'slip a bit']:
        assert bad not in text, f'Informal wording still present: {bad}'
    require_group_hits(text, [
        ['schedule', 'timeline'],
        ['vendor'],
        ['uncertain', 'uncertainty', 'risk', 'subject to change', 'subject to confirmation', 'may shift', 'may be revised', 'pending confirmation'],
    ], 'Edited sentence')
    sigs = [style_signature(ws[c]) for c in ['B7', 'B8', 'B9', 'B10']]
    assert sigs[1] == sigs[0] == sigs[2] == sigs[3], 'B8:B10 should match the body style of B7.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
