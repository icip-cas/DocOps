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

def row_has_fill(ws, row_num: int) -> bool:
    return any(cell_fill_rgb(ws.cell(row_num, c)) for c in range(1, 6))

def verify_task() -> None:
    wb = load_workbook(OUTPUT_PATH)
    assert str(wb['Summary']['B3'].value) in {'12500', '12500.0', '$12,500', '12500.00'}, 'Summary!B3 must contain the extracted approved amount.'
    assert row_has_fill(wb['Requests'], 3), 'CR-1184 row should be highlighted.'
    assert not row_has_fill(wb['Requests'], 2), 'Non-target row should not be highlighted.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
