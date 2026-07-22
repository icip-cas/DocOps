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
    return any(cell_fill_rgb(ws.cell(row_num, c)) for c in range(1, 7))


def verify_task() -> None:
    wb_out = load_workbook(OUTPUT_PATH)
    wb_in = load_workbook(INPUT_PATH)
    note_out = normalize_text(str(wb_out['Ops Summary']['B8'].value))
    note_in = normalize_text(str(wb_in['Ops Summary']['B8'].value))
    assert note_out != note_in, 'Ops Summary!B8 must be rewritten.'
    require_group_hits(note_out, [
        ['north'],
        ['steepest backlog', 'backlog increase', 'backlog rise'],
        ['north alpha', '14 incidents', '12 incident threshold', 'incident threshold', 'exceeded 12 incidents', 'site exceeded 12 incidents'],
    ], 'Corrected note')
    assert 'west' not in note_out, 'Corrected note must no longer attribute the issue to West.'
    assert str(wb_out['Ops Summary']['B9'].value) == str(wb_in['Ops Summary']['B9'].value), 'Only the targeted note should change.'
    ws = wb_out['Monthly Data']
    assert row_has_fill(ws, 2), 'North Alpha row should be highlighted.'
    assert row_has_fill(ws, 5), 'North Delta row should be highlighted.'
    assert not row_has_fill(ws, 3), 'West row should not be highlighted.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
