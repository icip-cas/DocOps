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
    wb_out, wb_out_values = load_xlsx_pair(OUTPUT_PATH)
    wb_in = load_workbook(INPUT_PATH)
    wb_in_values = load_workbook(INPUT_PATH, data_only=True)

    ws = wb_out['Program Outline']
    for row in [4, 7, 10]:
                assert int(ws.row_dimensions[row].outlineLevel or 0) == 0, f'Section row {row} should remain top-level.'
    for row in [5, 6, 8, 9, 11, 12]:
                assert int(ws.row_dimensions[row].outlineLevel or 0) >= 1, f'Detail row {row} was not grouped under its section.'
    assert workbook_values_signature(wb_out_values) == workbook_values_signature(wb_in_values), 'Outline task should not change the sheet values.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
