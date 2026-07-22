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

    ws_in = wb_in['Hiring Tracker']
    ws_out = wb_out['Hiring Tracker']
    exemplar = row_style_signature(ws_in, 4, [1, 2, 3, 4])
    assert row_style_signature(ws_out, 4, [1, 2, 3, 4]) == exemplar, 'Row 4 should remain the body-style exemplar.'
    for row in range(4, 11):
        assert row_style_signature(ws_out, row, [1, 2, 3, 4]) == exemplar, (
            f'Body row {row} does not match the input exemplar style from row 4.'
        )
    for row in [1, 3, 13]:
        assert row_style_signature(ws_out, row, [1, 2, 3, 4]) == row_style_signature(ws_in, row, [1, 2, 3, 4]), (
            f'Non-target row {row} style changed unexpectedly.'
        )
    assert workbook_values_signature(wb_out_values) == workbook_values_signature(wb_in_values), (
        'Style consistency task should not change cell values.'
    )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
