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

    ws = wb_out['Tracker']
    highlighted_rows = [5, 8]
    non_highlighted_rows = [3, 4, 6, 7]
    def row_has_visual_change(row):
                fills = [cell_fill_rgb(ws.cell(row, c)) for c in range(1, 7)]
                nondefault = [f for f in fills if f and f not in ('000000', 'FFFFFF', '00000000')]
                return len(nondefault) >= 3
    if not all(row_has_visual_change(r) for r in highlighted_rows):
                cf_ranges = ' '.join(all_cf_ranges(ws))
                assert any(str(r) in cf_ranges for r in highlighted_rows), 'High-risk rows are not visibly highlighted and no conditional-format ranges cover them.'
    for row in non_highlighted_rows:
                assert not row_has_visual_change(row), f'Row {row} should not be highlighted.'
    assert workbook_values_signature(wb_out_values) == workbook_values_signature(wb_in_values), 'Highlighting task should not change cell values.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
