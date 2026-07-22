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

    ws_in = wb_in['Status Draft']
    ws_out = wb_out['Status Draft']
    for row in range(1, ws_in.max_row + 1):
                for col in range(1, ws_in.max_column + 1):
                    coord = f'{get_column_letter(col)}{row}'
                    if coord == 'B8':
                        continue
                    assert ws_in[coord].value == ws_out[coord].value, f'Only B8 should change, but {coord} changed too.'
    new_text = str(ws_out['B8'].value or '')
    forbid_any(new_text, ['missed the deadline again', 'cannot trust the dates they gave us'], TASK_ID)
    require_any_group(new_text, [['timeline', 'schedule', 'dates'], ['uncertain', 'subject to', 'under review']], TASK_ID)



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
