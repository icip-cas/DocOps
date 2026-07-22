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
    return any(cell_fill_rgb(ws.cell(row_num, c)) for c in range(1, 5))

def verify_task() -> None:
    wb_formulas, wb_values = load_xlsx_pair(OUTPUT_PATH)
    ws = wb_formulas['Tracker']
    assert ws['D1'].value == 'Risk Flag', 'Risk Flag column must exist at D.'
    for row in [2, 3, 4]:
        assert ws[f'D{row}'].value == f'=IF(B{row}>=85,"Critical","Normal")'
    vals = [wb_values['Tracker'][f'D{r}'].value for r in [2,3,4]]
    assert vals == ['Critical', 'Normal', 'Critical'], f'Unexpected flag values: {vals}'
    assert row_has_fill(wb_formulas['Tracker'], 2) and row_has_fill(wb_formulas['Tracker'], 4), 'Critical rows must be highlighted.'
    assert not row_has_fill(wb_formulas['Tracker'], 3), 'Normal row should not be highlighted.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
