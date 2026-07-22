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

def tab_rgb(ws):
    color = ws.sheet_properties.tabColor
    if color is None or color.rgb is None:
        return None
    rgb = str(color.rgb)
    return rgb[-6:]


def verify_task() -> None:
    wb_formulas, wb_values = load_xlsx_pair(OUTPUT_PATH)
    tracker = wb_formulas['Tracker']
    assert tracker['D1'].value == 'Priority Score', 'Priority Score column must be inserted at column D.'
    expected = [
        '=VLOOKUP(B2,Weights!$A$2:$B$5,2,FALSE)*IF(C2>0,2,1)',
        '=VLOOKUP(B3,Weights!$A$2:$B$5,2,FALSE)*IF(C3>0,2,1)',
        '=VLOOKUP(B4,Weights!$A$2:$B$5,2,FALSE)*IF(C4>0,2,1)',
        '=VLOOKUP(B5,Weights!$A$2:$B$5,2,FALSE)*IF(C5>0,2,1)',
    ]
    for idx, formula in enumerate(expected, start=2):
        assert tracker[f'D{idx}'].value == formula, f'Unexpected formula in D{idx}'
    values = [wb_values['Tracker'][f'D{idx}'].value for idx in range(2, 6)]
    assert values == [10, 4, 6, 2], f'Priority scores are incorrect: {values}'
    for ws_name in ['Tracker', 'Dashboard', 'Weights']:
        ws = wb_formulas[ws_name]
        fill = cell_fill_rgb(ws['A1'])
        assert fill is not None, f'{ws_name} header must be themed.'
        assert_blueish(fill[-6:], f'{ws_name} header')
        tab = tab_rgb(ws)
        assert tab is not None, f'{ws_name} tab color must be set.'
        assert_blueish(tab, f'{ws_name} tab')


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
