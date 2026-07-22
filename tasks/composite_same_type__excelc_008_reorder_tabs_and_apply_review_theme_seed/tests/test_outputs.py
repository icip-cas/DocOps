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
    return None if color is None or color.rgb is None else str(color.rgb)[-6:]

def verify_task() -> None:
    wb = load_workbook(OUTPUT_PATH)
    assert wb.sheetnames == ['Cover', 'Metrics', 'Risks', 'Notes'], f'Unexpected sheet order: {wb.sheetnames}'
    for ws_name in wb.sheetnames:
        fill = cell_fill_rgb(wb[ws_name]['A1'])
        assert fill is not None, f'{ws_name} title band must be colored.'
        assert_blueish(fill[-6:], f'{ws_name} title band')
        assert_blueish(tab_rgb(wb[ws_name]), f'{ws_name} tab color')

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
