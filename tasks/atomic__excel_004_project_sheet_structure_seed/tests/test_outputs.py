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

    for sheet_name in ['Project_A', 'Project_B']:
                ws = wb_out[sheet_name]
                headers = [ws.cell(3, c).value for c in range(1, 6)]
                assert headers == ['Task', 'Priority', 'Owner', 'Due Date', 'Status'], f'{sheet_name} headers are incorrect: {headers}'
                expected = [
                    ['Drain survey', None, 'Nina', '2026-04-22', 'Open'],
                    ['Valve list cleanup', None, 'Karl', '2026-04-25', 'Open'],
                    ['Shutdown note review', None, 'Ava', '2026-04-28', 'Draft'],
                ]
                actual = [[ws.cell(r, c).value for c in range(1, 6)] for r in range(4, 7)]
                assert actual == expected, f'{sheet_name} rows were not preserved after inserting Priority: {actual}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
