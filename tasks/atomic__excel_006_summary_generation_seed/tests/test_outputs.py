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

    assert 'Executive Summary' in wb_out.sheetnames, 'Missing required Executive Summary sheet.'
    ws = wb_out['Executive Summary']
    entries = []
    for row in range(1, min(ws.max_row, 10) + 1):
                value = ws.cell(row, 1).value
                if isinstance(value, str) and value.strip():
                    entries.append(value.strip())
    assert len(entries) >= 3, f'Expected at least three generated bullets, found {entries}'
    joined = ' '.join(entries)
    require_any_group(joined, [['schedule', 'shutdown', 'slip']], TASK_ID)
    require_any_group(joined, [['paperwork', 'drawing', 'document']], TASK_ID)
    require_any_group(joined, [['vendor', 'engineer'], ['contact list', 'communications', 'comms']], TASK_ID)



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
