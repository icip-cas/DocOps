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
    wb = load_workbook(OUTPUT_PATH)
    ws = wb['Plan']
    headers = [ws.cell(1, c).value for c in range(1, 4)]
    assert headers == ['Milestone', 'Owner', 'Due Date'], f'Unexpected headers: {headers}'
    assert ws['B2'].value == 'Maya Chen'
    assert ws['B3'].value == 'Ravi Shah'
    assert ws['B4'].value == 'Elena Park'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
