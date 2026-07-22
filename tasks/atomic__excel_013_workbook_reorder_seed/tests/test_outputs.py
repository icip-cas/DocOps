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

    assert wb_out.sheetnames == ['Instructions', 'Lookup', 'Detail_Q1', 'Detail_Q2', 'Summary'], f'Workbook tabs are in the wrong order: {wb_out.sheetnames}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
