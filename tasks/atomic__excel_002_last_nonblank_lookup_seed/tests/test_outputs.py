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

    ws = wb_out['Selector']
    assert isinstance(ws['B4'].value, str) and ws['B4'].value.startswith('='), 'Selector!B4 must contain a formula.'
    formula_norm = normalize_text(ws['B4'].value)
    formula_refs_norm = formula_norm.replace('$', '')
    assert 'monthly data' in formula_norm or 'xlookup' in formula_norm or 'lookup' in formula_norm or 'filter' in formula_norm, 'Formula does not appear to read from Monthly Data.'
    assert 'b1' in formula_refs_norm or '2025' in formula_norm, 'Formula should use the selected year.'
    assert 'b2' in formula_refs_norm or 'gamma' in formula_norm, 'Formula should use the selected product.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
