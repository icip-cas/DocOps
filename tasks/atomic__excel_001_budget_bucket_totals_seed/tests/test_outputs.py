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


def code_map(ws_map):
    out = {}
    for row in range(2, ws_map.max_row + 1):
        code = ws_map[f'A{row}'].value
        bucket = ws_map[f'B{row}'].value
        if code and bucket:
            out[str(code).strip()] = str(bucket).strip()
    return out


def expected_totals(ws, mapping):
    totals = {'Maintenance': 0.0, 'Operations': 0.0}
    for row in range(4, 15):
        code = ws[f'C{row}'].value
        amount = ws[f'D{row}'].value
        bucket = mapping.get(str(code).strip())
        if bucket in totals and isinstance(amount, (int, float)):
            totals[bucket] += float(amount)
    return totals


def formula_matches_bucket(ws, mapping, cell_ref: str, bucket: str) -> bool:
    formula = ws[cell_ref].value
    assert isinstance(formula, str) and formula.startswith('='), f'{cell_ref} must contain a formula.'
    text = normalize_text(formula)
    compact = text.replace(' ', '').replace('$', '')
    bucket_l = bucket.lower()
    prefix = 'm-*' if bucket == 'Maintenance' else 'o-*'
    label_ref = 'h3' if cell_ref == 'I3' else 'h4'

    # Vectorized VLOOKUP inside SUMPRODUCT is rejected here because it is not a
    # reliably recalculable spreadsheet formula across engines for this task.
    if 'vlookup(' in compact and 'c4:c14' in compact and 'sumproduct' in compact:
        return False

    # Direct code-prefix SUMIF, e.g. =SUMIF(C4:C14,"M-*",D4:D14).
    if 'sumif' in compact and prefix in compact and 'c4:c14' in compact and 'd4:d14' in compact:
        return True

    # Bucket-column SUMIF, optionally with helper formulas in F4:F14.
    if 'sumif' in compact and bucket_l in compact and ('f:f' in compact or 'f4:f14' in compact):
        return True

    # Code-map driven formulas such as SUMPRODUCT(SUMIF(...Code Map...)*(...="Maintenance"))
    # or SUMPRODUCT((VLOOKUP(...Code Map...)="Maintenance")*amounts).
    if 'sumproduct' in compact and bucket_l in compact and 'codemap' in compact and 'd4:d14' in compact:
        return True

    # INDEX/MATCH variants over the Code Map are also acceptable dynamic formulas.
    if 'sumproduct' in compact and bucket_l in compact and 'match' in compact and 'index' in compact:
        return True

    if 'sumif' in compact and bucket_l in compact and 'codemap' in compact and 'd4:d14' in compact:
        return True

    if 'sumif' in compact and 'codemap' in compact and 'd4:d14' in compact and f'substitute({label_ref}' in compact:
        return True

    return False


def assert_formula_total(ws, mapping, cell_ref: str, bucket: str, expected: float) -> None:
    assert formula_matches_bucket(ws, mapping, cell_ref, bucket), (
        f'{cell_ref}: formula does not compute the {bucket} total from the cost-code mapping.'
    )
    actual = expected_totals(ws, mapping)[bucket]
    assert abs(actual - expected) < 0.01, (
        f'{bucket} total: expected source-data total {expected:.2f}, found {actual:.2f}'
    )


def verify_task() -> None:
    wb_out, wb_out_values = load_xlsx_pair(OUTPUT_PATH)
    wb_in = load_workbook(INPUT_PATH)
    wb_in_values = load_workbook(INPUT_PATH, data_only=True)

    ws = wb_out['Cost Tracker']
    assert isinstance(ws['I3'].value, str) and ws['I3'].value.startswith('='), 'Maintenance total cell I3 must contain a formula.'
    assert isinstance(ws['I4'].value, str) and ws['I4'].value.startswith('='), 'Operations total cell I4 must contain a formula.'
    assert isinstance(ws['I5'].value, str) and ws['I5'].value.startswith('='), 'Grand total cell I5 must contain a formula.'
    f_i3 = normalize_text(ws['I3'].value)
    f_i4 = normalize_text(ws['I4'].value)
    f_i5 = normalize_text(ws['I5'].value)
    assert 'maintenance' in f_i3 or 'sumif' in f_i3 or 'sumproduct' in f_i3, 'Maintenance formula does not look semantically correct.'
    assert 'operations' in f_i4 or 'sumif' in f_i4 or 'sumproduct' in f_i4, 'Operations formula does not look semantically correct.'
    assert 'sum' in f_i5 or 'i3' in f_i5 or 'i4' in f_i5, 'Grand total formula does not look semantically correct.'
    buckets = [ws[f'F{row}'].value for row in range(4, 15)]
    if all(buckets):
                assert buckets == ['Maintenance', 'Operations', 'Maintenance', 'Operations', 'Maintenance', 'Operations', 'Maintenance', 'Operations', 'Maintenance', 'Operations', 'Maintenance']

    mapping = code_map(wb_out['Code Map'])
    assert_formula_total(ws, mapping, 'I3', 'Maintenance', 9125.00)
    assert_formula_total(ws, mapping, 'I4', 'Operations', 3385.00)


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
