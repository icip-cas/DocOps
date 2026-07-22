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
    assert wb.sheetnames == ['Overview', 'Detail_Q1', 'Detail_Q2', 'Risks', 'Executive Summary'], f'Unexpected sheet order: {wb.sheetnames}'
    ws = wb['Executive Summary']
    bullets = [str(ws[f'A{i}'].value or '').strip() for i in range(2, 5)]
    assert all(bullets), 'Executive Summary must contain exactly three bullets in A2:A4.'
    assert_prefixed_items(bullets, ['Supplier:', 'Inspection:', 'Staffing:'], 'Executive Summary sheet')
    require_group_hits(bullets[0], [['supplier'], ['delay', 'backlog', 'lead time']], 'Supplier bullet')
    require_group_hits(bullets[1], [['inspection'], ['hold', 'backlog', 'repeat hold']], 'Inspection bullet')
    require_group_hits(
        bullets[2],
        [['staffing', 'weekend', 'crew'], ['pressure', 'constraint', 'risk', 'coverage gap', 'staffing gap', 'shortfall']],
        'Staffing bullet',
    )
    text = normalize_text(' '.join(bullets))
    assert '99' not in text, 'Summary should not introduce unsupported numbers.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
