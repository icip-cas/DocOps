import json
import re
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
    assert 'Archive_OLD' not in wb.sheetnames, 'Archive_OLD sheet must be deleted.'
    raw_note = str(wb['Cover']['B3'].value or '').strip()
    lines = [line.strip() for line in raw_note.splitlines() if line.strip()]
    assert_prefixed_items(lines, ['Supplier:', 'Inspection:'], 'Cover note')
    require_group_hits(lines[0], [['supplier'], ['delay', 'late', 'slip']], 'Supplier note line')
    require_group_hits(lines[1], [['inspection'], ['backlog', 'hold', 'queue']], 'Inspection note line')

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
