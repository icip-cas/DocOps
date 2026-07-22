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
    output_doc = Document(OUTPUT_PATH)
    table = output_doc.tables[0]
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    assert headers == ['Owner', 'Status', 'Due Date', 'Action'], f'Unexpected table headers: {headers}'
    first_row = [cell.text.strip() for cell in table.rows[1].cells]
    assert first_row[0] == 'Maya Chen' and first_row[2] == '14 Jun 2026' and 'vendor recovery sequence' in normalize_text(first_row[3]), 'Existing row content must be preserved around the new Status column.'
    texts = docx_texts(output_doc)
    note_candidates = [text for text in texts if text.startswith('Schedule risk:') or text.startswith('Client update:')]
    assert_prefixed_items(note_candidates, ['Schedule risk:', 'Client update:'], 'Draft notes')
    note1, note2 = note_candidates
    norm1 = normalize_text(note1)
    norm2 = normalize_text(note2)
    for forbidden in ['bit late', 'might slip', 'probably wait']:
        assert forbidden not in norm1 and forbidden not in norm2, f'Informal wording still present: {forbidden}'
    require_group_hits(norm1, [
        ['vendor'],
        ['one week', 'slip'],
        ['revised install date', 'revised installation date', 'updated install date'],
    ], 'Draft note 1')
    require_group_hits(norm2, [
        ['client'],
        ['revised install date', 'revised installation date', 'updated install date'],
        ['vendor'],
    ], 'Draft note 2')


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
