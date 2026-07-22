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
    doc = Document(OUTPUT_PATH)
    table = doc.tables[0]
    assert table.cell(1,1).text.strip() == '28 June 2026'
    assert table.cell(2,1).text.strip() == 'Elena Park'
    p1 = docx_para_by_text(doc, 'Final approved shutdown date: 28 June 2026.')
    p2 = docx_para_by_text(doc, 'Final action owner for the vendor contact list: Elena Park.')
    assert docx_para_has_highlight(p1) and docx_para_has_highlight(p2), 'Evidence lines must be highlighted.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
