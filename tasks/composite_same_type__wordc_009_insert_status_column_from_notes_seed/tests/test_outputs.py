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
    headers = [cell.text.strip() for cell in doc.tables[0].rows[0].cells]
    assert headers == ['Owner', 'Status', 'Due Date', 'Action']
    vals = [doc.tables[0].rows[i].cells[1].text.strip() for i in [1,2,3]]
    assert vals == ['On track', 'Waiting on vendor', 'Drafting']

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
