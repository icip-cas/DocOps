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
    prs = Presentation(OUTPUT_PATH)
    tbl = prs.slides[0].shapes[1].table
    headers = [tbl.cell(0, c).text.strip() for c in range(4)]
    assert headers == ['Milestone', 'Owner', 'Due Date', 'Notes']
    owners = [tbl.cell(r, 1).text.strip() for r in [1,2,3]]
    assert owners == ['Maya Chen', 'Ravi Shah', 'Elena Park']

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
