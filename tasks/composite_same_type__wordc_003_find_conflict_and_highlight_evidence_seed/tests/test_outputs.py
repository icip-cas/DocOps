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
    para1 = docx_para_by_text(output_doc, 'Standard purchases above $50,000 require CFO approval before release.')
    para2 = docx_para_by_text(output_doc, 'Emergency purchases above $50,000 may proceed with COO approval only and do not require CFO approval.')
    para3 = docx_para_by_text(output_doc, 'All emergency requests must still be logged in the central tracker within one business day.')
    assert docx_para_has_highlight(para1), 'First conflicting statement must be highlighted.'
    assert docx_para_has_highlight(para2), 'Second conflicting statement must be highlighted.'
    assert not docx_para_has_highlight(para3), 'Non-conflicting statement should not be highlighted.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
