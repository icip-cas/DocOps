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
    texts = docx_texts(doc)
    headings = [t for t in texts if t in ['Executive Summary','Findings','Next Steps','Appendix']]
    assert headings[:4] == ['Executive Summary','Findings','Next Steps','Appendix'], f'Unexpected order: {headings}'
    sigs = [docx_run_signature(docx_para_by_text(doc, t)) for t in headings[:4]]
    colors = [sig['color'] for sig in sigs]
    assert len(set(colors)) == 1 and colors[0] is not None, 'Top-level headings should share one themed color.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
