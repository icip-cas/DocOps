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
    pages = pdf_page_texts(OUTPUT_PATH)
    assert normalize_text(pages[0]).startswith('scope')
    assert normalize_text(pages[1]).startswith('findings')
    assert normalize_text(pages[2]).startswith('recommendations')
    assert normalize_text(pages[3]).startswith('appendix')
    cover = normalize_text(pages[0])
    require_ordered_anchors(cover, ['Permit:', 'Supplier:', 'Staffing:'], 'Cover summary box')
    require_group_hits(cover, [
        ['permit:'],
        ['delay', 'delayed'],
        ['supplier:'],
        ['backlog', 'lead time'],
        ['staffing:'],
        ['weekend', 'staffing', 'crew'],
    ], 'Cover summary box')

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
