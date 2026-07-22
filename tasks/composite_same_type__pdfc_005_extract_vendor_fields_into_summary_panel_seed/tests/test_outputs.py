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
    page1 = normalize_text(pages[0])
    assert 'polar mechanical services ltd' in page1
    assert any(ref in page1 for ref in ['q-7741-r2', 'q 7741 r2'])
    assert '54,400' in pages[0]
    assert '28 may 2026' in page1
    assert normalize_text(pages[1]) == normalize_text(pdf_page_texts(INPUT_PATH)[1]), 'Detailed quote page should remain unchanged.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
