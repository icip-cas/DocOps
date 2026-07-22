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
    assert len(pages) == 2, 'Superseded page should be deleted.'
    text = normalize_text('\n'.join(pages))
    assert 'superseded rate card' not in text
    colors = pdf_theme_rect_colors(OUTPUT_PATH)
    assert len(colors) >= 2
    for c in colors[:2]:
        assert_blueish(c, 'theme rect')

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
