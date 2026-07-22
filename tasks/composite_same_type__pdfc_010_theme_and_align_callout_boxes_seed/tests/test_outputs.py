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
    colors = pdf_theme_rect_colors(OUTPUT_PATH)
    assert len(colors) >= 3, 'Three themed callout boxes expected.'
    for c in colors[:3]:
        assert_blueish(c, 'callout theme color')
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        page = pdf.pages[0]
        rects = [r for r in page.rects if r.get('width',0) > 100 and r.get('height',0) > 40]
        assert len(rects) >= 3, 'Three callout rectangles expected.'
        tops = sorted(round(r.get('top',0),1) for r in rects[:3])
        assert tops[-1] - tops[0] <= 15, 'Callout boxes should align in one horizontal row.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
