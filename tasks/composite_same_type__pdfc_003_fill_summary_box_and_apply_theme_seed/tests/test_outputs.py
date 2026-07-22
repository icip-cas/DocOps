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
    page_texts = pdf_page_texts(OUTPUT_PATH)
    page1 = normalize_text(page_texts[0])
    page2_out = normalize_text(page_texts[1])
    page2_in = normalize_text(pdf_page_texts(INPUT_PATH)[1])
    assert 'summary box intentionally left blank' not in page1, 'Placeholder summary text should be removed.'
    require_ordered_anchors(page1, ['Supplier:', 'Inspection:', 'Staffing:'], 'Management Summary box')
    require_group_hits(page1, [
        ['supplier:'],
        ['backlog', 'lead time', 'release timing', 'releases lag', 'release lag'],
        ['inspection:'],
        ['re-inspection', 'reinspection', 'repeat inspection', 'repeat failures'],
        ['staffing:'],
        ['weekend', 'staffing', 'crew'],
    ], 'Management Summary box')
    assert page2_out == page2_in, 'Supporting notes page text should remain unchanged.'
    colors_out = pdf_theme_rect_colors(OUTPUT_PATH)
    blueish = [c for c in colors_out if len(c) == 6 and int(c[4:6], 16) >= int(c[0:2], 16) and int(c[4:6], 16) >= int(c[2:4], 16)]
    assert len(blueish) >= 3, 'Expected multiple blue-gray themed rectangles.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
