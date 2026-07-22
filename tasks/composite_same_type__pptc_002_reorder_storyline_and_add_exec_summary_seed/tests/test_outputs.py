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

def slide_titles(prs):
    return [slide_title_shape(slide).text for slide in prs.slides]


def verify_task() -> None:
    prs = Presentation(OUTPUT_PATH)
    assert slide_titles(prs) == ['Executive Summary', 'Root Cause', 'Implementation Plan', 'Appendix'], f'Unexpected slide order: {slide_titles(prs)}'
    slide = prs.slides[0]
    bullet_box = slide.shapes[1]
    texts = [' '.join(p.text.split()) for p in bullet_box.text_frame.paragraphs if p.text.strip()]
    assert len(texts) == 3, f'Executive Summary must contain exactly three bullets, found {len(texts)}'
    assert_prefixed_items(texts, ['Vendor:', 'Inspection:', 'Staffing:'], 'Executive Summary slide')
    require_group_hits(texts[0], [['vendor'], ['lead time', 'delay', 'slip']], 'Vendor bullet')
    require_group_hits(texts[1], [['inspection'], ['backlog', 'repeat hold']], 'Inspection bullet')
    require_group_hits(texts[2], [['staffing', 'crew'], ['risk', 'constraint', 'weekend']], 'Staffing bullet')


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
