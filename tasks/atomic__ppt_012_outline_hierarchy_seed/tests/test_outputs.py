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
    prs_out = Presentation(OUTPUT_PATH)
    slide = prs_out.slides[0]
    text_box = ppt_find_text_shape(slide, 'Program Controls')
    paras = [p for p in text_box.text_frame.paragraphs if ''.join(r.text for r in p.runs).strip()]
    levels = {''.join(r.text for r in p.runs).strip(): p.level for p in paras}
    assert levels.get('Program Controls') == 0, 'Program Controls should remain top-level.'
    assert levels.get('Mechanical Tasks') == 0, 'Mechanical Tasks should remain top-level.'
    assert levels.get('Electrical Tasks') == 0, 'Electrical Tasks should remain top-level.'
    for item in ['Site Readiness', 'confirm vendor contact list', 'Drain reroute', 'panel cleanup']:
        assert levels.get(item, 0) >= 1, f'{item} should be nested under a parent workstream.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
