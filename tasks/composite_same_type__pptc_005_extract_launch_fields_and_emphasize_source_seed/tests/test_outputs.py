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

def find_text_shape(slide, prefix):
    for shape in slide.shapes:
        if getattr(shape, 'has_text_frame', False) and ' '.join(shape.text.split()).startswith(prefix):
            return shape
    raise AssertionError(f'Shape not found for {prefix}')

def verify_task() -> None:
    prs = Presentation(OUTPUT_PATH)
    s1, s2 = prs.slides[0], prs.slides[1]
    assert '25 July' in find_text_shape(s1, 'Launch date').text
    assert 'Elena Park' in find_text_shape(s1, 'Owner').text
    target = find_text_shape(s2, 'Final launch date confirmed as 25 July.')
    para = target.text_frame.paragraphs[0]
    sig = ppt_paragraph_signature(para)
    assert sig['bold'] or (sig['color'] and sig['color'] != '000000'), 'Source line must be emphasized.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
