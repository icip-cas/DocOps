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

def shape_text(shape):
    return ' '.join(getattr(shape, 'text', '').split())


def verify_task() -> None:
    prs_in = Presentation(INPUT_PATH)
    prs_out = Presentation(OUTPUT_PATH)
    assert ppt_texts(prs_in) == ppt_texts(prs_out), 'Theme/style/alignment task should preserve slide text.'
    title_sigs = [ppt_paragraph_signature(slide_title_shape(slide).text_frame.paragraphs[0]) for slide in prs_out.slides]
    assert len({json.dumps(sig, sort_keys=True) for sig in title_sigs}) == 1, 'Slide title styles must be consistent.'
    slide2 = prs_out.slides[1]
    cards = [slide2.shapes[i] for i in [1, 2, 3]]
    fills = [shape_fill_rgb(card) for card in cards]
    assert all(fills), 'All workstream cards must have fill colors.'
    for fill in fills:
        assert_blueish(fill, 'workstream card fill')
    assert len(set(fills)) == 1, 'Workstream cards must share one theme color.'
    ys = [card.top for card in cards]
    assert max(ys) - min(ys) <= 10000, 'Cards on Slide 2 must align horizontally.'
    xs = [card.left for card in cards]
    gaps = [xs[1] - xs[0], xs[2] - xs[1]]
    assert abs(gaps[0] - gaps[1]) <= 12000, 'Cards on Slide 2 must have even spacing.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
