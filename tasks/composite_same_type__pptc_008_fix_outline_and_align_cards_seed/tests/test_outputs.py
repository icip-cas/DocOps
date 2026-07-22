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
    prs = Presentation(OUTPUT_PATH)
    slide = prs.slides[0]
    text_shape = slide.shapes[1]
    levels = [p.level for p in text_shape.text_frame.paragraphs if p.text.strip()]
    assert levels == [0, 1, 0, 1], f'Unexpected paragraph levels: {levels}'
    cards = [slide.shapes[i] for i in [2,3,4]]
    ys = [card.top for card in cards]
    assert max(ys) - min(ys) <= 10000, 'Cards must be aligned into one row.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
