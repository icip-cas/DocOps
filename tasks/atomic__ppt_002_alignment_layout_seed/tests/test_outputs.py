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
    prs_in = Presentation(INPUT_PATH)
    prs_out = Presentation(OUTPUT_PATH)
    input_text = ppt_texts(prs_in)
    output_text = ppt_texts(prs_out)

    slide = prs_out.slides[0]
    boxes = [slide.shapes[1], slide.shapes[2], slide.shapes[3]]
    tops = [int(s.top) for s in boxes]
    lefts = [int(s.left) for s in boxes]
    assert max(tops) - min(tops) <= 200000, f'Workstream boxes are not horizontally aligned: {tops}'
    gaps = [lefts[i + 1] - (lefts[i] + int(boxes[i].width)) for i in range(2)]
    assert abs(gaps[0] - gaps[1]) <= 220000, f'Workstream box spacing is not balanced: {gaps}'
    assert input_text == output_text, 'Layout task should not change slide text.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
