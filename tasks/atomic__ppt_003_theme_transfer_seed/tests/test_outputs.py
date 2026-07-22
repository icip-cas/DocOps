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

    assert input_text == output_text, 'Theme transfer should preserve slide text.'
    card_fills = []
    for slide in prs_out.slides:
                card = slide.shapes[1]
                fill = shape_fill_rgb(card)
                assert fill is not None, 'Each content card should have a fill color.'
                assert_blueish(fill, 'PPT card fill')
                card_fills.append(fill)
    assert len(set(card_fills)) == 1, 'All content cards should use one cohesive themed fill.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
