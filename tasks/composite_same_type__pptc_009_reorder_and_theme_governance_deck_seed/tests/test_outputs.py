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
    titles = [slide_title_shape(s).text for s in prs.slides]
    assert titles == ['Overview', 'Risks', 'Actions', 'Appendix'], f'Unexpected slide order: {titles}'
    fills = []
    for slide in prs.slides:
        title_shape = slide_title_shape(slide)
        # title is textbox, inspect first non-title content if available
        for shape in list(slide.shapes)[1:]:
            fill = shape_fill_rgb(shape)
            if fill:
                fills.append(fill)
                break
    assert fills, 'Expected themed fills on the deck.'
    for fill in fills:
        assert_blueish(fill, 'deck theme fill')

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
