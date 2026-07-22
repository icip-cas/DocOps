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

    assert len(prs_out.slides) == 2, 'Deck should still have exactly two slides.'
    slide2_text = '\n'.join(' '.join(getattr(shape, 'text', '').split()) for shape in prs_out.slides[1].shapes)
    assert 'TODO:' not in slide2_text, 'Risk-bullet placeholder is still present.'
    require_any_group(slide2_text, [['risk', 'delay', 'slip']], TASK_ID)
    require_any_group(slide2_text, [['variation order', 'paperwork', 'documentation']], TASK_ID)
    require_any_group(slide2_text, [['finance approval', 'recovery plan', 'leak risk']], TASK_ID)



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
