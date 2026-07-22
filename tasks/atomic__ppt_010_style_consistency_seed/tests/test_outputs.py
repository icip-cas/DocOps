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

    assert input_text == output_text, 'Style-consistency task should preserve slide text.'
    title_sigs = []
    body_sigs = []
    fills = []
    for slide in prs_out.slides:
                title_para = slide.shapes[0].text_frame.paragraphs[0]
                body_para = slide.shapes[1].text_frame.paragraphs[0]
                title_sigs.append(ppt_paragraph_signature(title_para))
                body_sigs.append(ppt_paragraph_signature(body_para))
                fills.append(shape_fill_rgb(slide.shapes[1]))
    assert len({json.dumps(sig, sort_keys=True, default=str) for sig in title_sigs}) == 1, 'Slide titles are still styled inconsistently.'
    assert len({json.dumps(sig, sort_keys=True, default=str) for sig in body_sigs}) == 1, 'Slide body text is still styled inconsistently.'
    assert len(set(fills)) == 1, 'Body-card appearance is still inconsistent across slides.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
