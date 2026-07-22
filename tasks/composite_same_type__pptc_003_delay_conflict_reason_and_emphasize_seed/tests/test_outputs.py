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

def find_paragraph(prs, target: str):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, 'has_text_frame', False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                if ' '.join(paragraph.text.split()) == target:
                    return paragraph
    raise AssertionError(f'Paragraph not found: {target}')


def verify_task() -> None:
    prs_in = Presentation(INPUT_PATH)
    prs_out = Presentation(OUTPUT_PATH)
    target = 'Critical: Launch remains on track for 18 July.'
    other = 'Training materials are still in review.'
    sig_in = ppt_paragraph_signature(find_paragraph(prs_in, target))
    sig_out = ppt_paragraph_signature(find_paragraph(prs_out, target))
    assert sig_out != sig_in, 'Conflicting line must be visually changed.'
    assert sig_out['bold'] or (sig_out['color'] and sig_out['color'] != sig_in['color']), 'Conflicting line must be emphasized with bold or color.'
    assert ppt_paragraph_signature(find_paragraph(prs_out, other)) == ppt_paragraph_signature(find_paragraph(prs_in, other)), 'Non-target bullet should remain unchanged.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
