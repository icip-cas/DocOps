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

    def run_is_emphasized(run) -> bool:
        font = run.font
        color = None
        try:
            if font.color and font.color.type is not None:
                color = str(font.color.rgb)
        except Exception:
            color = None
        return bool(font.bold) or (color is not None and color != '141414')

    slide = prs_out.slides[0]
    found = False
    for shape in slide.shapes:
                if not getattr(shape, 'has_text_frame', False):
                    continue
                for para in shape.text_frame.paragraphs:
                    text = ''.join(run.text for run in para.runs).strip()
                    if 'commissioning engineer unavailable' in normalize_text(text):
                        sig = ppt_paragraph_signature(para)
                        if (
                            sig['bold']
                            or (sig['color'] and sig['color'] != '141414')
                            or any(run.text.strip() and run_is_emphasized(run) for run in para.runs)
                        ):
                            found = True
    assert found, 'Critical milestone item is not visually emphasized.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
