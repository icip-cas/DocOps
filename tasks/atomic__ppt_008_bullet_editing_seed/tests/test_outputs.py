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
    prs_out = Presentation(OUTPUT_PATH)
    slide = prs_out.slides[0]
    text_box = ppt_find_text_shape(slide, 'Draft bullets:')
    lines = [line.strip() for line in text_box.text.splitlines() if line.strip()]
    assert lines and lines[0] == 'Draft bullets:', 'Expected the draft bullet label to remain.'
    bullet_lines = [line.lstrip('- ').strip() for line in lines[1:] if line.strip()]
    assert len(bullet_lines) == 3, f'Expected three rewritten bullets, found {len(bullet_lines)}.'
    bullet_text = '\n'.join(bullet_lines)

    slide_text = '\n'.join(ppt_texts(prs_out))
    forbid_any(slide_text, ['dragging', 'losing patience', 'furious'], TASK_ID)
    require_any_group(bullet_text, [['paperwork', 'correction'], ['corrected files'], ['documentation', 'updated']], TASK_ID)
    require_any_group(bullet_text, [['today'], ['promptly'], ['timely']], TASK_ID)
    require_any_group(bullet_text, [['steering team'], ['leadership'], ['executive']], TASK_ID)


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
