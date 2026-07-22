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
    assert titles[0] == 'Executive Summary', f'Expected inserted first slide, got {titles}'
    box = prs.slides[0].shapes[1]
    texts = [' '.join(p.text.split()) for p in box.text_frame.paragraphs if p.text.strip()]
    assert_prefixed_items(texts, ['Supplier:', 'Inspection:', 'Staffing:'], 'Inserted Executive Summary slide')
    require_group_hits(texts[0], [['supplier'], ['delay', 'lead time', 'slipped']], 'Supplier bullet')
    require_group_hits(texts[1], [['inspection'], ['backlog', 'repeat holds']], 'Inspection bullet')
    require_group_hits(texts[2], [['staffing', 'crew', 'weekend'], ['risk', 'constraint', 'capacity', 'recovery', 'support']], 'Staffing bullet')

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
