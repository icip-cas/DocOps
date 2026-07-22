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
    box = prs.slides[0].shapes[1]
    texts = [' '.join(p.text.split()) for p in box.text_frame.paragraphs if p.text.strip()]
    assert len(texts) == 3, f'Expected three rewritten bullets, found {len(texts)}.'
    combined = normalize_text(' '.join(texts))
    for bad in ['kinda', 'super', 'probably']:
        assert bad not in combined
    require_group_hits(combined, [
        ['vendor'],
        ['late', 'delay', 'behind schedule'],
        ['install date', 'installation date', 'revised install date', 'revised installation date', 'installation timeline', 'revised timeline'],
        ['fuzzy', 'uncertain', 'subject to change', 'pending confirmation', 'not yet finalized', 'not finalized', 'unconfirmed', 'to be confirmed'],
        ['client update', 'client communication'],
    ], 'Slide 2 bullets')
    sigs = [ppt_paragraph_signature(p) for p in box.text_frame.paragraphs if p.text.strip()]
    assert len({json.dumps(sig, sort_keys=True) for sig in sigs}) == 1, 'All rewritten bullets should share one style.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
