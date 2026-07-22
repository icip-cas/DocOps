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
    page = normalize_text(pdf_page_texts(OUTPUT_PATH)[0])
    require_group_hits(page, [
        ['30 calendar days', '30 day clause'],
        ['45 calendar days', '45 day schedule'],
        ['day 46', 'late payment fee', 'late fee trigger'],
    ], 'Cover note')
    assert 'minor wording differences' not in page, 'Old vague cover note should be replaced.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
