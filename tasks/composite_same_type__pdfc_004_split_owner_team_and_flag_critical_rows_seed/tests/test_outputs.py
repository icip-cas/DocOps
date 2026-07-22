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
    text = normalize_text(pdf_text(OUTPUT_PATH))
    assert 'owner / team' not in text, 'Combined header should be split.'
    assert 'owner' in text and 'team' in text, 'Separate Owner and Team headers must appear.'
    assert 'maya chen' in text and 'electrical' in text and 'ravi shah' in text and 'quality' in text, 'Row content must be preserved.'
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        page = pdf.pages[0]
        assert line_is_visually_highlighted(page, 'Valve kit release'), 'First critical row must be highlighted.'
        assert line_is_visually_highlighted(page, 'QA retest package'), 'Second critical row must be highlighted.'
        assert not line_is_visually_highlighted(page, 'Client update memo'), 'Watch row should not be highlighted.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
