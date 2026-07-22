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
    assert normalize_text(pdf_text(INPUT_PATH)) == normalize_text(pdf_text(OUTPUT_PATH)), 'Highlighting task must preserve text.'
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        page = pdf.pages[0]
        assert line_is_visually_highlighted(page, 'Clause 4.1 - Payment is due within 30 calendar days of invoice.'), '30-day clause must be highlighted.'
        assert line_is_visually_highlighted(page, 'Schedule B - Payment target: 45 calendar days from invoice receipt.'), '45-day schedule line must be highlighted.'
        assert line_is_visually_highlighted(page, 'Late payment fee applies from day 46.'), 'Late-fee trigger line must be highlighted.'
        assert not line_is_visually_highlighted(page, 'Operational note: site mobilization starts after the first invoice clears.'), 'Operational note should not be highlighted.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
