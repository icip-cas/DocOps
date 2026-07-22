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
    doc = Document(OUTPUT_PATH)
    assert doc.paragraphs, 'Output document must contain visible resume content.'
    assert doc.paragraphs[0].text.strip() == 'Jordan Hale', 'Blank opening page should be removed so the name appears first.'
    assert not docx_has_page_break(doc.paragraphs[0]), 'First paragraph should not retain the original page break.'
    texts = docx_texts(doc)
    assert texts[0] == 'Jordan Hale', 'Name should remain first visible text.'
    assert len(texts) == 3, f'Expected exactly two generated profile lines under the name, found {len(texts) - 1}.'
    assert_prefixed_items(texts[1:3], ['Commissioning:', 'Vendor coordination:'], 'Candidate profile')
    profile = normalize_text(' '.join(texts[1:3]))
    assert 'commissioning' in profile and 'vendor' in profile, 'Profile must mention commissioning leadership and vendor coordination.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
