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
    page_texts = pdf_page_texts(OUTPUT_PATH)
    assert normalize_text(page_texts[0]).startswith('scope'), 'Page 1 should be Scope.'
    assert normalize_text(page_texts[1]).startswith('findings'), 'Page 2 should be Findings.'
    assert normalize_text(page_texts[2]).startswith('recommendations'), 'Page 3 should be Recommendations.'
    assert normalize_text(page_texts[3]).startswith('appendix'), 'Page 4 should be Appendix.'
    outline = extract_pdf_outline_titles(OUTPUT_PATH)
    expected = [(0, 'Scope'), (0, 'Findings'), (1, 'Site A'), (1, 'Site B'), (0, 'Recommendations'), (0, 'Appendix')]
    assert outline == expected, f'Unexpected outline: {outline}'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
