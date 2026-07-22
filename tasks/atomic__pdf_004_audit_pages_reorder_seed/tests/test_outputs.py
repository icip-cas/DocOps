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
    reader_in = PdfReader(str(INPUT_PATH))
    reader_out = PdfReader(str(OUTPUT_PATH))
    text_out = pdf_text(OUTPUT_PATH)
    page_texts_out = pdf_page_texts(OUTPUT_PATH)

    order = [normalize_text(t).split()[0:2] for t in page_texts_out]
    assert normalize_text(page_texts_out[0]).startswith('1. scope'), 'Page 1 should be Scope.'
    assert normalize_text(page_texts_out[1]).startswith('2. findings'), 'Page 2 should be Findings.'
    assert normalize_text(page_texts_out[2]).startswith('3. recommendations'), 'Page 3 should be Recommendations.'
    assert normalize_text(page_texts_out[3]).startswith('4. appendix'), 'Page 4 should be Appendix.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
