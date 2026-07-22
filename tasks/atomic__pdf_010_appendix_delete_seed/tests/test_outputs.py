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

    assert len(reader_out.pages) == 2, 'Obsolete appendix page was not removed.'
    assert all('Superseded Rate Card' not in t for t in page_texts_out), 'Deleted page title still appears in output PDF.'
    assert normalize_text(page_texts_out[0]).startswith('audit packet'), 'Cover page should remain first.'
    assert normalize_text(page_texts_out[1]).startswith('findings'), 'Findings page should remain after deletion.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
