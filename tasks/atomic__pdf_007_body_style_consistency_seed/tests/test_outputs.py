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

    require_all(text_out, ['Operations Brief', 'Update A', 'Update B', 'Update C', 'Update D'], TASK_ID)
    with pdfplumber.open(OUTPUT_PATH) as pdf:
                page = pdf.pages[0]
                body_chars = [ch for ch in page.chars if ch.get('top', 0) > 90 and ch.get('top', 0) < 260]
                sigs = {(ch.get('fontname'), round(ch.get('size', 0), 1), str(ch.get('non_stroking_color'))) for ch in body_chars if ch.get('text', '').strip()}
                assert len(sigs) <= 3, f'Body paragraph styling is still too inconsistent: {sigs}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
