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

    assert normalize_text(pdf_text(INPUT_PATH)) == normalize_text(text_out), 'Theme-transfer task should not change the document text.'
    colors = pdf_theme_rect_colors(OUTPUT_PATH)
    assert colors, 'No colored theme rectangles found in output PDF.'
    blueish = [c for c in colors if len(c) == 6 and int(c[4:6], 16) >= int(c[0:2], 16) and int(c[4:6], 16) >= int(c[2:4], 16)]
    assert len(blueish) >= 4, f'Expected multiple blue-dominant themed elements, found: {colors[:10]}'
    assert len(set(blueish[:4])) <= 2, 'Theme colors are still too inconsistent across pages/elements.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
