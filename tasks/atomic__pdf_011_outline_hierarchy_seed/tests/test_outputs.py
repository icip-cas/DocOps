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

    outline = extract_pdf_outline_titles(OUTPUT_PATH)
    expected = [
                (0, 'Report Overview'),
                (0, 'Findings'),
                (1, 'Site A'),
                (1, 'Site B'),
                (0, 'Recommendations'),
            ]
    assert outline == expected, f'Unexpected bookmark hierarchy: {outline}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
