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

    in_text = pdf_text(INPUT_PATH)
    assert normalize_text(in_text) == normalize_text(text_out), 'Layout-control task should not change the PDF text.'
    with pdfplumber.open(OUTPUT_PATH) as pdf:
                rects = [r for r in pdf.pages[0].rects if r.get('width', 0) > 120 and r.get('height', 0) > 50]
                rects = sorted(rects, key=lambda r: r['x0'])[:6]
                fill_rects = [r for r in rects if r.get('fill')]
                assert len(fill_rects) >= 3, 'Expected three milestone boxes on the page.'
                selected = sorted(fill_rects[:3], key=lambda r: r['x0'])
                tops = [r['top'] for r in selected]
                assert max(tops) - min(tops) <= 20, f'Milestone boxes are not aligned into one horizontal row: {tops}'
                gaps = [selected[i + 1]['x0'] - selected[i]['x1'] for i in range(2)]
                assert abs(gaps[0] - gaps[1]) <= 25, f'Milestone box spacing is not even: {gaps}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
