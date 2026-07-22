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
    out = pdf_page_texts(OUTPUT_PATH)[0]
    norm = normalize_text(out)
    assert 'missed two promised dates again' not in norm
    require_group_hits(norm, [
        ['supplier'],
        ['two promised dates', 'two committed delivery dates', 'two delivery dates'],
        ['team concern', 'team has expressed significant concern', 'team is frustrated', 'team is increasingly frustrated', 'increasingly frustrated', 'repeated delays'],
    ], 'Rewritten paragraph')
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        page = pdf.pages[0]
        line1 = line_text_words(page, 'This paragraph is the body-style reference')
        line2 = line_text_words(page, 'This final paragraph is another body-style reference')
        assert line1 and line2, 'Reference lines must remain present.'

def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
