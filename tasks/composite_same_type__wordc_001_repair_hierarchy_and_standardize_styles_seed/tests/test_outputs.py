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
    output_doc = Document(OUTPUT_PATH)
    output_texts = docx_texts(output_doc)
    assert 'a. Isolation logic' in output_texts, 'Isolation logic subsection must restart at a.'
    assert 'c. Isolation logic' not in output_texts, 'Old incorrect subsection label still present.'
    top_sigs = [docx_run_signature(docx_para_by_text(output_doc, t)) for t in ['A. Project Context', 'B. Retrofit Scope', 'Appendix - Vendor Matrix']]
    assert len({json.dumps(sig, sort_keys=True) for sig in top_sigs}) == 1, 'Top-level headings must share one style.'
    second_sigs = [docx_run_signature(docx_para_by_text(output_doc, t)) for t in ['I. Site Conditions', 'I. Mechanical scope', 'II. Controls scope']]
    assert len({json.dumps(sig, sort_keys=True) for sig in second_sigs}) == 1, 'Second-level headings must share one style.'
    body_sigs = [docx_run_signature(docx_para_by_text(output_doc, t)) for t in [
        'Existing boiler room paragraph with uneven spacing and formatting.',
        'Ventilation notes use one body format.',
        'Access path notes use another body format.',
        'Mechanical scope body paragraph.',
    ]]
    assert len({json.dumps(sig, sort_keys=True) for sig in body_sigs}) == 1, 'Body paragraphs must share one style.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
