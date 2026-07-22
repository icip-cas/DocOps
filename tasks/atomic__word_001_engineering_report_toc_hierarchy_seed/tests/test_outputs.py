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

    assert output_texts.count('a. Isolation logic') == 1, 'Isolation logic subsection must restart at a in the body.'
    assert output_texts.count('a. Isolation logic	4') == 1, 'The manually typed TOC entry must also be updated to a.'
    assert not any('c. Isolation logic' in text for text in output_texts), 'Old incorrect subsection label still present.'

    section_b = 'B. Retrofit Scope' if 'B. Retrofit Scope' in output_texts else 'B.Retrofit Scope'
    top_level = ['A. Project Context', section_b, 'Appendix - Vendor Matrix']
    top_sigs = [docx_para_signature(docx_para_by_text(output_doc, t)) for t in top_level]
    assert len({json.dumps(sig, sort_keys=True) for sig in top_sigs}) == 1, 'Top-level headings are not styled consistently.'
    sub_sigs = [docx_para_signature(docx_para_by_text(output_doc, t)) for t in ['I. Site Conditions', 'I. Mechanical scope', 'II. Controls scope']]
    assert len({json.dumps(sig, sort_keys=True) for sig in sub_sigs}) == 1, 'Second-level headings are not styled consistently.'
    third_sigs = [docx_para_signature(docx_para_by_text(output_doc, t)) for t in ['1. Existing Boiler Room', '2. Temporary bypass plan']]
    assert len({json.dumps(sig, sort_keys=True) for sig in third_sigs}) == 1, 'Third-level headings are not styled consistently.'
    fourth_sigs = [docx_para_signature(docx_para_by_text(output_doc, t)) for t in ['a. Ventilation constraints', 'b. Access path', 'a. Isolation logic']]
    assert len({json.dumps(sig, sort_keys=True) for sig in fourth_sigs}) == 1, 'Fourth-level headings are not styled consistently.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
