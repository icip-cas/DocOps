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
    input_doc = Document(INPUT_PATH)
    output_doc = Document(OUTPUT_PATH)
    input_texts = docx_texts(input_doc)
    output_texts = docx_texts(output_doc)

    criticals = [
                'Critical: Supplier may suspend service after two unpaid invoices.',
                'Critical: Customer data may not be transferred outside the named region without approval.',
            ]
    noncriticals = [
                'Clause 2: Standard delivery window is ten business days.',
                'Clause 4: Liquidated damages apply only after formal written notice.',
                'Clause 6: Courtesy support calls are not billable.',
                'Clause 8: Annual review meeting is optional.',
            ]
    for text in criticals:
                assert docx_para_has_highlight(docx_para_by_text(output_doc, text)), f'Critical line not highlighted: {text}'
    for text in noncriticals:
                assert not docx_para_has_highlight(docx_para_by_text(output_doc, text)), f'Non-critical line should not be highlighted: {text}'
    assert input_texts == output_texts, 'Highlighting task should not rewrite the text.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
