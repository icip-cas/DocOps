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

    assert input_texts == output_texts, 'Theme transfer should keep document text unchanged.'
    for heading in ['Quarterly Review', 'Budget Snapshot', 'Action Items']:
                sig = docx_run_signature(docx_para_by_text(output_doc, heading))
                assert sig['color'] is not None, f'Heading {heading} does not have an explicit theme color.'
                assert_blueish(sig['color'], f'Heading {heading}')
    heading_sigs = [docx_run_signature(docx_para_by_text(output_doc, t)) for t in ['Quarterly Review', 'Budget Snapshot', 'Action Items']]
    assert len({json.dumps(sig, sort_keys=True) for sig in heading_sigs}) == 1, 'Headings do not share a unified professional theme.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
