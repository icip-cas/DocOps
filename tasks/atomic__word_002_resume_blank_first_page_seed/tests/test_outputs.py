import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from verifier_utils import *  # noqa: F401,F403

META_PATH = Path(os.environ.get('TASK_METADATA_PATH', '/tests/task_metadata.json'))
META = json.loads(META_PATH.read_text(encoding='utf-8'))
INPUT_PATH = Path(os.environ.get('INPUT_PATH', META['input_path']))
OUTPUT_PATH = Path(os.environ.get('OUTPUT_PATH', META['output_path']))
TASK_ID = META['task_id']
DOC_TYPE = META['doc_type']


def verify_task() -> None:
    input_doc = Document(INPUT_PATH)
    output_doc = Document(OUTPUT_PATH)
    input_texts = docx_texts(input_doc)
    output_texts = docx_texts(output_doc)

    input_nonempty = docx_nonempty_paragraphs(input_doc)
    output_nonempty = docx_nonempty_paragraphs(output_doc)
    assert output_nonempty, 'The output document has no resume content.'
    assert output_nonempty[0][0] < input_nonempty[0][0], 'The leading blank page was not reduced.'
    assert output_nonempty[0][1].text.strip() == 'Candidate Profile', 'Resume content no longer starts with the expected heading.'
    for p in output_doc.paragraphs[:output_nonempty[0][0]]:
        assert not docx_has_page_break(p), 'A page break still appears before the resume content.'
    assert input_texts == output_texts, 'Resume text changed unexpectedly while removing the blank page.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
