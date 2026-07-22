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

    table = output_doc.tables[0]
    headers = [c.text.strip() for c in table.rows[0].cells]
    assert headers == ['Item', 'Owner', 'Status', 'Due Date'], f'Unexpected table headers: {headers}'
    expected_rows = [
                ['Drain survey', 'Nina', '', '22 Apr'],
                ['Valve relabel', 'Karl', '', '24 Apr'],
                ['Shutdown note', 'Ava', '', '26 Apr'],
            ]
    actual_rows = [[c.text.strip() for c in row.cells] for row in table.rows[1:]]
    assert actual_rows == expected_rows, f'Table rows were not preserved while inserting Status column: {actual_rows}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
