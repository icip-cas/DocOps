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
    prs_in = Presentation(INPUT_PATH)
    prs_out = Presentation(OUTPUT_PATH)
    input_text = ppt_texts(prs_in)
    output_text = ppt_texts(prs_out)

    slide = prs_out.slides[0]
    table_shape = next(shape for shape in slide.shapes if shape.has_table)
    table = table_shape.table
    headers = [table.cell(0, c).text.strip() for c in range(len(table.columns))]
    assert headers == ['Milestone', 'Owner', 'Due Date', 'Status'], f'Unexpected table headers: {headers}'
    row1 = [table.cell(1, c).text.strip() for c in range(len(table.columns))]
    assert row1 == ['Variation order approved', '', '22 Apr', 'Open'], f'Row content was not preserved after inserting Owner column: {row1}'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
