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
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        tables = [table for page in pdf.pages for table in page.extract_tables() if table]
    assert tables, 'No table could be extracted from the output PDF.'
    table = tables[0]
    header = [normalize_text(cell or '') for cell in table[0]]
    assert 'owner / team' not in header, 'Combined Owner / Team header still present in the extracted table.'
    assert header[:5] == ['workstream', 'owner', 'team', 'status', 'next date'], (
        f'Expected split table header columns, found {table[0]!r}'
    )
    joined_rows = [' | '.join(cell or '' for cell in row) for row in table[1:]]
    assert line_has_all_tokens(joined_rows, 'Close', 'R. Shah', 'Finance', 'Apr 24'), 'Close row was not split correctly.'
    assert line_has_all_tokens(joined_rows, 'Cutover', 'J. Lim', 'Operations', 'Apr 26'), 'Cutover row was not split correctly.'
    assert line_has_all_tokens(joined_rows, 'QA', 'M. Cole', 'PMO', 'Apr 22'), 'QA row was not split correctly.'
    assert line_has_all_tokens(joined_rows, 'Comms', 'A. Diaz', 'Client Ops', 'Apr 25'), 'Comms row was not split correctly.'


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
