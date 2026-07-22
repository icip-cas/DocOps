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
    wb_out, wb_out_values = load_xlsx_pair(OUTPUT_PATH)
    wb_in = load_workbook(INPUT_PATH)
    wb_in_values = load_workbook(INPUT_PATH, data_only=True)

    ws_in = wb_in['Weekly Plan']
    ws_out = wb_out['Weekly Plan']
    ps = ws_out.page_setup
    assert str(ps.orientation).lower().endswith('landscape'), 'Page orientation must be landscape.'
    assert int(ps.fitToWidth or 0) == 1, 'Sheet should be configured to fit to one page wide.'
    titles = ws_out.print_title_rows or ''
    assert '$2:$2' in titles or '$1:$2' in titles, f'Header row should repeat in print titles, got {titles}'
    for row in [3, 4, 5, 6]:
        out_height = ws_out.row_dimensions[row].height or 0
        in_height = ws_in.row_dimensions[row].height or 0
        assert out_height >= in_height, f'Row {row} height should not shrink below the original wrapped-text height.'
    assert workbook_values_signature(wb_out_values) == workbook_values_signature(wb_in_values), (
        'Layout task should not change any cell values.'
    )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
