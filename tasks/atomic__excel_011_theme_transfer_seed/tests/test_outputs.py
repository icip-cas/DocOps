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

    assert wb_out.sheetnames == wb_in.sheetnames, 'Theme transfer should not reorder sheets.'
    tab_colors = []
    for sheet_name in wb_out.sheetnames:
                ws = wb_out[sheet_name]
                assert workbook_values_signature(wb_out_values)[sheet_name] == workbook_values_signature(wb_in_values)[sheet_name], f'Content changed on sheet {sheet_name} during theme transfer.'
                tab = ws.sheet_properties.tabColor.rgb if ws.sheet_properties.tabColor and ws.sheet_properties.tabColor.rgb else None
                if tab and len(tab) == 8:
                    tab = tab[2:]
                assert tab is not None, f'Sheet {sheet_name} is missing a tab color.'
                assert_blueish(tab, f'{sheet_name} tab color')
                tab_colors.append(tab)
                title_fill = cell_fill_rgb(ws['A1'])
                header_fill = cell_fill_rgb(ws['A3'])
                assert title_fill is not None and header_fill is not None, f'Sheet {sheet_name} is missing themed fills on title/header cells.'
                assert_blueish(title_fill, f'{sheet_name} title fill')
                assert_blueish(header_fill, f'{sheet_name} header fill')
    assert len(set(tab_colors)) == 1, 'All visible sheets should share one tab-color theme.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
