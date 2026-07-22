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
    reader_out = PdfReader(str(OUTPUT_PATH))
    page_texts_out = pdf_page_texts(OUTPUT_PATH)

    assert len(reader_out.pages) == 2, 'Generated summary PDF should retain both packet pages.'
    page1 = page_texts_out[0]
    assert 'Management Summary - intentionally blank' not in page1, 'Management Summary placeholder is still blank.'
    require_any_group(page1, [['outdated mapping file'], ['header labels']], TASK_ID)
    require_all(page1, ['outdated approver list'], TASK_ID)
    require_any_group(
        page1,
        [
            ['tagged differently across the APAC and EMEA queues'],
            ['inconsistent defect labeling across queues'],
            ['tickets were tagged differently across queues'],
            ['inconsistent ticket labeling across queues'],
        ],
        TASK_ID,
    )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
