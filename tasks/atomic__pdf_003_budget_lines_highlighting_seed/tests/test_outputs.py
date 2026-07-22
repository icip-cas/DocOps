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

    assert len(reader_out.pages) == 1, 'Highlighting task should still produce one page.'
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        page = pdf.pages[0]
        assert line_is_visually_highlighted(page, 'Line 02 Temporary chiller rental $10,050'), 'Line 02 was not highlighted.'
        assert line_is_visually_highlighted(page, 'Line 04 Weekend crane access $12,400'), 'Line 04 was not highlighted.'
        assert not line_is_visually_highlighted(page, 'Line 06 Valve actuator batch $10,000'), (
            'Line 06 should not be highlighted because the threshold is strict > $10,000.'
        )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
