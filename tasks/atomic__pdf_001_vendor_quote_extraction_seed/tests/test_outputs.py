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
    text = answer_text(OUTPUT_PATH)
    assert len(text) >= 10, 'Final answer is too short.'

    lines = nonempty_lines(text)
    assert len(lines) == 6, f'{TASK_ID}: expected exactly 6 labeled lines, found {len(lines)}.'
    assert line_has_all_tokens(lines, 'Vendor:', 'Polar Mechanical Services Ltd.'), f'{TASK_ID}: missing vendor field/value line.'
    assert line_has_all_tokens(lines, 'Quote Ref:', 'Q-7741-R2'), f'{TASK_ID}: missing quote-reference field/value line.'
    assert line_has_all_tokens(lines, 'Quoted Total Including Options:', '54,400'), f'{TASK_ID}: missing total field/value line.'
    assert line_has_all_tokens(lines, 'Delivery Window:', '28 May 2026'), f'{TASK_ID}: missing delivery-window field/value line.'
    assert line_has_all_tokens(lines, 'Warranty Term:', '18 months'), f'{TASK_ID}: missing warranty field/value line.'
    assert line_has_all_tokens(lines, 'Project Contact Email:', 'elena.brooks@polar-ms.example'), f'{TASK_ID}: missing contact-email field/value line.'
    assert_single_expected_money(text, '54,400', TASK_ID)


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
