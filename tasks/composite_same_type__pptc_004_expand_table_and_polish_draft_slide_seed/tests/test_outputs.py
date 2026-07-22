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
    prs = Presentation(OUTPUT_PATH)
    slide = prs.slides[0]
    table = slide.shapes[1].table
    headers = [table.cell(0, c).text.strip() for c in range(4)]
    assert headers == ['Milestone', 'Status', 'Due Date', 'Notes'], f'Unexpected table headers: {headers}'
    row1 = [table.cell(1, c).text.strip() for c in range(4)]
    assert row1[0] == 'Vendor recovery plan' and row1[2] == '14 Jun 2026', 'Milestone and due date must be preserved.'
    bullet_box = slide.shapes[2]
    bullets = [' '.join(p.text.split()) for p in bullet_box.text_frame.paragraphs if p.text.strip()]
    assert_prefixed_items(bullets, ['Vendor status:', 'Install date:', 'Client update:'], 'Milestone slide bullets')
    combined = normalize_text(' '.join(bullets))
    for forbidden in ['kinda', 'super', 'probably', 'a bit longer']:
        assert forbidden not in combined, f'Informal wording still present: {forbidden}'
    require_group_hits(
        bullets[0],
        [['vendor'], ['delay', 'late', 'slip', 'behind schedule', 'lagging', 'behind plan', 'at risk']],
        'Vendor status bullet',
    )
    require_group_hits(
        bullets[1],
        [['install date'], ['revised', 'updated', 'tentative'], ['uncertain', 'pending', 'subject to change', 'awaiting confirmation', 'unconfirmed']],
        'Install date bullet',
    )
    require_group_hits(
        bullets[2],
        [['client update'], ['dependency', 'pending vendor confirmation', 'after vendor confirmation', 'after the vendor confirms', 'install date is confirmed', 'until the install date is confirmed', 'once the install date is confirmed', 'once vendor confirms', 'hold until confirmation']],
        'Client update bullet',
    )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
