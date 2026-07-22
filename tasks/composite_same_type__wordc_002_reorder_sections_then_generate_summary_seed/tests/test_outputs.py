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
    output_doc = Document(OUTPUT_PATH)
    texts = docx_texts(output_doc)
    headings = [t for t in texts if t in ['Executive Summary', 'Findings', 'Next Steps', 'Appendix Materials']]
    assert headings[:4] == ['Executive Summary', 'Findings', 'Next Steps', 'Appendix Materials'], f'Unexpected section order: {headings}'
    idx = texts.index('Executive Summary')
    summary_items = texts[idx + 1: idx + 4]
    assert texts[idx + 4] == 'Findings', 'Executive Summary should contain exactly three inserted summary bullets before Findings.'
    assert_prefixed_items(summary_items, ['Permit:', 'Supplier:', 'Safety:'], 'Executive Summary section')
    require_group_hits(summary_items[0], [['permit'], ['delay', 'delayed']], 'Permit bullet')
    require_group_hits(summary_items[1], [['supplier'], ['backlog', 'late', 'second half of june']], 'Supplier bullet')
    require_group_hits(summary_items[2], [['safety'], ['workaround', 'temporary']], 'Safety bullet')


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
