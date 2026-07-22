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

    assert '[TODO:' not in '\n'.join(output_texts), 'Placeholder TODO text is still present.'
    idx = output_texts.index('Executive Summary Placeholder')
    assert idx + 1 < len(output_texts), 'No generated summary paragraph found after the placeholder heading.'
    summary = output_texts[idx + 1]
    assert len(summary) >= 100, 'Generated executive summary is too short.'
    require_any_group(summary, [['shutdown moved', 'schedule', 'timeline']], TASK_ID)
    require_any_group(summary, [['paperwork', 'documents', 'vendor']], TASK_ID)
    require_any_group(summary, [['owner', 'site'], ['finance', 'overtime', 'risk']], TASK_ID)
    assert output_texts[:idx] == input_texts[:idx], 'Unexpected edits occurred before the summary placeholder.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
