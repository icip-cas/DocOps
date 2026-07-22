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

    order = ['1. Executive Summary', '2. Findings', '4. Next Steps', '3. Appendix Materials']
    positions = {t: output_texts.index(t) for t in order}
    assert [positions[t] for t in order] == sorted(positions.values()), 'Sections are not in the requested logical order.'
    assert positions['1. Executive Summary'] + 1 < positions['2. Findings'], 'Executive Summary content was not kept with its heading.'
    assert positions['2. Findings'] + 1 < positions['4. Next Steps'], 'Findings content was not kept with its heading.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
