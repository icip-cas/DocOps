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

    assert input_texts == output_texts, 'Image layout task should not rewrite document text.'
    in_sizes = docx_inline_shape_sizes(input_doc)
    out_sizes = docx_inline_shape_sizes(output_doc)
    assert len(out_sizes) == 3, 'Expected exactly three images after layout cleanup.'
    assert coeff_of_var([w for w, _ in out_sizes]) <= coeff_of_var([w for w, _ in in_sizes]), 'Image widths are not more uniform than the input layout.'
    assert coeff_of_var([h for _, h in out_sizes]) <= coeff_of_var([h for _, h in in_sizes]), 'Image heights are not more uniform than the input layout.'



def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
