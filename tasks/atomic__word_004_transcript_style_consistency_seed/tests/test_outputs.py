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

    assert input_texts == output_texts, 'Text content should stay unchanged for a style-consistency task.'
    heading_texts = ['Session 1', 'session 2', 'SESSION 3', 'Session Four', 'Wrap-Up']
    body_texts = [
        'Interviewer: Please describe the maintenance bottleneck you observed during the shutdown.',
        'Participant: The main problem was that valve tags did not match the drawings, so every isolation step took longer than expected.',
        'Interviewer: Did the team have enough notice before the work windows changed?',
        'Participant: Not really. The final sequence was updated the night before, and the printouts were already wrong by morning.',
        'Interviewer: What would you prioritize fixing first if the same job were repeated next month?',
    ]
    heading_sigs = [docx_para_signature(docx_para_by_text(output_doc, t)) for t in heading_texts]
    body_sigs = [docx_para_signature(docx_para_by_text(output_doc, t)) for t in body_texts]
    assert len({json.dumps(sig, sort_keys=True) for sig in heading_sigs}) == 1, 'Transcript heading formatting is still inconsistent.'
    assert len({json.dumps(sig, sort_keys=True) for sig in body_sigs}) == 1, 'Transcript body formatting is still inconsistent.'
    for text in [
        'Interview Transcript Packet',
        'Formatting note: the transcript was assembled from multiple team members and pasted together without cleanup.',
    ]:
        assert docx_para_signature(docx_para_by_text(output_doc, text)) == docx_para_signature(docx_para_by_text(input_doc, text)), (
            f'Non-target paragraph formatting changed unexpectedly for: {text}'
        )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
