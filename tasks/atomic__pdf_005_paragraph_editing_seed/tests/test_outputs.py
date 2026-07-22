import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from verifier_utils import *  # noqa: F401,F403

META_PATH = Path(os.environ.get('TASK_METADATA_PATH', '/tests/task_metadata.json'))
META = json.loads(META_PATH.read_text(encoding='utf-8'))
INPUT_PATH = Path(os.environ.get('INPUT_PATH', META['input_path']))
OUTPUT_PATH = Path(os.environ.get('OUTPUT_PATH', META['output_path']))
TASK_ID = META['task_id']
DOC_TYPE = META['doc_type']


def verify_task() -> None:
    text_out = pdf_text(OUTPUT_PATH)

    forbid_any(text_out, ['missed two promised dates again', 'cannot rely on the timeline they gave us'], TASK_ID)
    require_any_group(text_out, [['missed two promised dates'], ['two promised dates were missed']], TASK_ID)
    require_any_group(
        text_out,
        [['incomplete package'], ['package remained incomplete'], ['package was incomplete'], ['latest package was incomplete']],
        TASK_ID,
    )
    require_any_group(
        text_out,
        [
            ['timeline should remain uncertain until the revised materials are confirmed'],
            ['timeline remains uncertain until the revised materials are confirmed'],
            ['timeline remains uncertain until revised materials are confirmed'],
            ['timeline should be treated as uncertain until revised materials are confirmed'],
        ],
        TASK_ID,
    )
    require_all(
        text_out,
        ['Client Notice Draft', 'delayed package includes both the reconciled data feed and the revised regional mapping file'],
        TASK_ID,
    )


def test_semantic_verifier() -> None:
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    verify_task()
