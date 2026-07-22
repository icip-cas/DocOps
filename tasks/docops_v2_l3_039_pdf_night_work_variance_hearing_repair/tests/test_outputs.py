import json
import os
import re
from pathlib import Path

from pypdf import PdfReader

from verifier_utils import normalize_text, pdf_page_texts, pdf_theme_rect_colors, run_preflight

META_PATH = Path(os.environ.get("TASK_METADATA_PATH", "/tests/task_metadata.json"))
if not META_PATH.exists():
    META_PATH = Path(__file__).parent / "task_metadata.json"
META = json.loads(META_PATH.read_text(encoding="utf-8"))
EXPECT = META["verifier_expectations"]
INPUT_PATH = Path(os.environ.get("INPUT_PATH", META["input_path"]))
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH", META["output_path"]))


def metadata_text(reader):
    meta = reader.metadata or {}
    return "\n".join(str(value or "") for value in meta.values())


def flatten_outline(outline):
    items = []
    for item in outline:
        if isinstance(item, list):
            items.extend(flatten_outline(item))
        else:
            items.append(str(getattr(item, "title", item.get("/Title", item))))
    return items


def all_text(path):
    return "\n".join(pdf_page_texts(path))


def require_all(text, phrases, label):
    normalized = normalize_text(text)
    missing = [phrase for phrase in phrases if normalize_text(phrase) not in normalized]
    assert not missing, f"{label}: missing {missing}"


def require_any(text, phrases, label):
    normalized = normalize_text(text)
    assert any(normalize_text(phrase) in normalized for phrase in phrases), f"{label}: expected one of {phrases}"


def cleanup_context(text, phrase):
    pattern = re.compile(rf"\b{re.escape(phrase)}\b", re.IGNORECASE)
    for match in pattern.finditer(text):
        window = text[max(0, match.start() - 120):match.end() + 180]
        if not re.search(r"\b(no|not|removed|excluded|without|cleaned|public|corrected)\b", window, re.IGNORECASE):
            return False
    return True


def assert_no_forbidden(text):
    normalized = normalize_text(text)
    hard = [
        "do not circulate", "NV-24-7781", "Ridgeway Paving", "Maple Avenue",
        "2026-08-08", "2026-09-01 to 2026-09-05", "555-0166",
        "enforcement score", "staff-only", "phone tree",
    ]
    hits = [phrase for phrase in hard if normalize_text(phrase) in normalized]
    assert not hits, f"Forbidden error remnants still present: {hits}"
    contextual = [
        phrase for phrase in ["DRAFT", "INTERNAL ONLY", "PRIVATE contact"]
        if re.search(rf"\b{re.escape(phrase)}\b", text, re.IGNORECASE) and not cleanup_context(text, phrase)
    ]
    assert not contextual, f"Forbidden cleanup markers still present: {contextual}"


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    assert OUTPUT_PATH != INPUT_PATH


def test_pages_required_text_tables_and_forbidden_errors_removed():
    reader = PdfReader(str(OUTPUT_PATH))
    assert len(reader.pages) == EXPECT["page_count"]
    page_texts = pdf_page_texts(OUTPUT_PATH)
    for idx, title in enumerate(EXPECT["page_titles"]):
        assert title in page_texts[idx], f"Page {idx + 1}: missing title {title!r}"
    require_all(
        "\n".join(page_texts),
        [
            "NV-2026-417",
            "Metro Utilities Joint Trench Program",
            "Cedar Street between 3rd Avenue and 7th Avenue",
            "2026-09-14 to 2026-09-18",
            "9:00 PM to 5:00 AM",
            "2026-08-27",
            "2026-08-25 at 5:00 PM",
            "noisevariance@city.example",
        ],
        "corrected public facts",
    )
    require_any(page_texts[2], ["sequence", "completed in sequence"], "hearing workflow")
    require_any(page_texts[3], ["monitoring", "matrix"], "monitoring matrix")
    require_any(page_texts[4], ["resident", "comment"], "resident routing")
    require_any(page_texts[5], ["release", "checklist", "approve"], "release checklist")
    joined = "\n".join(page_texts)
    assert_no_forbidden(joined + "\n" + metadata_text(reader))


def test_outline_metadata_page_size_and_style():
    reader = PdfReader(str(OUTPUT_PATH))
    assert flatten_outline(reader.outline) == EXPECT["outline_titles"]
    meta = reader.metadata or {}
    assert meta.get("/Title") == EXPECT["metadata"]["title"]
    assert meta.get("/Author") == EXPECT["metadata"]["author"]
    assert meta.get("/Subject") == EXPECT["metadata"]["subject"]
    for page in reader.pages:
        assert abs(float(page.mediabox.width) - EXPECT["style"]["page_width"]) < 1
        assert abs(float(page.mediabox.height) - EXPECT["style"]["page_height"]) < 1
    colors = pdf_theme_rect_colors(OUTPUT_PATH)
    assert EXPECT["style"]["header_bar"] in colors
    assert EXPECT["style"]["accent"] in colors
    for page_text in pdf_page_texts(OUTPUT_PATH):
        assert EXPECT["style"]["footer_text"] in page_text


def test_source_artifact_was_not_modified():
    reader = PdfReader(str(INPUT_PATH))
    text = normalize_text(all_text(INPUT_PATH) + "\n" + metadata_text(reader))
    missing = [phrase for phrase in EXPECT["source_must_contain"] if normalize_text(phrase) not in text]
    assert not missing, f"Source artifact no longer contains expected errors: {missing}"
