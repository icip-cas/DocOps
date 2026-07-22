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
        window = text[max(0, match.start() - 120): match.end() + 180]
        if not re.search(r"\b(no|not|removed|cleaned|excluded|without|superseded|corrected|replaced|pass)\b", window, re.IGNORECASE):
            return False
    return True


def assert_no_forbidden(text):
    hard = ["BN-2026-8428", "$824", "2026-06-08", "2026-07-08"]
    normalized = normalize_text(text)
    hits = [phrase for phrase in hard if normalize_text(phrase) in normalized]
    assert not hits, f"Forbidden error remnants still present: {hits}"
    contextual = [
        phrase
        for phrase in ["DRAFT", "PRIVATE", "caseworker direct line", "do not mail", "scratch"]
        if normalize_text(phrase) in normalized and not cleanup_context(text, phrase)
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
    require_all(page_texts[0], ["BN-2026-8842", "Benefits Service Center"], "title page")
    require_all(page_texts[1], ["2026-06-18", "$842", "2026-07-18"], "corrected summary")
    require_all(page_texts[2], ["$842"], "eligibility/payment page")
    require_any(page_texts[2], ["Eligible", "Eligibility", "Approved"], "eligibility/payment page")
    require_all(page_texts[3], ["2026-07-18", "Benefits Service Center"], "appeal workflow")
    require_any(page_texts[3], ["File appeal", "appeal"], "appeal workflow")
    require_any(page_texts[4], ["Errata", "Release", "Checklist", "Complete", "PASS"], "release checklist")
    joined = "\n".join(page_texts)
    require_all(joined, ["BN-2026-8842", "2026-06-18", "$842", "2026-07-18", "Benefits Service Center", "1-800-555-0188"], "corrected public values")
    assert_no_forbidden(joined + "\n" + metadata_text(reader))


def test_outline_metadata_page_size_and_style():
    reader = PdfReader(str(OUTPUT_PATH))
    assert flatten_outline(reader.outline) == EXPECT["outline_titles"]
    meta = reader.metadata or {}
    assert "Public Benefits Notice" in str(meta.get("/Title", ""))
    assert "Benefits Service Center" in str(meta.get("/Author", ""))
    assert str(meta.get("/Subject", "")).strip()
    assert_no_forbidden(metadata_text(reader))
    for page in reader.pages:
        assert abs(float(page.mediabox.width) - EXPECT["style"]["page_width"]) < 1
        assert abs(float(page.mediabox.height) - EXPECT["style"]["page_height"]) < 1
    colors = pdf_theme_rect_colors(OUTPUT_PATH)
    assert colors, "PDF should contain styled colored rectangles"
    for page_text in pdf_page_texts(OUTPUT_PATH):
        assert "Page" in page_text and (
            "BN-2026-8842" in page_text
            or "Benefits Service Center" in page_text
            or "Public Benefits Notice Repair Packet" in page_text
        )


def test_source_artifact_was_not_modified():
    reader = PdfReader(str(INPUT_PATH))
    text = normalize_text(all_text(INPUT_PATH) + "\n" + metadata_text(reader))
    missing = [phrase for phrase in EXPECT["source_must_contain"] if normalize_text(phrase) not in text]
    assert not missing, f"Source artifact no longer contains expected errors: {missing}"
