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


def require_all_terms(text, phrases, label):
    text_norm = normalize_text(text)
    missing = [phrase for phrase in phrases if normalize_text(phrase) not in text_norm]
    assert not missing, f"{label}: missing required public terms: {missing}"


def require_any(text, phrases, label):
    text_norm = normalize_text(text)
    assert any(normalize_text(phrase) in text_norm for phrase in phrases), (
        f"{label}: expected at least one of {phrases!r}"
    )


def allowed_cleanup_context(text, start, end):
    window = text[max(0, start - 120): min(len(text), end + 180)]
    return re.search(r"\b(no|removed|clean|cleared|without|not included|free of)\b", window, re.I)


def assert_no_forbidden_text(text):
    text_norm = normalize_text(text)
    hard_forbidden = [
        "Greenvile",
        "permti",
        "publc",
        "adress",
        "inspeciton",
        "reciept",
        "conditon",
        "P-2046-17D",
        "2026-08-04",
        "2026-09-08",
        "2026-08-19",
        "$1,580",
        "5 conditions",
        "duplicate-page",
        "OCR-error",
        "do-not-post",
    ]
    hits = [phrase for phrase in hard_forbidden if normalize_text(phrase) in text_norm]
    assert not hits, f"Forbidden error remnants still present: {hits}"

    contextual_hits = []
    for phrase in ["DRAFT", "PRIVATE", "scratch", "obsolete"]:
        flags = 0 if phrase in {"DRAFT", "PRIVATE"} else re.I
        for match in re.finditer(rf"\b{re.escape(phrase)}\b", text, flags):
            if not allowed_cleanup_context(text, match.start(), match.end()):
                contextual_hits.append(phrase)
                break
    assert not contextual_hits, f"Draft/private/scratch markers still present: {contextual_hits}"


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    assert OUTPUT_PATH != INPUT_PATH


def test_pages_required_text_tables_and_forbidden_errors_removed():
    reader = PdfReader(str(OUTPUT_PATH))
    assert len(reader.pages) == EXPECT["page_count"]
    page_texts = pdf_page_texts(OUTPUT_PATH)
    for idx, title in enumerate(EXPECT["page_titles"]):
        assert title in page_texts[idx], f"Page {idx + 1}: missing title {title!r}"

    joined = "\n".join(page_texts)
    require_all_terms(
        joined,
        [
            "Greenville",
            "Arbor House Cooperative",
            "P-2046-17B",
            "412 Cedar Avenue",
            "2026-08-14",
            "2026-09-18",
            "2026-08-29",
            "$1,850",
        ],
        "corrected public permit facts",
    )
    assert_no_forbidden_text(joined + "\n" + metadata_text(reader))
    require_any(joined, ["7 public conditions", "Public condition count", "condition count is listed as 7", "Condition 7", "Seven public conditions"], "condition count")

    require_all_terms(page_texts[0], ["Public Permit Notice Repair Packet"], "cover page")
    require_any(page_texts[0], ["Arbor House Cooperative", "P-2046-17B", "412 Cedar Avenue"], "cover page")
    require_any(page_texts[1], ["Errata", "Corrections", "Corrected"], "errata page")
    require_all_terms(page_texts[1], ["Greenville"], "errata page")
    require_any(page_texts[1], ["permit", "public", "address", "inspection", "receipt", "condition"], "errata page")
    require_all_terms(page_texts[2], ["Arbor House Cooperative", "P-2046-17B", "412 Cedar Avenue"], "applicant and parcel summary")
    require_all_terms(page_texts[3], ["2026-08-14", "2026-09-18", "2026-08-29"], "inspection schedule")
    require_any(page_texts[3], ["Notice", "posting"], "inspection schedule")
    require_any(page_texts[3], ["Inspection", "Site Inspection"], "inspection schedule")
    require_any(page_texts[3], ["Hearing", "Public Hearing"], "inspection schedule")
    require_all_terms(page_texts[4], ["$1,850"], "fee and condition matrix")
    require_any(page_texts[4], ["7 public conditions", "Public Conditions", "condition count", "Condition 7", "Seven public conditions"], "fee and condition matrix")
    require_any(page_texts[5], ["Checklist", "Posting Checks", "posting"], "public posting checklist")


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
