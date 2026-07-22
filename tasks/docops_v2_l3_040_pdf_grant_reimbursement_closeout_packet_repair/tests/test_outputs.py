import json
import os
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


def contains_all(text, phrases):
    normalized = normalize_text(text)
    return all(normalize_text(phrase) in normalized for phrase in phrases)


def contains_groups(text, groups):
    normalized = normalize_text(text)
    for group in groups:
        if not any(normalize_text(term) in normalized for term in group):
            return False
    return True


def forbidden_hits(text):
    normalized = normalize_text(text)
    hits = []
    for phrase in EXPECT["forbidden_phrases"]:
        phrase_norm = normalize_text(phrase)
        if phrase_norm not in normalized:
            continue
        # Public release-control statements such as "draft material removed" are
        # acceptable; old values and account/identity fragments are not.
        if phrase_norm in {"draft", "internal only"}:
            allowed_contexts = [
                f"no {phrase_norm}",
                f"{phrase_norm} material removed",
                f"{phrase_norm} markers removed",
                f"{phrase_norm} notes removed",
                f"{phrase_norm} content appears",
            ]
            if any(context in normalized for context in allowed_contexts):
                continue
        hits.append(phrase)
    return hits


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

    corrected_facts = [
        "GRC-2026-118",
        "North Valley Food Bank",
        "2026-04-01 to 2026-06-30",
        "$184,620",
        "$179,420",
        "$5,200",
        "grants-closeout@state.example",
    ]
    assert contains_all(joined, corrected_facts)
    assert "$184,600" not in joined

    page_semantics = {
        1: [["claim"], ["eligible"], ["disallowed"], ["final payable", "payable"]],
        2: [["cost"], ["claimed", "submitted"], ["disallowed"], ["eligible"]],
        3: [["evidence"], ["review", "accepted", "verified", "matched"], ["status", "result", "outcome", "ready"]],
        4: [["step", "workflow", "hold"], ["owner", "responsible", "reviewer", "office", "lead"], ["release", "payment", "payable", "authorization"]],
        5: [["certification", "checklist"], ["restricted", "public-safe", "public release", "release"], ["footer", "metadata", "bookmark", "contact", "payment-account"]],
    }
    for idx, groups in page_semantics.items():
        assert contains_groups(page_texts[idx], groups), f"Page {idx + 1}: missing semantic table terms {groups}"

    combined = normalize_text(joined + "\n" + metadata_text(reader))
    hits = forbidden_hits(combined)
    assert not hits, f"Forbidden error remnants still present: {hits}"


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
