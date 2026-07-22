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


def all_text(path):
    return "\n".join(pdf_page_texts(path))


def norm(text):
    return normalize_text(text)


def require_any(text, phrases, label):
    text_norm = norm(text)
    assert any(norm(phrase) in text_norm for phrase in phrases), (
        f"{label}: expected at least one of {phrases!r}"
    )


def require_all_terms(text, phrases, label):
    text_norm = norm(text)
    missing = [phrase for phrase in phrases if norm(phrase) not in text_norm]
    assert not missing, f"{label}: missing required public terms: {missing}"


def flatten_outline(outline):
    items = []
    for item in outline:
        if isinstance(item, list):
            items.extend(flatten_outline(item))
        else:
            title = getattr(item, "title", None)
            if title is None:
                try:
                    title = item.get("/Title")
                except Exception:
                    title = str(item)
            items.append(str(title))
    return items


def root_obj(reader):
    return reader.trailer["/Root"].get_object()


def has_acroform(reader):
    root = root_obj(reader)
    acro = root.get("/AcroForm")
    if not acro:
        return False
    acro = acro.get_object()
    return bool(acro.get("/Fields"))


def has_annotations(reader):
    for page in reader.pages:
        annots = page.get("/Annots")
        if annots:
            return True
    return False


def has_embedded_files(reader):
    root = root_obj(reader)
    names = root.get("/Names")
    if not names:
        return False
    names = names.get_object()
    return bool(names.get("/EmbeddedFiles"))


def has_javascript(reader):
    root = root_obj(reader)
    if root.get("/OpenAction"):
        return True
    names = root.get("/Names")
    if names:
        names = names.get_object()
        if names.get("/JavaScript"):
            return True
    raw = OUTPUT_PATH.read_bytes() if OUTPUT_PATH.exists() else b""
    return b"/JavaScript" in raw or b"/JS" in raw


def allowed_removed_context(text, start, end):
    window = text[max(0, start - 140): min(len(text), end + 240)]
    return re.search(
        r"\b(removed|excluded|cleared|replaced|no|without|not included|do not include|free of)\b",
        window,
        re.IGNORECASE,
    ) is not None


def assert_no_private_text(path):
    reader = PdfReader(str(path))
    raw_text = all_text(path) + "\n" + metadata_text(reader)
    text = normalize_text(raw_text)

    hard_forbidden = [
        "Vivian Holt",
        "Marco Ruiz",
        "private reviewer",
        "do not publish",
        "2026-03-04",
        "2026-09-01",
    ]
    hits = [phrase for phrase in hard_forbidden if normalize_text(phrase) in text]
    assert not hits, f"Forbidden private/source values still present: {hits}"

    assert not re.search(r"(?<!\d)112(?!\d)", raw_text), "Old public comment count 112 is still present"

    contextual_private = ["personal phone", "legal hold", "unreleased deliberation", "staff-only"]
    contextual_hits = []
    for phrase in contextual_private:
        pattern = re.compile(re.escape(phrase), re.IGNORECASE)
        for match in pattern.finditer(raw_text):
            if not allowed_removed_context(raw_text, match.start(), match.end()):
                contextual_hits.append(phrase)
                break
    assert not contextual_hits, f"Private material still present outside removal context: {contextual_hits}"

    marker_hits = []
    for match in re.finditer(r"\b(?:DRAFT|PRIVATE|INTERNAL)\b", raw_text):
        if not allowed_removed_context(raw_text, match.start(), match.end()):
            marker_hits.append(match.group(0))
    assert not marker_hits, f"Private status markers still present outside removal context: {marker_hits}"


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    assert OUTPUT_PATH != INPUT_PATH


def test_page_count_titles_required_text_and_privacy_cleanup():
    reader = PdfReader(str(OUTPUT_PATH))
    assert len(reader.pages) == EXPECT["page_count"]
    page_texts = pdf_page_texts(OUTPUT_PATH)
    for idx, title in enumerate(EXPECT["page_titles"]):
        assert title in page_texts[idx], f"Page {idx + 1}: missing title {title!r}"

    require_all_terms(page_texts[0], ["ZA-2026-041", "Public Hearing Desk"], "cover page")
    require_any(page_texts[0], ["public release", "public distribution", "public packet"], "cover page")

    require_all_terms(
        page_texts[1],
        ["2026-03-14", "128", "2026-09-15", "Public Hearing Desk", "ZA-2026-041"],
        "corrected hearing summary",
    )
    require_any(page_texts[1], ["hearing date", "hearing"], "corrected hearing summary")
    require_any(page_texts[1], ["comment count", "comments"], "corrected hearing summary")
    require_any(page_texts[1], ["response deadline", "deadline"], "corrected hearing summary")

    require_all_terms(page_texts[2], ["Category", "Count"], "comment category matrix")
    require_any(page_texts[2], ["128", "deduplicated", "public comments"], "comment category matrix")
    require_any(page_texts[2], ["handling", "response", "theme", "lead lane", "public"], "comment category matrix")

    require_any(page_texts[3], ["workflow", "response"], "response workflow")
    require_any(page_texts[3], ["ZA-2026-041", "Public Hearing Desk", "2026-09-15"], "response workflow")
    require_any(page_texts[3], ["publish", "release", "archive", "draft", "quality", "review"], "response workflow")

    require_any(page_texts[4], ["AcroForm", "form fields", "form controls", "interactive elements"], "release controls")
    require_any(page_texts[4], ["Embedded", "attachments", "embedded content"], "release controls")
    require_any(page_texts[4], ["JavaScript", "script actions", "document scripts", "Document actions", "executable"], "release controls")
    require_any(page_texts[4], ["metadata", "public metadata", "Nonpublic metadata"], "release controls")
    require_any(page_texts[4], ["Appendix", "appendix references", "index"], "release controls")

    assert_no_private_text(OUTPUT_PATH)


def test_pdf_objects_are_flattened_and_clean():
    reader = PdfReader(str(OUTPUT_PATH))
    assert not has_acroform(reader), "Output still has AcroForm fields"
    assert not has_annotations(reader), "Output still has annotations"
    assert not has_embedded_files(reader), "Output still has embedded files"
    assert not has_javascript(reader), "Output still has JavaScript/OpenAction"


def test_outline_metadata_page_size_and_style():
    reader = PdfReader(str(OUTPUT_PATH))
    assert flatten_outline(reader.outline) == EXPECT["outline_titles"]
    meta = reader.metadata or {}
    assert meta.get("/Title") == EXPECT["metadata"]["title"]
    assert meta.get("/Author") == EXPECT["metadata"]["author"]
    assert meta.get("/Subject") == EXPECT["metadata"]["subject"]
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        assert abs(width - EXPECT["style"]["page_width"]) < 1
        assert abs(height - EXPECT["style"]["page_height"]) < 1
    colors = pdf_theme_rect_colors(OUTPUT_PATH)
    assert EXPECT["style"]["header_bar"] in colors
    assert EXPECT["style"]["accent"] in colors
    for page_text in pdf_page_texts(OUTPUT_PATH):
        assert EXPECT["style"]["footer_text"] in page_text


def test_source_artifact_was_not_modified():
    reader = PdfReader(str(INPUT_PATH))
    source = EXPECT["source_must_have"]
    assert has_acroform(reader) == source["acroform"]
    assert has_annotations(reader) == source["annotations"]
    assert has_embedded_files(reader) == source["embedded_files"]
    assert has_javascript(reader) == source["javascript"]
    assert source["metadata_author"] in metadata_text(reader)
    text = normalize_text(all_text(INPUT_PATH) + "\n" + metadata_text(reader))
    missing = [phrase for phrase in source["phrases"] if normalize_text(phrase) not in text]
    assert not missing, f"Source artifact no longer contains expected defects: {missing}"
