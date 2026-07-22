import json
import os
import re
from pathlib import Path

import pdfplumber
from pypdf import PdfReader

from verifier_utils import normalize_text, run_preflight

META_PATH = Path(os.environ.get("TASK_METADATA_PATH", "/tests/task_metadata.json"))
if not META_PATH.exists():
    META_PATH = Path(__file__).parent / "task_metadata.json"
META = json.loads(META_PATH.read_text(encoding="utf-8"))
EXPECT = META["verifier_expectations"]
INPUT_PATH = Path(os.environ.get("INPUT_PATH", META["input_path"]))
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH", META["output_path"]))


def extract_pages(path):
    with pdfplumber.open(path) as pdf:
        return [page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages]


def color_to_hex(color):
    if color is None:
        return None
    if isinstance(color, (int, float)):
        val = max(0, min(255, round(float(color) * 255)))
        return f"{val:02X}{val:02X}{val:02X}"
    vals = list(color)
    if len(vals) < 3:
        return None
    rgb = [max(0, min(255, round(float(v) * 255))) for v in vals[:3]]
    return "".join(f"{v:02X}" for v in rgb)


def page_has_header_band(page, expected_hex):
    for rect in page.rects:
        if rect.get("top", 9999) <= 40 and rect.get("height", 0) >= 30:
            if color_to_hex(rect.get("non_stroking_color")) == expected_hex:
                return True
    return False


def page_text_norm(text):
    return normalize_text(text)


def assert_terms(text, terms, label):
    normalized = page_text_norm(text)
    missing = [term for term in terms if normalize_text(term) not in normalized]
    assert not missing, f"{label}: missing {missing}"


def forbidden_hits(text):
    hits = []
    raw = text
    normalized = normalize_text(text)
    for phrase in EXPECT["forbidden_phrases"]:
        if phrase in {"DRAFT", "PRIVATE", "INTERNAL"}:
            if re.search(rf"\\b{phrase}\\b", raw):
                hits.append(phrase)
        elif phrase == "scratch":
            if re.search(r"\\bscratch\\b", normalized):
                hits.append(phrase)
        elif normalize_text(phrase) in normalized:
            hits.append(phrase)
    return hits


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    assert OUTPUT_PATH != INPUT_PATH


def test_page_count_titles_and_private_material_removed():
    pages = extract_pages(OUTPUT_PATH)
    assert len(pages) == EXPECT["page_count"]
    for idx, expected_title in enumerate(EXPECT["page_titles"]):
        first_lines = pages[idx].splitlines()[:4]
        assert any(expected_title in line for line in first_lines), f"Page {idx + 1}: missing title {expected_title!r}"
    hits = forbidden_hits("\n".join(pages))
    assert not hits, f"Forbidden phrases still present: {hits}"


def test_required_public_lines_by_page():
    pages = extract_pages(OUTPUT_PATH)
    assert_terms(pages[0], ["public", "annual", "compliance report", "permit year", "2026"], "Page 1 cover")
    assert_terms(pages[1], ["Substantial compliance", "8 active outfalls", "5 of 6", "3 business days"], "Page 2 summary")
    for line in EXPECT["required_lines_by_page"]["3"]:
        assert line in pages[2], f"Page 3: missing required milestone line {line!r}"
    for line in [
        "Outfall 07 | E. coli | 140 | 235 | Within threshold",
        "Outfall 12 | Turbidity | 7 NTU | 10 NTU | Within threshold",
        "Outfall 18 | pH | 7.4 | 6.5-8.5 | Within threshold",
    ]:
        assert line in pages[3], f"Page 4: missing monitoring line {line!r}"
    for line in [
        "CA-001 | Repair sediment control at Maple Yard | 2026-08-25 | Field Operations",
        "CA-002 | Replace damaged inlet marker set | 2026-08-28 | Public Works",
        "CA-003 | Update contractor spill checklist | 2026-09-05 | Compliance Office",
    ]:
        assert line in pages[4], f"Page 5: missing corrective action line {line!r}"
    for line in EXPECT["required_lines_by_page"]["6"]:
        assert line in pages[5], f"Page 6: missing complaint line {line!r}"
    for line in EXPECT["required_lines_by_page"]["7"]:
        assert line in pages[6], f"Page 7: missing infrastructure line {line!r}"
    for line in EXPECT["required_lines_by_page"]["8"][:5]:
        assert line in pages[7], f"Page 8: missing QA line {line!r}"
    assert_terms(
        pages[8],
        ["Header band height", "40", "243B53", "007C89", "2E7D32", "C55A11", "7030A0", EXPECT["footer"]],
        "Page 9 style guide",
    )
    for line in EXPECT["required_lines_by_page"]["10"]:
        assert line in pages[9], f"Page 10: missing appendix line {line!r}"


def test_footer_and_page_numbering():
    pages = extract_pages(OUTPUT_PATH)
    for idx, text in enumerate(pages, start=1):
        assert EXPECT["footer"] in text, f"Page {idx}: missing public footer"
        assert f"Page {idx} of {EXPECT['page_count']}" in text, f"Page {idx}: wrong page numbering"


def test_header_band_style_migration():
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        assert len(pdf.pages) == len(EXPECT["header_colors"])
        for idx, expected_hex in enumerate(EXPECT["header_colors"]):
            assert page_has_header_band(pdf.pages[idx], expected_hex), f"Page {idx + 1}: missing header band {expected_hex}"


def test_pdf_metadata_cleaned():
    reader = PdfReader(str(OUTPUT_PATH))
    info = reader.metadata
    meta_text = "\n".join(
        str(value or "")
        for value in [info.title, info.subject, info.author, info.get("/Keywords")]
    )
    hits = forbidden_hits(meta_text)
    assert not hits, f"Metadata contains forbidden material: {hits}"
    assert_terms(info.title or "", ["Riverbend", "Stormwater", "Compliance Report"], "PDF title")
    assert_terms((info.subject or "") + "\n" + (info.get("/Keywords") or ""), ["stormwater", "compliance", "2026"], "PDF subject/keywords")
    assert_terms(info.author or "", ["Riverbend"], "PDF author")


def test_source_artifact_was_not_modified():
    source_pages = extract_pages(INPUT_PATH)
    assert len(source_pages) > EXPECT["page_count"]
    text = normalize_text("\n".join(source_pages))
    assert "draft" in text
    assert "contractor dispute" in text
    assert "resident personal phone" in text
