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


def normalized(text):
    return normalize_text(text)


def assert_terms(text, terms, label):
    norm = normalized(text)
    missing = [term for term in terms if normalized(term) not in norm]
    assert not missing, f"{label}: missing public terms {missing}"


def assert_any(text, terms, label):
    norm = normalized(text)
    assert any(normalized(term) in norm for term in terms), f"{label}: missing one of {terms}"


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


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    assert OUTPUT_PATH != INPUT_PATH


def test_page_count_titles_and_private_material_removed():
    pages = extract_pages(OUTPUT_PATH)
    assert len(pages) == EXPECT["page_count"]
    for idx, expected_title in enumerate(EXPECT["page_titles"]):
        first_lines = pages[idx].splitlines()[:4]
        assert any(expected_title in line for line in first_lines), f"Page {idx + 1}: missing title {expected_title!r}"
    text = normalize_text("\n".join(pages))
    hits = [p for p in EXPECT["forbidden_phrases"] if normalize_text(p) in text]
    assert not hits, f"Forbidden phrases still present: {hits}"


def test_required_public_lines_by_page():
    pages = extract_pages(OUTPUT_PATH)
    all_text = "\n".join(pages)
    assert_terms(all_text, ["97.2%", "Station C-71", "2026-02-08", "Nia Rowan", "2026-03-10"], "correction facts")

    assert_terms(pages[0], ["MetroCharge", "EV", "Reliability", "Q1 2026"], "cover page")
    assert_any(pages[0], ["public reliability report", "public release", "public report"], "cover page public status")

    assert_terms(pages[1], ["Network uptime", "97.2%", "Active public stations", "42", "2026-02-08", "2026-03-10"], "reliability summary page")

    assert_terms(pages[2], ["Station", "Station C-71", "East Loop", "97.2%", "Station A-04", "98.1%", "Station B-19", "96.8%"], "station uptime page")
    assert_any(pages[2], ["Meets target", "Monitor"], "station uptime status")

    assert_terms(pages[3], ["OUT-001", "Payment gateway outage", "2026-02-08", "OUT-002", "Station C-71"], "outage timeline page")
    assert_any(pages[3], ["Resolved", "Closed", "Complete", "Published"], "outage timeline status")

    assert_terms(pages[4], ["MA-001", "Station C-71", "Nia Rowan", "2026-02-12", "Complete"], "maintenance action page")

    assert_terms(pages[5], ["EQ-001", "two chargers", "East Loop", "Active"], "equity access page")
    assert_any(pages[5], ["EQ-002", "downtime notices", "public access"], "equity access commitments")

    assert_terms(pages[6], ["Payment outage", "Gateway configuration fault", "Closed", "2026-03-10"], "payment incident page")

    assert_terms(pages[7], ["Reliability bulletin", "Drivers", "2026-03-10", "Council"], "public communications page")
    assert_any(pages[7], ["Posted", "Published", "Sent"], "public communications status")

    assert_terms(pages[8], ["Header band", "40", "0B132B", "0F766E", "B45309", "1D4ED8", "6D28D9"], "publication style guide page")
    assert_terms(pages[8], [EXPECT["footer"]], "publication style guide footer")

    assert_terms(pages[9], ["Appendix", "Uptime", "Release approval", "Reliability Office Director"], "appendix register page")


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
    metadata_text = "\n".join(str(value or "") for value in [info.title, info.subject, info.author, info.get("/Keywords")])
    assert_terms(metadata_text, ["MetroCharge", "Reliability"], "PDF metadata")
    assert_any(metadata_text, ["EV", "charging"], "PDF metadata EV scope")
    assert_any(metadata_text, ["public", "Q1 2026", "2026"], "PDF metadata public scope")
    hits = [p for p in EXPECT["forbidden_phrases"] if normalize_text(p) in normalize_text(metadata_text)]
    assert not hits, f"Metadata contains forbidden material: {hits}"


def test_source_artifact_was_not_modified():
    source_pages = extract_pages(INPUT_PATH)
    assert len(source_pages) > EXPECT["page_count"]
    text = normalize_text("\n".join(source_pages))
    assert "draft" in text
    assert "vendor penalty note" in text
    assert "driver complaint phone" in text
