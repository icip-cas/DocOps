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


def collapse_doubled_letters(text):
    # Some PDFs extract bold/overlaid text as "MMeeaall". Collapse repeated
    # adjacent characters so title checks are about content, not extraction noise.
    return re.sub(r"([A-Za-z])\1", r"\1", text)


def normalized(text):
    return normalize_text(collapse_doubled_letters(text))


def assert_terms(text, terms, label):
    norm = normalized(text)
    missing = [term for term in terms if normalized(term) not in norm]
    assert not missing, f"{label}: missing public terms {missing}"


def assert_any(text, terms, label):
    norm = normalized(text)
    assert any(normalized(term) in norm for term in terms), f"{label}: missing one of {terms}"


def forbidden_scan_text(text):
    norm = normalize_text(text)
    # Do not reject benign public wording such as "internal review materials
    # excluded"; still reject sensitive source phrases and old corrected values.
    norm = re.sub(r"internal [a-z ]{0,40}excluded", "public exclusion note", norm)
    return norm


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
        first_text = "\n".join(pages[idx].splitlines()[:6])
        assert normalized(expected_title) in normalized(first_text), f"Page {idx + 1}: missing title {expected_title!r}"
    text = forbidden_scan_text("\n".join(pages))
    public_label_hits = [
        phrase
        for phrase in ["DRAFT", "PRIVATE", "INTERNAL", "scratch"]
        if phrase in "\n".join(pages)
    ]
    sensitive_or_old = [
        "student medical note",
        "vendor dispute",
        "staff disciplinary note",
        "personal phone",
        "legal settlement draft",
        "do not release",
        "internal scoring memo",
        "18,420",
        "Site B-12",
        "$9,800",
        "2026-05-03",
        "2026-06-30",
    ]
    hits = public_label_hits + [p for p in sensitive_or_old if normalize_text(p) in text]
    assert not hits, f"Forbidden phrases still present: {hits}"


def test_required_public_lines_by_page():
    pages = extract_pages(OUTPUT_PATH)
    all_text = "\n".join(pages)
    assert_terms(all_text, ["18,240", "Site B-21", "$980", "2026-05-13", "2026-06-15"], "correction facts")

    assert_terms(pages[0], ["Metro Schools", "Meal Program", "Public", "Audit", "School Year 2026"], "cover page")
    assert_any(pages[0], ["public report", "public audit", "public release"], "cover page public status")

    assert_terms(
        pages[1],
        ["Overall finding", "Substantial compliance", "Meals reconciled", "18,240", "Sites reviewed", "8", "4 of 5"],
        "executive findings page",
    )

    assert_terms(pages[2], ["Meal Type", "Breakfast", "7,880", "Lunch", "10,360", "Total meals", "18,240"], "meal reconciliation page")
    assert_any(pages[2], ["Reconciled", "Final", "Corrected"], "meal reconciliation status")

    assert_terms(pages[3], ["Site", "Site B-21", "Meal count controls", "Site C-04", "Temperature log", "Site D-09"], "site compliance page")
    assert_any(pages[3], ["Complete", "Open", "Corrected"], "site compliance status")

    assert_terms(pages[4], ["Vendor", "FreshStart", "$980", "Northline Dairy", "$0", "City Bakery"], "vendor invoice page")
    assert_any(pages[4], ["Corrected overage", "Overage", "Resolved", "Under Review"], "vendor invoice finding")

    assert_terms(pages[5], ["Allergen", "AL-001", "Sesame", "2026-05-13", "AL-002", "AL-003"], "allergen log page")
    assert_any(pages[5], ["Posted", "Sent", "Distributed"], "allergen log status")

    assert_terms(pages[6], ["CA-001", "Site B-21", "2026-06-15", "Nutrition Data Team", "CA-002"], "corrective action page")
    assert_any(pages[6], ["Open", "Complete"], "corrective action status")

    assert_terms(pages[7], ["OUT-001", "Families", "audit summary", "OUT-002", "Site"], "outreach page")
    assert_any(pages[7], ["Ready", "Complete", "Planned"], "outreach status")

    assert_terms(pages[8], ["Meal pattern compliance", "Civil rights notice", "Food safety logs", "Complete"], "nutrition checklist page")

    assert_terms(
        pages[9],
        ["Header band", "40", "1F2937", "0F766E", "166534", "7C2D12", "B91C1C"],
        "publication style guide page",
    )
    assert_terms(pages[10], ["Appendix A", "Meal count", "Appendix B", "Site review", "Release approval", "Nutrition Services Director"], "appendix register page")


def test_footer_and_page_numbering():
    pages = extract_pages(OUTPUT_PATH)
    for idx, text in enumerate(pages, start=1):
        page_text = normalized(text)
        assert normalized(EXPECT["footer"]) in page_text, f"Page {idx}: missing public footer"
        assert normalized(f"Page {idx} of {EXPECT['page_count']}") in page_text, f"Page {idx}: wrong page numbering"


def test_header_band_style_migration():
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        assert len(pdf.pages) == len(EXPECT["header_colors"])
        for idx, expected_hex in enumerate(EXPECT["header_colors"]):
            assert page_has_header_band(pdf.pages[idx], expected_hex), f"Page {idx + 1}: missing header band {expected_hex}"


def test_pdf_metadata_cleaned():
    reader = PdfReader(str(OUTPUT_PATH))
    info = reader.metadata
    metadata_text = "\n".join(str(value or "") for value in [info.title, info.subject, info.author, info.get("/Keywords")])
    assert_terms(metadata_text, ["Metro Schools", "Meal", "Audit"], "PDF metadata")
    assert_terms(metadata_text, ["Nutrition"], "PDF metadata")
    assert_any(metadata_text, ["public", "school year 2026", "2026"], "PDF metadata public scope")
    hits = [p for p in ["DRAFT", "PRIVATE", "INTERNAL", "student medical note", "vendor dispute", "personal phone"] if normalize_text(p) in normalize_text(metadata_text)]
    assert not hits, f"Metadata contains forbidden material: {hits}"


def test_source_artifact_was_not_modified():
    source_pages = extract_pages(INPUT_PATH)
    assert len(source_pages) > EXPECT["page_count"]
    text = normalize_text("\n".join(source_pages))
    assert "draft" in text
    assert "vendor dispute" in text
    assert "personal phone" in text
