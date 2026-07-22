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


def all_text(path):
    return "\n".join(extract_pages(path))


def norm(text):
    return re.sub(r"[^a-z0-9+$%.]+", " ", str(text).lower()).strip()


def compact(text):
    return re.sub(r"[^a-z0-9]+", "", str(text).lower())


def squash_repeated_chars(text):
    return re.sub(r"(.)\1+", r"\1", text)


def require_parts(text, parts, label):
    normalized = norm(text)
    compacted = compact(text)
    squashed = squash_repeated_chars(compacted)
    missing = []
    for part in parts:
        normalized_part = norm(part)
        tokens = [tok for tok in normalized_part.split() if tok]
        if normalized_part in normalized:
            continue
        if compact(part) and compact(part) in compacted:
            continue
        if compact(part) and compact(part) in squashed:
            continue
        if tokens and all(tok in normalized for tok in tokens):
            continue
        missing.append(part)
    assert not missing, f"{label}: missing semantic parts {missing!r}"


def forbidden_hits(text):
    cleaned = normalize_text(text)
    allowed_internal_contexts = [
        "internal only pages removed",
        "internal-only pages removed",
        "internal only material was removed",
        "internal-only material was removed",
        "internal only material removed",
        "internal-only material removed",
    ]
    for phrase in allowed_internal_contexts:
        cleaned = cleaned.replace(normalize_text(phrase), "")
    return [p for p in EXPECT["forbidden_phrases"] if normalize_text(p) in cleaned]


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


def color_family(hex_color):
    if not hex_color:
        return None
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    if b > g > r and b >= 70 and (b - r) >= 35:
        return "dark_blue"
    if r > g > b and r >= 160 and 45 <= g <= 125:
        return "terracotta"
    if r >= 170 and g >= 120 and b <= 90 and r >= g:
        return "gold"
    if g > r and g >= b and 70 <= g <= 140:
        return "green"
    return None


EXPECTED_HEADER_FAMILIES = [
    "dark_blue",
    "terracotta",
    "gold",
    "gold",
    "green",
    "green",
    "green",
    "dark_blue",
]


def page_has_header_family(page, expected_family):
    for rect in page.rects:
        if rect.get("top", 9999) <= 80 and rect.get("height", 0) >= 20:
            if color_family(color_to_hex(rect.get("non_stroking_color"))) == expected_family:
                return True
    return False


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    assert OUTPUT_PATH != INPUT_PATH


def test_page_count_titles_and_private_material_removed():
    pages = extract_pages(OUTPUT_PATH)
    assert len(pages) == EXPECT["page_count"]
    for idx, expected_title in enumerate(EXPECT["page_titles"]):
        first_text = "\n".join(pages[idx].splitlines()[:5])
        assert compact(expected_title) in compact(first_text), f"Page {idx + 1}: missing title {expected_title!r}"
    hits = forbidden_hits("\n".join(pages))
    assert not hits, f"Forbidden phrases still present: {hits}"


def test_required_public_lines_by_page():
    pages = extract_pages(OUTPUT_PATH)
    require_parts(pages[0], ["Public incident field guide", "2026-07-15", "public safety coordination"], "Page 1")
    quick_action_groups = [
        [["Confirm material identity", "placard", "shipping paper", "facility contact"]],
        [["Establish initial isolation", "rescue", "containment"]],
        [["public information officer", "community messages"], ["public information officer", "community information"]],
        [["unified command", "protective action distances", "one mile"]],
    ]
    for idx, alternatives in enumerate(quick_action_groups, start=1):
        failures = []
        for action in alternatives:
            try:
                require_parts(pages[1], action, f"Page 2 action {idx}")
                break
            except AssertionError as exc:
                failures.append(str(exc))
        else:
            raise AssertionError(f"Page 2 action {idx}: none of the accepted variants matched: {failures!r}")

    for item in [
        ["1005", "Ammonia, anhydrous", "125", "Toxic inhalation gas"],
        ["1017", "Chlorine", "124", "Toxic inhalation gas"],
        ["1203", "Gasoline", "128", "Flammable liquid"],
        ["1977", "Nitrogen", "refrigerated", "120", "Cryogenic gas"],
        ["1993", "Flammable liquids", "n.o.s.", "128", "Flammable liquid"],
    ]:
        require_parts(pages[2], item, "Page 3 substance index")

    for item in [
        ["Chlorine", "Small", "100 ft", "0.3 mi", "1.1 mi"],
        ["Chlorine", "Large", "800 ft", "1.7 mi", "4.2 mi"],
        ["Ammonia, anhydrous", "Small", "100 ft", "0.2 mi", "0.6 mi"],
        ["Gasoline", "Large", "150 ft", "0.3 mi", "0.5 mi"],
    ]:
        require_parts(pages[3], item, "Page 4 protective distances")

    for item in [
        ["Hot Zone", "trained responders", "assigned PPE"],
        ["Warm Zone", "Decontamination corridor", "exposure monitoring"],
        ["Cold Zone", "Command", "triage", "media", "family assistance"],
        ["Public Perimeter", "Community notification", "traffic control"],
    ]:
        require_parts(pages[4], item, "Page 5 isolation zones")

    for item in [
        ["Toxic inhalation gas", "Positive-pressure SCBA", "Level A or B", "Continuous air monitoring"],
        ["Flammable liquid", "SCBA", "vapor", "Structural turnout", "Level B", "Eliminate ignition sources"],
        ["Cryogenic gas", "Face shield", "insulated gloves", "Thermal protection", "Ventilate low areas"],
    ]:
        require_parts(pages[5], item, "Page 6 PPE matrix")

    for item in [
        ["Gross rinse", "medical evaluation", "contamination is visible"],
        ["Remove outer garments", "bag", "contaminated material"],
        ["Document", "person", "time", "substance", "symptoms"],
        ["Transfer clean patients", "triage", "exposure tag"],
    ]:
        require_parts(pages[6], item, "Page 7 decontamination workflow")

    require_parts(pages[7], ["Release owner", "Capstone Harbor Emergency Management"], "Page 8 release owner")
    require_parts(pages[7], ["Public release date", "2026-07-15"], "Page 8 release date")
    require_parts(pages[7], ["Correction", "applied", "Yes"], "Page 8 correction log")
    require_parts(pages[7], ["public", "pages", "remove"], "Page 8 nonpublic pages")


def test_footer_and_page_numbering():
    pages = extract_pages(OUTPUT_PATH)
    for idx, text in enumerate(pages, start=1):
        require_parts(text, ["Capstone Harbor"], f"Page {idx} footer")
        assert f"Page {idx}" in text or f"Page {idx}" in squash_repeated_chars(text), f"Page {idx}: wrong page numbering"


def test_header_band_style_migration():
    with pdfplumber.open(OUTPUT_PATH) as pdf:
        assert len(pdf.pages) == len(EXPECTED_HEADER_FAMILIES)
        for idx, expected_family in enumerate(EXPECTED_HEADER_FAMILIES):
            assert page_has_header_family(pdf.pages[idx], expected_family), f"Page {idx + 1}: missing {expected_family} header band"


def test_pdf_metadata_cleaned():
    reader = PdfReader(str(OUTPUT_PATH))
    info = reader.metadata
    metadata_text = " ".join(str(v or "") for v in [info.title, info.subject, info.author, info.get("/Keywords")])
    hits = forbidden_hits(metadata_text)
    assert not hits, f"Metadata still contains forbidden phrases: {hits}"
    require_parts(info.title or "", ["Capstone Harbor", "HazMat Field Guide"], "metadata title")
    require_parts(info.author or "", ["Capstone Harbor"], "metadata author")
    require_parts(metadata_text, ["public"], "metadata public marker")


def test_source_artifact_was_not_modified():
    source_pages = extract_pages(INPUT_PATH)
    assert len(source_pages) > EXPECT["page_count"]
    text = normalize_text("\n".join(source_pages))
    assert "draft" in text
    assert "chlorine plume rumor" in text
    assert "personal phone" in text
