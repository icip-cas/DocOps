import json
import os
import re
import sys
from pathlib import Path
from datetime import date, datetime

import pdfplumber
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).parent))
from verifier_utils import *  # noqa: F401,F403

META_PATH = Path(os.environ.get('TASK_METADATA_PATH', '/tests/task_metadata.json'))
META = json.loads(META_PATH.read_text(encoding='utf-8'))
INPUT_PATH = Path(os.environ.get('INPUT_PATH', META['input_path']))
OUTPUT_PATH = Path(os.environ.get('OUTPUT_PATH', META['output_path']))
EXPECT = META['verifier_expectations']


def _vopt_norm_text(value):
    if value is None:
        return ''
    if isinstance(value, datetime):
        if value.hour or value.minute or value.second:
            return value.strftime('%Y-%m-%d %H:%M')
        return value.strftime('%Y-%m-%d')
    if isinstance(value, date):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value)
    text = text.replace('\u2013', '-').replace('\u2014', '-').replace('\u2212', '-')
    text = text.replace('\xa0', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _vopt_norm_formula(value):
    return re.sub(r'\s+', '', _vopt_norm_text(value)).upper()


def _vopt_assert_formula(actual, expected, label='formula'):
    assert _vopt_norm_formula(actual) == _vopt_norm_formula(expected), f"{label}: expected {expected!r}, found {actual!r}"


def _vopt_norm_rows(rows):
    return [[_vopt_norm_text(cell) for cell in row] for row in rows]


def _vopt_rows_equal(actual, expected):
    return _vopt_norm_rows(actual) == _vopt_norm_rows(expected)


def _vopt_rows_text(rows):
    return '\n'.join('|'.join(_vopt_norm_text(cell) for cell in row) for row in rows)


def _vopt_header_has(header, required):
    actual = {_vopt_norm_text(cell).lower() for cell in header}
    missing = [cell for cell in required if _vopt_norm_text(cell).lower() not in actual]
    assert not missing, f"Missing expected header columns {missing!r}; actual={header!r}"


def _vopt_section_text(section_part):
    return '\n'.join(p.text for p in getattr(section_part, 'paragraphs', []) if p.text is not None)


def _vopt_header_text(doc):
    return '\n'.join(_vopt_section_text(section.header) for section in doc.sections)


def _vopt_footer_text(doc):
    return '\n'.join(_vopt_section_text(section.footer) for section in doc.sections)


def _vopt_table_ref(ws, expected_name=None):
    tables = ws.tables
    if expected_name:
        for name in tables.keys():
            if str(name).lower() == str(expected_name).lower():
                return tables[name].ref
    vals = list(tables.values())
    assert vals, f"No Excel native table found on sheet {ws.title!r}"
    return vals[0].ref


def _vopt_clean_area(value):
    text = _vopt_norm_text(value).replace("'", '').replace('$', '').replace(' ', '')
    if '!' in text:
        text = text.split('!', 1)[1]
    return text.upper()


def _vopt_assert_print_area(ws, expected):
    actual = _vopt_clean_area(ws.print_area)
    target = _vopt_clean_area(expected)
    assert target in actual or actual in target, f"{ws.title}: expected print area {expected!r}, found {ws.print_area!r}"


def _vopt_freeze_pane(value):
    return getattr(value, 'coordinate', value)


def _vopt_validated_cells(ws):
    cells = set()
    for dv in ws.data_validations.dataValidation:
        for rng in dv.cells.ranges:
            text = str(rng).replace('$', '')
            try:
                min_col, min_row, max_col, max_row = range_boundaries(text)
            except ValueError:
                cells.add(text)
                continue
            if max_row > 10000:
                max_row = max(min_row, 200)
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    cells.add(f"{get_column_letter(col)}{row}")
    return cells


class _VoptValidationRanges(list):
    def __init__(self, ws, ranges):
        super().__init__(ranges)
        self.ws = ws

    def __contains__(self, expected):
        expected = str(expected).replace('$', '')
        if super().__contains__(expected):
            return True
        try:
            min_col, min_row, max_col, max_row = range_boundaries(expected)
        except ValueError:
            return super().__contains__(expected)
        target = {f"{get_column_letter(col)}{row}" for row in range(min_row, max_row + 1) for col in range(min_col, max_col + 1)}
        return target.issubset(_vopt_validated_cells(self.ws))


def _vopt_dv_ranges(ws):
    ranges = []
    for dv in ws.data_validations.dataValidation:
        ranges.extend(str(rng).replace('$', '') for rng in dv.cells.ranges)
    return _VoptValidationRanges(ws, ranges)


def _vopt_ordered_subset(expected, actual):
    actual_iter = iter([_vopt_norm_text(item) for item in actual])
    for item in [_vopt_norm_text(value) for value in expected]:
        for candidate in actual_iter:
            if item == candidate:
                break
        else:
            return False
    return True



def _page_texts(path):
    with pdfplumber.open(str(path)) as pdf:
        return [page.extract_text() or '' for page in pdf.pages]


def _page_titles(page_texts):
    titles = []
    for text in page_texts:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        titles.append(lines[0] if lines else '')
    return titles


def _assert_required_pdf_semantics(full_text):
    required = [
        'Status Final Client Release',
        'Page Count 4',
        'Invoice ID INV-1001',
        'Amount $12,400',
        'Invoice ID INV-1002',
        'Amount $8,750',
        'Federal Tax ID [REDACTED TAX ID]',
        'Bank Routing [REDACTED BANK ROUTING]',
        'Appendix ID REM-2026-07',
    ]
    require_all(full_text, required, 'PDF text')


def _assert_forbidden_pdf_semantics(full_text):
    norm = normalize_text(full_text)
    forbidden = [
        'duplicate do not release',
        'duplicate copy',
        'internal only',
        'hidden staging batch',
        'status draft',
        'page count 5',
        'contains duplicate invoice page and internal routing page',
        '12-3456789',
        '98-7654321',
        '021000021',
        '011000015',
    ]
    hits = [phrase for phrase in forbidden if phrase in norm]
    assert not hits, f'PDF text: forbidden draft/internal/sensitive content remains: {hits}'


def _compile_pattern(pattern):
    try:
        if '\\\\b' in pattern:
            pattern = pattern.encode('utf-8').decode('unicode_escape')
    except Exception:
        pass
    return re.compile(pattern)


def _outline_has(outline_titles, expected):
    expected_norm = normalize_text(expected)
    title_norms = [normalize_text(title) for title in outline_titles]
    if expected_norm in title_norms:
        return True
    aliases = {
        'release cover': ['cover', 'invoice release packet'],
        'invoices': ['invoices'],
        'alpine labs inv-1001': ['invoice - alpine labs inv-1001', 'alpine labs invoice inv-1001', 'alpine labs inv-1001'],
        'birch clinic inv-1002': ['invoice - birch clinic inv-1002', 'birch clinic invoice inv-1002', 'birch clinic inv-1002'],
        'approved remittance appendix': ['approved remittance appendix'],
    }
    options = aliases.get(expected_norm, [expected_norm])
    if expected_norm == 'invoices':
        return (
            any(option in title_norms for option in options)
            or _outline_has(outline_titles, 'Alpine Labs INV-1001') and _outline_has(outline_titles, 'Birch Clinic INV-1002')
        )
    return any(option in title_norms for option in options)


def test_output_exists_and_is_pdf():
    run_preflight(META, INPUT_PATH, OUTPUT_PATH)
    reader = PdfReader(str(OUTPUT_PATH))
    assert len(reader.pages) == EXPECT['page_count'], f"Expected {EXPECT['page_count']} pages, found {len(reader.pages)}"


def test_page_order_and_text_anchors():
    texts = _page_texts(OUTPUT_PATH)
    for expected_title, page_text in zip(EXPECT['page_titles'], texts):
        assert expected_title in page_text, f"Missing page title {expected_title!r} on its expected page"
    full_text = '\n'.join(texts)
    _assert_required_pdf_semantics(full_text)
    _assert_forbidden_pdf_semantics(full_text)
    for pattern in EXPECT['sensitive_patterns_absent']:
        assert not _compile_pattern(pattern).search(full_text), f"Sensitive pattern still present: {pattern}"


def test_bookmark_outline():
    reader = PdfReader(str(OUTPUT_PATH))
    outline_titles = [title for _level, title in flatten_outline(reader.outline)]
    for expected in EXPECT['outline_titles']:
        assert _outline_has(outline_titles, expected), f"Missing outline title/category: {expected}; found {outline_titles!r}"
    assert "Alpine Labs INV-1001 Duplicate" not in outline_titles
    assert "Internal Payment Routing" not in outline_titles


def test_seed_would_not_pass_core_requirements():
    seed_text = '\n'.join(_page_texts(INPUT_PATH))
    assert "Duplicate copy" in seed_text
    assert "INTERNAL ONLY" in seed_text
    assert "12-3456789" in seed_text
