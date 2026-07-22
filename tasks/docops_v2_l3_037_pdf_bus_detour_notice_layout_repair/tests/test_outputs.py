import json, os, re
from pathlib import Path
from pypdf import PdfReader
from verifier_utils import normalize_text, pdf_page_texts, pdf_theme_rect_colors, run_preflight
META_PATH=Path(os.environ.get('TASK_METADATA_PATH','/tests/task_metadata.json'))
if not META_PATH.exists(): META_PATH=Path(__file__).parent/'task_metadata.json'
META=json.loads(META_PATH.read_text()); EXPECT=META['verifier_expectations']; INPUT_PATH=Path(os.environ.get('INPUT_PATH',META['input_path'])); OUTPUT_PATH=Path(os.environ.get('OUTPUT_PATH',META['output_path']))
def metadata_text(reader): return '\n'.join(str(v or '') for v in (reader.metadata or {}).values())
def flatten_outline(outline):
    out=[]
    for item in outline:
        if isinstance(item,list): out+=flatten_outline(item)
        else: out.append(str(getattr(item,'title',item.get('/Title',item))))
    return out
def all_text(path): return '\n'.join(pdf_page_texts(path))
def require_all(text, phrases, label):
    nt=normalize_text(text); missing=[p for p in phrases if normalize_text(p) not in nt]; assert not missing, f'{label} missing {missing}'
def require_any(text, phrases, label):
    nt=normalize_text(text); assert any(normalize_text(p) in nt for p in phrases), f'{label} expected one of {phrases}'
def cleanup_context(text, phrase):
    pattern=re.compile(rf'\\b{re.escape(phrase)}\\b', re.I)
    for m in pattern.finditer(text):
        window=text[max(0,m.start()-120):m.end()+180]
        if not re.search(r'\\b(no|not|removed|cleaned|excluded|without|superseded|corrected|replaced)\\b', window, re.I):
            return False
    return True
def assert_no_forbidden(text):
    nt=normalize_text(text)
    hard=['Route 21','2026-10-13','2026-10-27','7 stops','30 minutes','operator personal phone','do not post','duplicate draft']
    hits=[p for p in hard if normalize_text(p) in nt]; assert not hits, hits
    contextual=[p for p in ['DRAFT','PRIVATE'] if re.search(rf'\\b{re.escape(p)}\\b', text, re.I) and not cleanup_context(text,p)]
    assert not contextual, contextual
def test_output_exists_and_is_pdf(): run_preflight(META,INPUT_PATH,OUTPUT_PATH); assert OUTPUT_PATH!=INPUT_PATH
def test_pages_text_tables_and_forbidden_removed():
    reader=PdfReader(str(OUTPUT_PATH)); assert len(reader.pages)==EXPECT['page_count']; pages=pdf_page_texts(OUTPUT_PATH)
    for i,t in enumerate(EXPECT['page_titles']): assert t in pages[i]
    require_all(pages[0], ['Route 12','311 Transit Desk'], 'title page')
    require_all(pages[0], ['Public','detour'], 'title page')
    require_any(pages[1], ['2026-10-03','October 3','Oct. 3','Oct 3'], 'detour summary')
    require_any(pages[1], ['2026-10-17','October 17','Oct. 17','Oct 17'], 'detour summary')
    require_any(pages[1], ['9','nine'], 'detour summary')
    require_all(pages[2], ['Shuttle','15 minutes'], 'stop matrix')
    require_any(pages[2], ['Stop ID','Stop','Closed stop'], 'stop matrix')
    require_any(pages[2], ['Closure Status','Closed'], 'stop matrix')
    require_any(pages[3], ['2026-10-03','October 3','Oct. 3','Oct 3'], 'communication timeline')
    require_any(pages[3], ['Public alert','alert','communication','rider'], 'communication timeline')
    require_all(pages[4], ['Release Checklist'], 'release checklist')
    require_any(pages[4], ['ready','complete','release','public'], 'release checklist')
    joined='\n'.join(pages)
    require_all(joined, ['Route 12','2026-10-03','2026-10-17','9','15 minutes','311 Transit Desk'], 'corrected public values')
    assert_no_forbidden(joined+'\n'+metadata_text(reader))
def test_outline_metadata_style_and_orientation():
    reader=PdfReader(str(OUTPUT_PATH)); assert flatten_outline(reader.outline)==EXPECT['outline_titles']; meta=reader.metadata or {}
    assert 'Public Bus Detour Notice Packet' in str(meta.get('/Title',''))
    assert str(meta.get('/Author','')).strip()
    assert str(meta.get('/Subject','')).strip()
    assert_no_forbidden(metadata_text(reader))
    for idx,page in enumerate(reader.pages):
        w,h=float(page.mediabox.width),float(page.mediabox.height)
        if idx in EXPECT['landscape_pages']:
            assert abs(w-EXPECT['style']['landscape_width'])<1 and abs(h-EXPECT['style']['landscape_height'])<1
        else:
            assert abs(w-EXPECT['style']['portrait_width'])<1 and abs(h-EXPECT['style']['portrait_height'])<1
    colors=pdf_theme_rect_colors(OUTPUT_PATH); assert colors, 'PDF should contain styled colored rectangles'
    for page_text in pdf_page_texts(OUTPUT_PATH):
        assert 'Page' in page_text and ('Route 12' in page_text or 'Public Bus Detour Notice Packet' in page_text or '311 Transit Desk' in page_text)
def test_source_artifact_was_not_modified():
    reader=PdfReader(str(INPUT_PATH)); text=normalize_text(all_text(INPUT_PATH)+'\n'+metadata_text(reader)); missing=[p for p in EXPECT['source_must_contain'] if normalize_text(p) not in text]; assert not missing, missing
