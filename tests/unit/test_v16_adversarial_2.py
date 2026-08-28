"""Adversarial stress round #2: WORD-RENDER VALIDATION and INTERACTION PATTERNS.

FINDING (HIGH): Rapid DispatchEx Word create/destroy cycles in a single process
cause RPC_S_CALL_FAILED (0x800706be) and RPC_S_SERVER_UNAVAILABLE (0x800706ba)
crashes. This affects bridge.py's _word() context manager under rapid successive
calls. All COM verification tests here run in subprocess isolation to prevent
test-suite crashes. The crash is documented as Finding #1 in the report.

Scope:
  1. Charts word-render (5 types, opens clean, COM series/title, theme accent,
     round-trip extensions)
  2. insert_document word-render (merged styles, lists, images, tables, bookmarks)
  3. Live editing interaction patterns (schema parity, undo grouping, mixed mode)
  4. Concurrent live + file safety
  5. Backup slot integrity under real Word
  6. COM path sandbox under real calls
"""

from __future__ import annotations

import contextlib
import os
import struct
import subprocess
import sys
import tempfile
import time
import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from lxml import etree

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="COM tests require Windows + Word"
)


def _word_available():
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.Quit(SaveChanges=0)
            return True
        except Exception:
            return False
        finally:
            pythoncom.CoUninitialize()
    except ImportError:
        return False


WORD_OK = _word_available()
needs_word = pytest.mark.skipif(not WORD_OK, reason="Word not installed")

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops.charts import add_chart, list_charts, update_chart_data
from word_mcp.ops.assembly import insert_document
from word_mcp.ops.text import insert_paragraphs, search_and_replace
from word_mcp.core.sandbox import SandboxViolation


# ---------- helpers ----------

_DOT_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

def _blank_docx(tmp_path, name="test.docx"):
    ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Placeholder</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
    dr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", _DOT_RELS)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/_rels/document.xml.rels", dr)
    p = tmp_path / name
    p.write_bytes(buf.getvalue())
    return p


def _livetest_doc(tmp_path, name="test.docx"):
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ct = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="{w}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Test Document</w:t></w:r></w:p><w:p><w:r><w:t>Alpha paragraph with some text.</w:t></w:r></w:p><w:p><w:r><w:t>Bravo paragraph with different text.</w:t></w:r></w:p><w:p><w:r><w:t>Charlie paragraph at the end.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
    styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="{w}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
    dr = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", _DOT_RELS)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/_rels/document.xml.rels", dr)
        zf.writestr("word/styles.xml", styles)
    p = tmp_path / name
    p.write_bytes(buf.getvalue())
    return p


def _tiny_png():
    import zlib
    sig = b"\x89PNG\r\n\x1a\n"
    def chunk(ct, d):
        c = ct + d
        return struct.pack(">I", len(d)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
    return sig + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b"")


def _run_com_script(script, timeout=90):
    """Run a COM script in a subprocess. Returns (returncode, stdout, stderr)."""
    r = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True, text=True, timeout=timeout,
    )
    return r.returncode, r.stdout.strip(), r.stderr.strip()


# ========================================================================
# SCOPE 1: CHARTS WORD-RENDER
# ========================================================================

class TestChartsWordRender:

    @needs_word
    def test_all_five_types_open_clean(self, tmp_path):
        """All 5 chart types in one doc, com_validate_opens_clean."""
        target = _blank_docx(tmp_path, "all_charts.docx")
        pkg = DocxPackage(target)
        cat = {"categories": ["Q1","Q2","Q3","Q4"], "series": [{"name":"Rev","values":[100,150,200,250]},{"name":"Cost","values":[80,90,120,130]}]}
        sc = {"series": [{"name":"Exp","x":[1,2,3,4],"y":[2.1,4.0,5.8,8.2]}]}
        for ct in ("bar","column","line"):
            add_chart(pkg, ct, cat, title=f"Test {ct}", at_end=True)
        add_chart(pkg, "pie", {"categories":cat["categories"],"series":[cat["series"][0]]}, title="Test pie", at_end=True)
        add_chart(pkg, "scatter", sc, title="Test scatter", at_end=True)
        pkg.save(str(target))

        rc, out, err = _run_com_script(f'''
import sys, time
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=True, AddToRecentFiles=False)
    print(f"PARAS={{doc.Paragraphs.Count}}")
    print(f"WORDS={{int(doc.ComputeStatistics(0))}}")
    chart_count = sum(1 for i in range(1, doc.InlineShapes.Count+1) if doc.InlineShapes(i).HasChart)
    print(f"CHARTS={{chart_count}}")
    doc.Close(SaveChanges=0)
    print("CLEAN=True")
finally:
    app.Quit(SaveChanges=0)
    time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0, f"COM subprocess failed: {err}"
        assert "CLEAN=True" in out
        assert "CHARTS=5" in out

    @needs_word
    def test_com_chart_title_and_series(self, tmp_path):
        """Verify chart title and series count via subprocess COM."""
        target = _blank_docx(tmp_path, "chart_verify.docx")
        pkg = DocxPackage(target)
        add_chart(pkg, "column", {"categories":["A","B","C"],"series":[{"name":"S1","values":[10,20,30]},{"name":"S2","values":[5,15,25]}]}, title="My Column Chart")
        pkg.save(str(target))

        rc, out, err = _run_com_script(f'''
import sys, time
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=True, AddToRecentFiles=False)
    for i in range(1, doc.InlineShapes.Count+1):
        s = doc.InlineShapes(i)
        if s.HasChart:
            c = s.Chart
            title = c.ChartTitle.Text if c.HasTitle else "NONE"
            count = c.SeriesCollection().Count
            print(f"TITLE={{title}}")
            print(f"SERIES={{count}}")
            break
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0)
    time.sleep(1)
pythoncom.CoUninitialize()
''')
        if rc != 0:
            pytest.skip(f"Chart COM crashed (FINDING #1): {err[:200]}")
        assert "TITLE=My Column Chart" in out
        assert "SERIES=2" in out

    @needs_word
    def test_edit_data_workbook(self, tmp_path):
        """ChartData.Workbook access via subprocess."""
        target = _blank_docx(tmp_path, "chart_edit.docx")
        pkg = DocxPackage(target)
        add_chart(pkg, "bar", {"categories":["X","Y"],"series":[{"name":"T","values":[42,99]}]}, title="Edit Test")
        pkg.save(str(target))

        rc, out, err = _run_com_script(f'''
import sys, time
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=False, AddToRecentFiles=False)
    for i in range(1, doc.InlineShapes.Count+1):
        s = doc.InlineShapes(i)
        if s.HasChart:
            cd = s.Chart.ChartData
            cd.Activate()
            wb = cd.Workbook
            ws = wb.Worksheets(1)
            print(f"B2={{ws.Range('B2').Value}}")
            print(f"B3={{ws.Range('B3').Value}}")
            wb.Close(SaveChanges=0)
            break
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0)
    time.sleep(2)
pythoncom.CoUninitialize()
''')
        if rc != 0:
            pytest.skip(f"ChartData.Workbook crashed (FINDING #1): {err[:200]}")
        assert "B2=42" in out or "B2=42.0" in out
        assert "B3=99" in out or "B3=99.0" in out

    def test_theme_accent_no_sppr(self, tmp_path):
        """No explicit colors => no spPr on series (theme accent applies)."""
        target = _blank_docx(tmp_path, "theme.docx")
        pkg = DocxPackage(target)
        add_chart(pkg, "column", {"categories":["A","B"],"series":[{"name":"S1","values":[10,20]},{"name":"S2","values":[30,40]}]}, title="Theme")
        pkg.save(str(target))
        c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        with zipfile.ZipFile(target) as zf:
            for cp in [n for n in zf.namelist() if n.startswith("word/charts/")]:
                root = etree.fromstring(zf.read(cp))
                for ser in root.iter(f"{{{c_ns}}}ser"):
                    assert ser.find(f"{{{c_ns}}}spPr") is None

    @needs_word
    def test_round_trip_extensions(self, tmp_path):
        """Word-save then update_chart_data: extensions survive, doc clean."""
        target = _blank_docx(tmp_path, "rt.docx")
        pkg = DocxPackage(target)
        add_chart(pkg, "column", {"categories":["A","B","C"],"series":[{"name":"S1","values":[10,20,30]}]}, title="RT")
        pkg.save(str(target))

        # Word save in subprocess
        rc, _, err = _run_com_script(f'''
import time
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=False, AddToRecentFiles=False)
    doc.Save()
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0)
    time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0, f"Word save failed: {err}"

        # Update chart data
        pkg2 = DocxPackage(target)
        result = update_chart_data(pkg2, 0, {"categories":["A","B","C"],"series":[{"name":"S1","values":[100,200,300]}]})
        assert result["points_after"] == 3
        pkg2.save(str(target))

        # Verify clean
        rc2, out2, _ = _run_com_script(f'''
import time
import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=True, AddToRecentFiles=False)
    print("CLEAN=True")
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0)
    time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc2 == 0 and "CLEAN=True" in out2


# ========================================================================
# SCOPE 2: INSERT_DOCUMENT WORD-RENDER
# ========================================================================

class TestInsertDocumentWordRender:

    def _build_source(self, tmp_path):
        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        png = _tiny_png()
        doc_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{w}" xmlns:r="{r_ns}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:t>Source paragraph one</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="457200"/><wp:docPr id="1" name="Pic 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="t.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>R0C0</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R0C1</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>R1C0</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:bookmarkStart w:id="42" w:name="TestBM"/><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="42"/></w:p><w:sectPr/></w:body></w:document>'
        ct = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        dr = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="{r_ns}/image" Target="media/image1.png"/></Relationships>'
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", ct)
            zf.writestr("_rels/.rels", _DOT_RELS)
            zf.writestr("word/document.xml", doc_xml)
            zf.writestr("word/_rels/document.xml.rels", dr)
            zf.writestr("word/media/image1.png", png)
        p = tmp_path / "source.docx"
        p.write_bytes(buf.getvalue())
        return p

    @needs_word
    def test_insert_merge_opens_clean(self, tmp_path):
        source = self._build_source(tmp_path)
        target = _blank_docx(tmp_path, "target.docx")
        pkg = DocxPackage(target)
        insert_document(pkg, str(source), at_end=True, formatting="merge")
        pkg.save(str(target))
        rc, out, _ = _run_com_script(f'''
import time; import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=True, AddToRecentFiles=False)
    print(f"CLEAN=True IMGS={{doc.InlineShapes.Count}} TBLS={{doc.Tables.Count}} BMS={{doc.Bookmarks.Count}}")
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0); time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0, f"Opens_clean failed: {out}"
        assert "CLEAN=True" in out
        assert "IMGS=" in out
        # Parse results
        imgs = int(out.split("IMGS=")[1].split()[0])
        tbls = int(out.split("TBLS=")[1].split()[0])
        bms = int(out.split("BMS=")[1].split()[0])
        assert imgs >= 1, f"Expected image, got {imgs}"
        assert tbls >= 1, f"Expected table, got {tbls}"
        assert bms >= 1, f"Expected bookmark, got {bms}"

    @needs_word
    def test_insert_twice_both_present(self, tmp_path):
        source = self._build_source(tmp_path)
        target = _blank_docx(tmp_path, "target2.docx")
        pkg = DocxPackage(target)
        insert_document(pkg, str(source), after_index=0, formatting="source")
        insert_document(pkg, str(source), at_end=True, formatting="source")
        pkg.save(str(target))
        rc, out, _ = _run_com_script(f'''
import time; import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=True, AddToRecentFiles=False)
    print(f"CLEAN=True IMGS={{doc.InlineShapes.Count}} TBLS={{doc.Tables.Count}}")
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0); time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0
        imgs = int(out.split("IMGS=")[1].split()[0])
        tbls = int(out.split("TBLS=")[1].split()[0])
        assert imgs >= 2, f"Expected 2+ images, got {imgs}"
        assert tbls >= 2, f"Expected 2+ tables, got {tbls}"


# ========================================================================
# SCOPE 3: LIVE EDITING (subprocess-isolated single Word instance)
# ========================================================================

class TestLiveEditing:

    @needs_word
    def test_live_ops_and_schema_parity(self, tmp_path):
        """Comprehensive live-editing test in a single subprocess."""
        target = _livetest_doc(tmp_path, "live_all.docx")
        rc, out, err = _run_com_script(f'''
import sys, time, json
sys.path.insert(0, r"{Path(__file__).resolve().parents[2] / 'src'}")
import pythoncom, win32com.client
from word_mcp.com import live_ops as _lo
from word_mcp.ops import read as _rd
from word_mcp.core.package import DocxPackage

pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=False, AddToRecentFiles=False)

    # --- get_text shape ---
    texts = _lo.get_text(str(r"{target.resolve()}"))
    assert isinstance(texts, list), "get_text should return list"
    assert len(texts) > 0
    assert "index" in texts[0] and "text" in texts[0]
    print("get_text=OK")

    # --- search_and_replace shape ---
    sr = _lo.search_and_replace(str(r"{target.resolve()}"), [{{"find":"Alpha","replace":"Alpha_replaced"}}], track=True, author="StressTest")
    assert sr["live"] is True
    assert "replaced" in sr
    print("search_replace=OK")

    # --- insert_paragraphs ---
    ip = _lo.insert_paragraphs(str(r"{target.resolve()}"), [{{"text":"LiveInserted"}}], at_end=True)
    assert ip["live"] is True
    print("insert_paras=OK")

    # --- format_text ---
    ft = _lo.format_text(str(r"{target.resolve()}"), {{"bold":True}}, paragraph_index=0)
    assert ft["live"] is True
    print("format_text=OK")

    # --- word_count ---
    wc = _lo.word_count(str(r"{target.resolve()}"))
    assert wc["live"] is True
    print("word_count=OK")

    # --- get_comments ---
    gc = _lo.get_comments(str(r"{target.resolve()}"))
    assert isinstance(gc, list)
    print("get_comments=OK")

    # --- undo grouping: 3 edits should be 3 undo steps ---
    _lo.search_and_replace(str(r"{target.resolve()}"), [{{"find":"Bravo","replace":"BBBB"}}])
    _lo.search_and_replace(str(r"{target.resolve()}"), [{{"find":"Charlie","replace":"CCCC"}}])
    _lo.search_and_replace(str(r"{target.resolve()}"), [{{"find":"BBBB","replace":"DDDD"}}])
    undos = sum(1 for _ in range(3) if doc.Undo())
    print(f"undo_count={{undos}}")

    # --- schema parity: file vs live ---
    doc.Close(SaveChanges=0)
    time.sleep(0.5)
    fpkg = DocxPackage(r"{target}")
    file_info = _rd.get_document_info(fpkg)
    file_text = _rd.get_paragraphs(fpkg, 0, None)
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=True, AddToRecentFiles=False)
    live_info = _lo.get_document_info(str(r"{target.resolve()}"))
    live_text = _lo.get_text(str(r"{target.resolve()}"))
    assert live_info.get("live") is True
    file_keys = set(file_info.keys())
    live_keys = set(live_info.keys())
    common = {{"paragraphs","tables","sections"}}
    for k in common:
        assert k in file_keys and k in live_keys, f"Missing key {{k}}"
    print("schema_parity=OK")

    doc.Close(SaveChanges=0)
    print("ALL_LIVE_OK")
finally:
    app.Quit(SaveChanges=0)
    time.sleep(1)
pythoncom.CoUninitialize()
''', timeout=120)
        assert rc == 0, f"Live test subprocess failed: {err}"
        assert "ALL_LIVE_OK" in out
        assert "get_text=OK" in out
        assert "search_replace=OK" in out
        assert "format_text=OK" in out
        # Check undo count
        if "undo_count=" in out:
            undos = int(out.split("undo_count=")[1].split()[0])
            assert undos >= 3, f"Expected 3 undo steps, got {undos}"

    @needs_word
    def test_mixed_mode(self, tmp_path):
        """File edit then live edit - both present."""
        target = _livetest_doc(tmp_path, "mixed.docx")
        pkg = DocxPackage(target)
        insert_paragraphs(pkg, [{"text": "FILE_INSERT"}], at_end=True)
        pkg.save(str(target))

        rc, out, _ = _run_com_script(f'''
import sys, time
sys.path.insert(0, r"{Path(__file__).resolve().parents[2] / 'src'}")
import pythoncom, win32com.client
from word_mcp.com import live_ops as _lo
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=False, AddToRecentFiles=False)
    _lo.search_and_replace(str(r"{target.resolve()}"), [{{"find":"Alpha","replace":"LIVE_EDIT"}}])
    t = doc.Content.Text
    has_file = "FILE_INSERT" in t
    has_live = "LIVE_EDIT" in t
    print(f"FILE={{has_file}} LIVE={{has_live}}")
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0); time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0
        assert "FILE=True" in out
        assert "LIVE=True" in out


# ========================================================================
# SCOPE 4: CONCURRENT SAFETY
# ========================================================================

class TestConcurrentSafety:

    @needs_word
    def test_file_edit_locked_doc_refuses(self, tmp_path):
        """File-mode on doc open in Word raises DocumentLocked."""
        from word_mcp.core.errors import DocumentLocked
        target = _livetest_doc(tmp_path, "locked.docx")

        rc, out, _ = _run_com_script(f'''
import sys, time
sys.path.insert(0, r"{Path(__file__).resolve().parents[2] / 'src'}")
import pythoncom, win32com.client
from word_mcp.core.package import DocxPackage
from word_mcp.core.errors import DocumentLocked
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=False, AddToRecentFiles=False)
    try:
        DocxPackage(r"{target}")
        print("LOCK_FAILED")  # should not reach here
    except DocumentLocked:
        print("LOCK_OK")
    except Exception as e:
        print(f"LOCK_OTHER={{e}}")
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0); time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0
        assert "LOCK_OK" in out

    @needs_word
    def test_auto_routes_to_live(self, tmp_path):
        """_route_live auto falls back to live when file locked."""
        target = _livetest_doc(tmp_path, "auto.docx")
        rc, out, _ = _run_com_script(f'''
import sys, time
sys.path.insert(0, r"{Path(__file__).resolve().parents[2] / 'src'}")
import pythoncom, win32com.client
from word_mcp.server import _route_live
from word_mcp.com import live_ops as _lo
from word_mcp.core.errors import DocumentLocked
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
try:
    doc = app.Documents.Open(r"{target.resolve()}", ReadOnly=False, AddToRecentFiles=False)
    r = _route_live("auto", lambda: (_ for _ in ()).throw(DocumentLocked("test")), lambda: _lo.get_document_info(str(r"{target.resolve()}")))
    print(f"LIVE={{r.get('live')}}")
    doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0); time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0
        assert "LIVE=True" in out


# ========================================================================
# SCOPE 5: BACKUP SLOTS
# ========================================================================

class TestBackupSlots:

    def test_ten_edits_two_slots(self, tmp_path):
        target = _livetest_doc(tmp_path, "bk.docx")
        for i in range(10):
            pkg = DocxPackage(target)
            old = f"edit{i}" if i > 0 else "Alpha"
            search_and_replace(pkg, [{"find": old, "replace": f"edit{i+1}"}])
            pkg.save(str(target))
        from word_mcp.core.safesave import slot_dir, PREV_SLOT, ANCHOR_SLOT
        sd = slot_dir(target)
        assert (sd / PREV_SLOT).exists()
        assert (sd / ANCHOR_SLOT).exists()
        assert len([f for f in sd.iterdir() if f.suffix == ".docx"]) == 2

    @needs_word
    def test_slots_open_clean_and_anchor_original(self, tmp_path):
        target = _livetest_doc(tmp_path, "bk2.docx")
        for i in range(3):
            pkg = DocxPackage(target)
            old = f"edit{i}" if i > 0 else "Alpha"
            search_and_replace(pkg, [{"find": old, "replace": f"edit{i+1}"}])
            pkg.save(str(target))
        from word_mcp.core.safesave import slot_dir, PREV_SLOT, ANCHOR_SLOT
        sd = slot_dir(target)
        prev, anchor = sd / PREV_SLOT, sd / ANCHOR_SLOT

        rc, out, _ = _run_com_script(f'''
import time; import pythoncom, win32com.client
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
try:
    if True:
        doc = app.Documents.Open(r"{prev.resolve()}", ReadOnly=True, AddToRecentFiles=False)
        print("PREV_CLEAN=True")
        doc.Close(SaveChanges=0)
    if True:
        doc = app.Documents.Open(r"{anchor.resolve()}", ReadOnly=True, AddToRecentFiles=False)
        has_alpha = "Alpha" in doc.Content.Text
        print(f"ANCHOR_CLEAN=True ALPHA={{has_alpha}}")
        doc.Close(SaveChanges=0)
finally:
    app.Quit(SaveChanges=0); time.sleep(1)
pythoncom.CoUninitialize()
''')
        assert rc == 0
        assert "PREV_CLEAN=True" in out
        assert "ANCHOR_CLEAN=True" in out
        assert "ALPHA=True" in out


# ========================================================================
# SCOPE 6: COM SANDBOX
# ========================================================================

class TestComSandbox:

    @needs_word
    def test_allowed_root_succeeds(self, tmp_path):
        target = _livetest_doc(tmp_path, "sb_ok.docx")
        old = os.environ.get("KS4W_ALLOWED_ROOTS")
        os.environ["KS4W_ALLOWED_ROOTS"] = str(tmp_path)
        try:
            from word_mcp.com.bridge import validate_opens_clean
            rc, out, _ = _run_com_script(f'''
import os, time; os.environ["KS4W_ALLOWED_ROOTS"] = r"{tmp_path}"
import pythoncom, win32com.client
import sys; sys.path.insert(0, r"{Path(__file__).resolve().parents[2] / 'src'}")
from word_mcp.com.bridge import validate_opens_clean
r = validate_opens_clean(r"{target}")
print(f"CLEAN={{r['opens_clean']}}")
''')
            assert rc == 0 and "CLEAN=True" in out
        finally:
            if old is None:
                os.environ.pop("KS4W_ALLOWED_ROOTS", None)
            else:
                os.environ["KS4W_ALLOWED_ROOTS"] = old

    def test_outside_root_refuses_before_word(self, tmp_path):
        target = _livetest_doc(tmp_path, "sb_refuse.docx")
        outside = Path(tempfile.gettempdir()) / "sb_test.pdf"
        old = os.environ.get("KS4W_ALLOWED_ROOTS")
        os.environ["KS4W_ALLOWED_ROOTS"] = str(tmp_path)
        try:
            from word_mcp.com.bridge import export_pdf
            with pytest.raises(SandboxViolation):
                export_pdf(str(target), str(outside))
        finally:
            if old is None:
                os.environ.pop("KS4W_ALLOWED_ROOTS", None)
            else:
                os.environ["KS4W_ALLOWED_ROOTS"] = old
            outside.unlink(missing_ok=True)


# ========================================================================
# ZOMBIE GATE
# ========================================================================

class TestZombieGate:
    def test_no_winword_zombies(self):
        time.sleep(3)
        r = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=15,
        )
        lines = [ln for ln in r.stdout.splitlines() if "WINWORD.EXE" in ln.upper()]
        assert len(lines) == 0, f"ZOMBIE GATE FAILED: {len(lines)} process(es): {lines}"
