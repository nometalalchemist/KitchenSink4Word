"""Regressions for the v1.4 adversarial-round findings."""

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage


def _template(tmp_path):
    tpl = tmp_path / "tpl.docx"
    srv.create_document(str(tpl))
    srv.insert_paragraphs(
        str(tpl), [{"text": "Hello {{name}}."}], backup=False
    )
    return tpl


def test_f1_long_filename_pattern_capped_not_midrun_crash(tmp_path):
    tpl = _template(tmp_path)
    out = tmp_path / "out"
    rows = [{"name": "ok"}, {"name": "X" * 300}, {"name": "also ok"}]
    r = srv.mail_merge(str(tpl), rows, str(out), filename_pattern="{name}.docx")
    assert r["rows"] == 3
    produced = list(out.glob("*.docx"))
    assert len(produced) == 3
    assert all(len(p.name) <= 185 for p in produced)


def test_f1_midrun_oserror_reports_partial_outputs(tmp_path, monkeypatch):
    tpl = _template(tmp_path)
    out = tmp_path / "out2"
    from word_mcp.ops import mailmerge as mm

    original_save = DocxPackage.save
    calls = {"n": 0}

    def failing_save(self, *a, **k):
        calls["n"] += 1
        if calls["n"] == 2:
            raise OSError(28, "No space left on device")
        return original_save(self, *a, **k)

    monkeypatch.setattr(DocxPackage, "save", failing_save)
    with pytest.raises(WordMcpError, match="already been written|already written"):
        srv.mail_merge(
            str(tpl), [{"name": "a"}, {"name": "b"}, {"name": "c"}], str(out)
        )
    assert len(list(out.glob("*.docx"))) == 1  # honest partial state


def test_f3_protection_creates_missing_settings_part(tmp_path):
    doc = tmp_path / "bare.docx"
    srv.create_document(str(doc))
    srv.insert_paragraphs(
        str(doc), [{"text": "needs protection"}], backup=False
    )
    # strip settings.xml the way bare third-party OOXML producers omit it
    import shutil
    import zipfile

    stripped = tmp_path / "stripped.docx"
    with zipfile.ZipFile(doc) as zin, zipfile.ZipFile(
        stripped, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename == "word/settings.xml":
                continue
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b'<Override PartName="/word/settings.xml" '
                    b'ContentType="application/vnd.openxmlformats-'
                    b'officedocument.wordprocessingml.settings+xml"/>',
                    b"",
                )
            zout.writestr(item, data)
    r = srv.set_document_protection(
        str(stripped), protection="readOnly", password="pw", backup=False
    )
    assert r["protection"] == "readOnly"
    state = srv.get_protection(str(stripped))
    assert state["protected"] is True
    # remove on a doc with no settings.xml is a clean no-op, not an error
    shutil.copy(doc, tmp_path / "n.docx")
    with zipfile.ZipFile(doc) as zin, zipfile.ZipFile(
        tmp_path / "nosettings.docx", "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename != "word/settings.xml":
                zout.writestr(item, zin.read(item.filename))
    r2 = srv.set_document_protection(protection="none", file_path=str(tmp_path / "nosettings.docx"),
                                        backup=False)
    assert r2["removed"] is False
