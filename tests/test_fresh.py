import datetime as dt
import io
import zipfile

import docx
import pytest

from docx_fresh import TEMPLATE_TIME, fresh
from docx_fresh.__main__ import main

NOW = dt.datetime(2026, 11, 24, 9, 30, tzinfo=dt.timezone.utc)


def saved(document):
    buf = io.BytesIO()
    document.save(buf)
    buf.seek(0)
    return buf


def members(buf):
    with zipfile.ZipFile(buf) as z:
        return {name: z.read(name) for name in z.namelist()}


def test_default_template_is_stale_without_fresh():
    files = members(saved(docx.Document()))
    assert b"python-docx" in files["docProps/core.xml"]
    assert b"2013-12-23" in files["docProps/core.xml"]
    assert b"Macintosh" in files["docProps/app.xml"]
    assert "docProps/thumbnail.jpeg" in files


def test_fresh_replaces_template_metadata():
    document = fresh(docx.Document(), now=NOW)
    files = members(saved(document))
    core, app = files["docProps/core.xml"], files["docProps/app.xml"]

    assert b"python-docx" not in core
    assert b"2013" not in core
    assert b"2026-11-24T09:30:00Z" in core
    for stale in (b"Macintosh", b"<Pages>", b"<Words>", b"<AppVersion>", b"<TotalTime>"):
        assert stale not in app
    assert b"<Template>Normal.dotm</Template>" in app
    assert not any(name.startswith("docProps/thumbnail") for name in files)
    assert b"thumbnail" not in files["_rels/.rels"]


def test_saved_file_reopens_with_fresh_values():
    reopened = docx.Document(saved(fresh(docx.Document(), author="Report bot", now=NOW)))
    props = reopened.core_properties
    assert props.author == "Report bot"
    assert props.last_modified_by == "Report bot"
    assert props.comments == ""
    assert props.created == NOW
    assert props.modified == NOW


def test_current_time_is_used_by_default():
    before = dt.datetime.now(dt.timezone.utc).replace(microsecond=0)
    props = fresh(docx.Document()).core_properties
    assert props.created >= before
    assert props.modified == props.created


def test_own_template_is_left_alone():
    template = docx.Document()
    props = template.core_properties
    props.author = "Alice"
    props.last_modified_by = "Bob"
    props.comments = "Quarterly report"
    props.created = dt.datetime(2020, 1, 2, tzinfo=dt.timezone.utc)
    props.modified = dt.datetime(2021, 3, 4, tzinfo=dt.timezone.utc)
    template_bytes = saved(template)
    app_before = members(template_bytes)["docProps/app.xml"]
    template_bytes.seek(0)

    document = fresh(docx.Document(template_bytes), now=NOW)
    files = members(saved(document))
    props = document.core_properties
    assert (props.author, props.last_modified_by, props.comments) == (
        "Alice",
        "Bob",
        "Quarterly report",
    )
    assert props.created == dt.datetime(2020, 1, 2, tzinfo=dt.timezone.utc)
    assert props.modified == dt.datetime(2021, 3, 4, tzinfo=dt.timezone.utc)
    assert files["docProps/app.xml"] == app_before
    assert "docProps/thumbnail.jpeg" in files


def test_fresh_twice_is_harmless():
    document = fresh(fresh(docx.Document(), now=NOW), now=dt.datetime(2030, 1, 1, tzinfo=dt.timezone.utc))
    assert document.core_properties.created == NOW


def test_naive_now_is_rejected():
    with pytest.raises(ValueError):
        fresh(docx.Document(), now=dt.datetime(2026, 11, 24))


def test_body_is_unchanged():
    document = docx.Document()
    document.add_heading("Title", 1)
    document.add_paragraph("Hello")
    files_before = members(saved(document))
    files_after = members(saved(fresh(document, now=NOW)))
    assert files_after["word/document.xml"] == files_before["word/document.xml"]


def test_cli_fixes_a_file_in_place(tmp_path):
    path = tmp_path / "report.docx"
    docx.Document().save(path)
    assert main([str(path), "--author", "CLI"]) == 0
    props = docx.Document(path).core_properties
    assert props.author == "CLI"
    assert props.created != TEMPLATE_TIME


def test_cli_writes_to_output(tmp_path):
    src, out = tmp_path / "in.docx", tmp_path / "out.docx"
    docx.Document().save(src)
    main([str(src), "-o", str(out)])
    assert docx.Document(src).core_properties.created == TEMPLATE_TIME
    assert docx.Document(out).core_properties.created != TEMPLATE_TIME
