"""Synthetic packages only: no manuscript or external file is a test fixture."""

from copy import deepcopy
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET
import pytest

MODULE = Path(__file__).resolve().parents[1] / "src" / "equation_audit.py"
sys.path.insert(0, str(MODULE.parent))
from equation_audit import audit_docx, q, NS


def node(name, text=None, **attrs):
    e = ET.Element(q(name))
    e.text = text
    for key, value in attrs.items():
        e.set(q(key.replace("_", ":")), str(value))
    return e


def field_runs(p, code, cached, simple=False):
    if simple:
        f = node("w:fldSimple", w_instr=code)
        r = node("w:r"); r.append(node("w:t", cached)); f.append(r); p.append(f)
        return
    for kind, value in [("begin", None), ("code", code), ("separate", None), ("text", cached), ("end", None)]:
        r = node("w:r")
        r.append(node("w:instrText", value) if kind == "code" else node("w:t", value) if kind == "text"
                 else node("w:fldChar", w_fldCharType=kind))
        p.append(r)


def synthetic(equations=1, simple=False):
    d = ET.Element(q("w:document"), nsmap=NS)
    body = ET.SubElement(d, q("w:body"))
    for number in range(1, equations + 1):
        table = node("w:tbl"); body.append(table)
        props = node("w:tblPr"); props.append(node("w:tblW", w_w=9360, w_type="dxa")); table.append(props)
        borders = node("w:tblBorders")
        for side in ["top", "bottom", "left", "right", "insideH", "insideV"]:
            borders.append(node("w:" + side, w_val="nil"))
        props.append(borders)
        grid = node("w:tblGrid"); table.append(grid)
        row = node("w:tr"); table.append(row)
        for width in [1310, 6740, 1310]:
            grid.append(node("w:gridCol", w_w=width))
            cell = node("w:tc"); row.append(cell)
            cp = node("w:tcPr"); cp.append(node("w:tcW", w_w=width, w_type="dxa")); cell.append(cp)
            cell.append(node("w:p"))
        cells = row.findall("w:tc", NS)
        p = cells[1].find("w:p", NS)
        pp = node("w:pPr"); pp.append(node("w:jc", w_val="center")); p.append(pp)
        mp = node("m:oMathPara"); math = node("m:oMath"); mp.append(math); p.append(mp)
        r = node("m:r"); r.append(node("m:t", "y = x + 1")); math.append(r)
        p = cells[2].find("w:p", NS)
        pp = node("w:pPr"); pp.append(node("w:jc", w_val="right")); p.append(pp)
        p.append(node("w:bookmarkStart", w_id=number, w_name=f"Eq{number}"))
        field_runs(p, " SEQ Eq ", str(number), simple)
        p.append(node("w:bookmarkEnd", w_id=number))
    for number in range(1, equations + 1):
        p = node("w:p"); body.append(p)
        field_runs(p, f" REF Eq{number} \\h ", str(number), simple)
    section = node("w:sectPr"); body.append(section)
    section.append(node("w:pgSz", w_w=12240, w_h=15840))
    section.append(node("w:pgMar", w_left=1440, w_right=1440, w_top=1440, w_bottom=1440))
    section.append(node("w:lnNumType", w_countBy=1, w_restart="continuous"))
    return d


def write_doc(tmp_path, doc=None, name="synthetic.docx", extras=None):
    p = tmp_path / name
    with ZipFile(p, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", ET.tostring(doc if doc is not None else synthetic()))
        for part, data in (extras or {}).items():
            z.writestr(part, data)
    return p


def codes(report):
    return {i["code"] for i in report["issues"]}


@pytest.mark.parametrize("simple", [False, True])
def test_valid_synthetic_sample_and_no_modification(tmp_path, simple):
    path = write_doc(tmp_path, synthetic(3, simple))
    before = hashlib.sha256(path.read_bytes()).digest()
    r = audit_docx(path)
    assert r["status"] == "pass", r["issues"]
    assert r["counts"]["seq_eq_fields"] == 3
    assert r["counts"]["ref_fields"] == 3
    assert r["counts"]["matched_bookmark_pairs"] == 3
    assert r["stored_numbering"]["status"] == "pass"
    assert r["input_unchanged"]
    assert before == hashlib.sha256(path.read_bytes()).digest()


def test_missing_referenced_bookmark(tmp_path):
    d = synthetic()
    for tag in ["bookmarkStart", "bookmarkEnd"]:
        e = d.find(".//w:" + tag, NS); e.getparent().remove(e)
    assert "MISSING_REF_TARGET" in codes(audit_docx(write_doc(tmp_path, d)))


def test_unmatched_bookmark(tmp_path):
    d = synthetic(); e = d.find(".//w:bookmarkEnd", NS); e.getparent().remove(e)
    assert "UNMATCHED_BOOKMARK" in codes(audit_docx(write_doc(tmp_path, d)))


def test_duplicate_bookmark_names(tmp_path):
    d = synthetic(2); d.findall(".//w:bookmarkStart", NS)[1].set(q("w:name"), "Eq1")
    assert "DUPLICATE_BOOKMARK_NAME" in codes(audit_docx(write_doc(tmp_path, d)))


def test_bookmark_end_precedes_start(tmp_path):
    d = synthetic(); start = d.find(".//w:bookmarkStart", NS); end = d.find(".//w:bookmarkEnd", NS)
    p = start.getparent(); p.remove(end); p.insert(0, end)
    assert "UNMATCHED_BOOKMARK" in codes(audit_docx(write_doc(tmp_path, d)))


def test_malformed_equation_table_four_columns(tmp_path):
    d = synthetic(); t = d.find(".//w:tbl", NS)
    t.find("w:tblGrid", NS).append(node("w:gridCol", w_w=1310))
    t.find("w:tr", NS).append(deepcopy(t.find("w:tr/w:tc", NS)))
    assert "MALFORMED_EQUATION_TABLE" in codes(audit_docx(write_doc(tmp_path, d)))


def test_missing_seq_field(tmp_path):
    d = synthetic(); p = d.findall(".//w:tbl/w:tr/w:tc", NS)[2].find("w:p", NS)
    for r in p.findall("w:r", NS):
        p.remove(r)
    r = node("w:r"); r.append(node("w:t", "1")); p.append(r)
    assert "MISSING_SEQ" in codes(audit_docx(write_doc(tmp_path, d)))


def test_split_complex_instructions(tmp_path):
    d = synthetic()
    for e in list(d.findall(".//w:instrText", NS)):
        r = e.getparent(); p = r.getparent(); at = list(p).index(r)
        text = e.text; p.remove(r)
        for offset, chunk in enumerate([text[:2], text[2:5], text[5:]]):
            nr = node("w:r"); nr.append(node("w:instrText", chunk)); p.insert(at + offset, nr)
    assert audit_docx(write_doc(tmp_path, d))["status"] == "pass"


def test_nested_field_cached_results_are_not_instructions(tmp_path):
    d = synthetic(); p = node("w:p"); body = d.find("w:body", NS); body.insert(0, p)
    for kind in ["begin", "separate"]:
        r = node("w:r"); r.append(node("w:fldChar", w_fldCharType=kind)); p.append(r)
        if kind == "begin":
            r = node("w:r"); r.append(node("w:instrText", " IF 1 = 1 ")); p.append(r)
    field_runs(p, " REF Eq1 \\h ", "1")
    r = node("w:r"); r.append(node("w:fldChar", w_fldCharType="end")); p.append(r)
    r = audit_docx(write_doc(tmp_path, d))
    assert r["status"] == "pass"
    assert r["counts"]["ref_fields"] == 2


def test_unclosed_field(tmp_path):
    d = synthetic(); e = next(e for e in d.findall(".//w:fldChar", NS) if e.get(q("w:fldCharType")) == "end")
    e.getparent().remove(e)
    assert "MALFORMED_FIELD" in codes(audit_docx(write_doc(tmp_path, d)))


def test_noncontinuous_cached_sequence(tmp_path):
    d = synthetic(2)
    d.findall(".//w:tbl", NS)[1].find(".//w:t", NS).text = "9"
    assert "NONCONTINUOUS_SEQUENCE" in codes(audit_docx(write_doc(tmp_path, d)))


def test_supported_sequence_restart(tmp_path):
    d = synthetic(2)
    tables = d.findall(".//w:tbl", NS)
    tables[0].find(".//w:instrText", NS).text = " SEQ Eq \\r 7 "
    for t, value in zip(tables, ["7", "8"]):
        t.find(".//w:t", NS).text = value
    # Update the synthetic REF caches too.
    for p, value in zip(d.find("w:body", NS).findall("w:p", NS), ["7", "8"]):
        p.find(".//w:t", NS).text = value
    assert audit_docx(write_doc(tmp_path, d))["status"] == "pass"


def test_unknown_sequence_switch_is_unverified(tmp_path):
    d = synthetic(); d.find(".//w:instrText", NS).text = " SEQ Eq \\s 1 "
    r = audit_docx(write_doc(tmp_path, d))
    assert r["stored_numbering"]["status"] == "unverified"
    assert r["stored_numbering"]["fields"][0]["expected"] is None


def test_empty_sequence_cache_is_unverified(tmp_path):
    d = synthetic(); d.find(".//w:t", NS).text = ""
    assert audit_docx(write_doc(tmp_path, d))["stored_numbering"]["status"] == "unverified"


def test_stale_ref_cache(tmp_path):
    d = synthetic(); d.find("w:body/w:p", NS).find(".//w:t", NS).text = "99"
    assert "STALE_REF_CACHE" in codes(audit_docx(write_doc(tmp_path, d)))


def test_relative_ref_is_not_compared_as_number(tmp_path):
    d = synthetic(); p = d.find("w:body/w:p", NS)
    p.find(".//w:instrText", NS).text = " REF Eq1 \\p "
    p.find(".//w:t", NS).text = "above"
    r = audit_docx(write_doc(tmp_path, d))
    assert r["status"] == "pass"
    assert r["references"][0]["cache_status"] == "unverified"


def test_continuous_line_configuration_is_information(tmp_path):
    r = audit_docx(write_doc(tmp_path))
    assert r["line_numbering"][0]["attributes"] == {"countBy": "1", "restart": "continuous"}
    assert r["line_numbering"][0]["visible_range"] == "unverified"
    assert "Visible line-number range" in r["unverified_checks"]
    assert r["status"] == "pass"


@pytest.mark.parametrize("problem,expected", [
    ("width", "TABLE_TOO_WIDE"), ("asymmetry", "ASYMMETRIC_SPACERS"),
    ("span", "MALFORMED_EQUATION_TABLE"), ("alignment", "NUMBER_ALIGNMENT"),
])
def test_table_geometry_errors(tmp_path, problem, expected):
    d = synthetic(); t = d.find(".//w:tbl", NS)
    if problem == "width":
        t.find("w:tblPr/w:tblW", NS).set(q("w:w"), "10669")
    elif problem == "asymmetry":
        t.find("w:tblGrid/w:gridCol", NS).set(q("w:w"), "1500")
    elif problem == "span":
        t.find("w:tr/w:tc/w:tcPr", NS).append(node("w:gridSpan", w_val=2))
    else:
        t.findall("w:tr/w:tc", NS)[2].find("w:p/w:pPr/w:jc", NS).set(q("w:val"), "left")
    assert expected in codes(audit_docx(write_doc(tmp_path, d)))


def test_malformed_matrix(tmp_path):
    d = synthetic(); math = d.find(".//m:oMath", NS); matrix = node("m:m"); math.append(matrix)
    for width in [2, 1]:
        row = node("m:mr"); matrix.append(row)
        for _ in range(width):
            row.append(node("m:e"))
    assert "MALFORMED_MATRIX" in codes(audit_docx(write_doc(tmp_path, d)))


def test_cli_summary_json_and_input_hash(tmp_path):
    p = write_doc(tmp_path); before = p.read_bytes(); report = tmp_path / "report.json"
    run = subprocess.run([sys.executable, str(MODULE), str(p), "--json", str(report)], capture_output=True, text=True)
    assert run.returncode == 0, run.stderr
    assert "static OOXML only" in run.stdout
    assert json.loads(report.read_text(encoding="utf-8"))["input_unchanged"]
    assert p.read_bytes() == before


def test_cli_json_stdout_is_parseable(tmp_path):
    p = write_doc(tmp_path)
    run = subprocess.run([sys.executable, str(MODULE), str(p), "--json", "-"], capture_output=True, text=True)
    assert run.returncode == 0
    assert json.loads(run.stdout)["status"] == "pass"
    assert "static OOXML only" in run.stderr


@pytest.mark.parametrize("alias", [False, True])
def test_cli_refuses_json_input_overwrite(tmp_path, alias):
    p = write_doc(tmp_path); before = p.read_bytes(); destination = p
    if alias:
        destination = tmp_path / "alias.json"
        try:
            os.link(p, destination)
        except OSError:
            pytest.skip("Hard links are unavailable on this filesystem")
    run = subprocess.run([sys.executable, str(MODULE), str(p), "--json", str(destination)], capture_output=True, text=True)
    assert run.returncode == 2
    assert p.read_bytes() == before


def test_cli_error_exit_code(tmp_path):
    p = tmp_path / "invalid.docx"; p.write_bytes(b"not a ZIP")
    run = subprocess.run([sys.executable, str(MODULE), str(p)], capture_output=True, text=True)
    assert run.returncode == 2
    assert "could not run" in run.stderr


def test_cli_detected_error_exit_code(tmp_path):
    d = synthetic(); e = d.find(".//w:bookmarkEnd", NS); e.getparent().remove(e)
    p = write_doc(tmp_path, d)
    run = subprocess.run([sys.executable, str(MODULE), str(p)], capture_output=True, text=True)
    assert run.returncode == 1
    assert "UNMATCHED_BOOKMARK" in run.stdout


def test_reject_doctype(tmp_path):
    p = tmp_path / "dtd.docx"
    with ZipFile(p, "w") as z:
        z.writestr("word/document.xml", b'<!DOCTYPE x [<!ENTITY x "bad">]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    with pytest.raises(ValueError, match="DOCTYPE"):
        audit_docx(p)


def test_split_visible_reference_error(tmp_path):
    d = synthetic(); p = node("w:p"); d.find("w:body", NS).insert(0, p)
    for text in ["Error! Reference ", "source not found."]:
        r = node("w:r"); r.append(node("w:t", text)); p.append(r)
    assert "VISIBLE_REFERENCE_ERROR" in codes(audit_docx(write_doc(tmp_path, d)))


def test_unsupported_document_namespace_is_not_false_pass(tmp_path):
    d = ET.Element('{http://purl.oclc.org/ooxml/wordprocessingml/main}document')
    with pytest.raises(ValueError, match="transitional"):
        audit_docx(write_doc(tmp_path, d))
