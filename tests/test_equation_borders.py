"""Explicit border declarations and conservative inheritance handling."""
import hashlib
import json
import subprocess
import sys
from zipfile import ZipFile

from lxml import etree as ET
import pytest

from test_equation_audit import MODULE, NS, node, synthetic, write_doc, codes
from equation_audit import audit_docx, q, summary


@pytest.mark.parametrize("side", ["top", "bottom", "left", "right", "insideH", "insideV"])
def test_explicit_visible_table_border_is_an_error(tmp_path, side):
    doc = synthetic()
    doc.find(f".//w:tblPr/w:tblBorders/w:{side}", NS).set(q("w:val"), "single")
    path = write_doc(tmp_path, doc); before = path.read_bytes()
    report = audit_docx(path)
    assert report["status"] == "fail"
    assert "VISIBLE_EQUATION_BORDER" in codes(report)
    assert report["equation_tables"][0]["borders"]["status"] == "error"
    assert path.read_bytes() == before


@pytest.mark.parametrize("value", ["none", "nil"])
def test_explicit_invisible_borders_are_valid(tmp_path, value):
    doc = synthetic()
    for border in doc.findall(".//w:tblBorders/*", NS):
        border.set(q("w:val"), value)
    report = audit_docx(write_doc(tmp_path, doc))
    assert report["issues"] == []
    assert report["equation_tables"][0]["borders"]["status"] == "pass"


@pytest.mark.parametrize("style", ["double", "dashed", "wave"])
def test_other_explicit_line_styles_are_detected(tmp_path, style):
    doc = synthetic(); doc.find(".//w:tblBorders/w:top", NS).set(q("w:val"), style)
    assert "VISIBLE_EQUATION_BORDER" in codes(audit_docx(write_doc(tmp_path, doc)))


def test_direct_cell_border_is_not_hidden_by_table_none(tmp_path):
    doc = synthetic(); borders = node("w:tcBorders"); borders.append(node("w:top", w_val="single"))
    doc.find(".//w:tcPr", NS).append(borders)
    assert "VISIBLE_EQUATION_BORDER" in codes(audit_docx(write_doc(tmp_path, doc)))


def test_row_border_exception_is_detected(tmp_path):
    doc = synthetic(); exception = node("w:tblPrEx"); borders = node("w:tblBorders")
    borders.append(node("w:top", w_val="single")); exception.append(borders)
    doc.find(".//w:tr", NS).insert(0, exception)
    assert "VISIBLE_EQUATION_BORDER" in codes(audit_docx(write_doc(tmp_path, doc)))


@pytest.mark.parametrize("partial", [False, True])
def test_absent_border_declarations_are_unverified(tmp_path, partial):
    doc = synthetic(); borders = doc.find(".//w:tblBorders", NS)
    if partial:
        borders.remove(borders.find("w:bottom", NS))
    else:
        borders.getparent().remove(borders)
    report = audit_docx(write_doc(tmp_path, doc))
    assert report["error_count"] == 0
    assert "BORDER_STYLE_UNVERIFIED" in codes(report)
    assert report["equation_tables"][0]["borders"]["status"] == "unverified"
    assert "PASS WITH WARNINGS" in summary(report)


def styled_report(tmp_path, style_xml, direct=True):
    doc = synthetic(); props = doc.find(".//w:tblPr", NS)
    props.insert(0, node("w:tblStyle", w_val="Grid"))
    if not direct:
        props.remove(props.find("w:tblBorders", NS))
    styles = ('<w:styles xmlns:w="' + NS['w'] + '">' + style_xml + '</w:styles>').encode()
    return audit_docx(write_doc(tmp_path, doc, extras={"word/styles.xml": styles}))


def test_direct_none_overrides_known_inherited_table_grid(tmp_path):
    r = styled_report(tmp_path, '<w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:style>')
    assert r["issues"] == []
    assert r["equation_tables"][0]["borders"]["status"] == "pass"


def test_inherited_table_borders_are_not_claimed_resolved(tmp_path):
    r = styled_report(tmp_path, '<w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:style>', direct=False)
    assert r["error_count"] == 0
    assert r["equation_tables"][0]["borders"]["status"] == "unverified"
    assert "BORDER_STYLE_UNVERIFIED" in codes(r)


@pytest.mark.parametrize("style_xml", [
    '',
    '<w:style w:type="table" w:styleId="Grid"><w:basedOn w:val="Grid"/></w:style>',
    '<w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style>',
])
def test_missing_cyclic_or_conditional_cell_styles_are_unverified(tmp_path, style_xml):
    r = styled_report(tmp_path, style_xml)
    assert r["error_count"] == 0
    assert "BORDER_STYLE_UNVERIFIED" in codes(r)
    assert r["equation_tables"][0]["borders"]["status"] == "unverified"


def test_unknown_border_value_is_unverified_not_fabricated_visible(tmp_path):
    doc = synthetic(); doc.find(".//w:tblBorders/w:top", NS).set(q("w:val"), "customUnknown")
    r = audit_docx(write_doc(tmp_path, doc))
    assert r["error_count"] == 0 and "BORDER_STYLE_UNVERIFIED" in codes(r)


def test_independently_constructed_single_border_cli_regression(tmp_path):
    # Construct a separate OPC package rather than mutating the shared synthetic builder.
    w, m = NS['w'], NS['m']
    doc = f'''<w:document xmlns:w="{w}" xmlns:m="{m}"><w:body><w:tbl>
    <w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>
    <w:top w:val="single"/><w:left w:val="none"/><w:bottom w:val="none"/>
    <w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
    </w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1310"/><w:gridCol w:w="6740"/><w:gridCol w:w="1310"/></w:tblGrid><w:tr>
    <w:tc><w:tcPr><w:tcW w:w="1310" w:type="dxa"/></w:tcPr><w:p/></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="6740" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><m:oMathPara><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="1310" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:fldSimple w:instr="SEQ Eq"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:tc>
    </w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>'''
    path = tmp_path / "independent-border.docx"
    with ZipFile(path, "w") as z:
        z.writestr("word/document.xml", doc)
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    run = subprocess.run([sys.executable, str(MODULE), str(path), "--json", "-"], capture_output=True, text=True)
    assert run.returncode == 1
    report = json.loads(run.stdout)
    assert codes(report) == {"VISIBLE_EQUATION_BORDER"}
    assert report["equation_tables"][0]["borders"]["status"] == "error"
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before
