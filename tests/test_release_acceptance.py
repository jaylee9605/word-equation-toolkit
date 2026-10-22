"""Acceptance regressions discovered by reading code, not fixture-only happy paths."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile

from lxml import etree as ET
import pytest

from test_equation_audit import MODULE, NS, node, synthetic, write_doc, codes
from equation_audit import audit_docx, main, q

PROJECT = Path(__file__).resolve().parents[1]
BASELINE_SHA256 = "e8a9658a9a26184319ba0e0d2a673d4c36549c9fbffbbcfde6ce79c5d207ca84"


def test_reviewed_baseline_is_unchanged_and_has_no_errors():
    path = PROJECT / "examples/equation_sample.docx"
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    assert before == BASELINE_SHA256
    report = audit_docx(path)
    assert report["issues"] == []
    assert report["stored_numbering"]["status"] == "pass"
    assert all(r["cache_status"] == "pass" for r in report["references"])
    assert report["counts"] == {
        "omml_objects": 14, "display_math_objects": 10, "inline_math_objects": 4,
        "numbered_equation_rows": 10, "seq_eq_fields": 10, "ref_fields": 10,
        "bookmark_starts": 10, "bookmark_ends": 10, "matched_bookmark_pairs": 10,
        "equation_tables": 9,
    }
    assert report["matrices"][0]["rows"] == 5
    assert report["matrices"][0]["cells_per_row"] == [2] * 5
    assert report["line_numbering"][0]["attributes"] == {"countBy": "1", "restart": "continuous"}
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before


def test_template_static_audit_does_not_change_template():
    path = PROJECT / "templates/equation-template.dotx"
    before = path.read_bytes()
    assert audit_docx(path)["issues"] == []
    assert path.read_bytes() == before


def test_deterministic_errors_across_hash_seeds(tmp_path):
    doc = synthetic(4)
    for e in doc.findall(".//w:bookmarkEnd", NS):
        e.getparent().remove(e)
    path = write_doc(tmp_path, doc)
    outputs = []
    for seed in ["1", "7", "31"]:
        env = dict(os.environ, PYTHONHASHSEED=seed)
        run = subprocess.run([sys.executable, str(MODULE), str(path), "--json", "-"],
                             capture_output=True, text=True, env=env)
        assert run.returncode == 1
        outputs.append(run.stdout)
    assert outputs[0] == outputs[1] == outputs[2]


def test_required_opc_parts_cannot_false_pass(tmp_path):
    path = tmp_path / "not-a-docx.docx"
    with ZipFile(path, "w") as z:
        z.writestr("word/document.xml", ET.tostring(synthetic()))
    with pytest.raises(ValueError, match="OPC"):
        audit_docx(path)


def test_percent_encoded_internal_relationship_is_valid(tmp_path):
    rels = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="style%20sample.xml"/></Relationships>'
    path = write_doc(tmp_path, extras={"word/_rels/document.xml.rels": rels,
                                     "word/style sample.xml": b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'})
    assert "MISSING_RELATIONSHIP_PART" not in codes(audit_docx(path))


def test_grid_span_one_is_not_a_merged_cell(tmp_path):
    d = synthetic()
    d.find(".//w:tcPr", NS).append(node("w:gridSpan", w_val=1))
    assert audit_docx(write_doc(tmp_path, d))["issues"] == []


def test_top_gutter_does_not_reduce_horizontal_text_area(tmp_path):
    d = synthetic(); d.find(".//w:pgMar", NS).set(q("w:gutter"), "400")
    settings = b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:gutterAtTop/></w:settings>'
    r = audit_docx(write_doc(tmp_path, d, extras={"word/settings.xml": settings}))
    assert "TABLE_TOO_WIDE" not in codes(r)
    assert r["equation_tables"][0]["available_text_width_twips"] == 9360


def test_missing_native_math_is_detected(tmp_path):
    d = synthetic(); math = d.find(".//m:oMathPara", NS); math.getparent().remove(math)
    assert "MISSING_OMML" in codes(audit_docx(write_doc(tmp_path, d)))


def test_conflicting_sequence_controls_remain_unverified(tmp_path):
    d = synthetic(); d.find(".//w:instrText", NS).text = " SEQ Eq \\r 7 \\c "
    r = audit_docx(write_doc(tmp_path, d))
    assert r["stored_numbering"]["status"] == "unverified"
    assert "NONCONTINUOUS_SEQUENCE" not in codes(r)


def test_custom_default_paragraph_style_is_resolved(tmp_path):
    d = synthetic(); p = d.find(".//w:tbl/w:tr", NS).findall("w:tc", NS)[1].find("w:p", NS)
    p.find("w:pPr", NS).remove(p.find("w:pPr/w:jc", NS))
    styles = b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Body"><w:name w:val="Body"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style></w:styles>'
    assert "EQUATION_ALIGNMENT" not in codes(audit_docx(write_doc(tmp_path, d, extras={"word/styles.xml": styles})))


def test_report_replace_does_not_follow_alias_created_after_guard(tmp_path, monkeypatch):
    path = write_doc(tmp_path); before = path.read_bytes(); output = tmp_path / "report.json"
    real_replace = os.replace

    def create_alias_then_replace(source, destination):
        os.link(path, output)
        real_replace(source, destination)

    monkeypatch.setattr(os, "replace", create_alias_then_replace)
    assert main([str(path), "--json", str(output)]) == 0
    assert json.loads(output.read_text(encoding="utf-8"))["input_unchanged"]
    assert path.read_bytes() == before


def test_atomic_report_failure_leaves_input_unchanged(tmp_path, monkeypatch):
    path = write_doc(tmp_path); before = path.read_bytes()

    def refuse_replace(*args):
        raise PermissionError("synthetic output permission failure")

    monkeypatch.setattr(os, "replace", refuse_replace)
    assert main([str(path), "--json", str(tmp_path / "report.json")]) == 2
    assert path.read_bytes() == before
    assert not list(tmp_path.glob(".equation-audit-*"))


def test_cli_invalid_input_json_mode_stays_clean(tmp_path):
    path = tmp_path / "invalid.docx"; path.write_bytes(b"invalid")
    run = subprocess.run([sys.executable, str(MODULE), str(path), "--json", "-"], capture_output=True, text=True)
    assert run.returncode == 2 and not run.stdout and "Traceback" not in run.stderr


@pytest.mark.parametrize("suffix", [".docx", ".dotx", ".docm", ".dotm"])
def test_report_cannot_overwrite_a_different_word_artifact(tmp_path, suffix):
    path = write_doc(tmp_path)
    other = tmp_path / ("other" + suffix); other.write_bytes(b"preserved artifact")
    assert main([str(path), "--json", str(other)]) == 2
    assert other.read_bytes() == b"preserved artifact"


def test_macro_enabled_container_is_explicitly_unsupported(tmp_path):
    path = write_doc(tmp_path)
    with ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        b"application/vnd.ms-word.document.macroEnabled.main+xml")
    with ZipFile(path, "w") as z:
        for name, data in parts.items():
            z.writestr(name, data)
    with pytest.raises(ValueError, match="macro-free"):
        audit_docx(path)


def test_empty_matrix_rows_are_not_false_pass(tmp_path):
    d = synthetic(); matrix = node("m:m"); matrix.append(node("m:mr"))
    d.find(".//m:oMath", NS).append(matrix)
    assert "MALFORMED_MATRIX" in codes(audit_docx(write_doc(tmp_path, d)))
