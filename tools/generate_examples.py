"""Regenerate synthetic negative fixtures and read-only validation reports.

Run after installing requirements-dev.txt. Never edits the reviewed baseline.
"""
from pathlib import Path
import hashlib
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import json

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))
from test_equation_audit import synthetic, write_doc, node, NS
from equation_audit import audit_docx


def main():
    baseline = ROOT / "examples/equation_sample.docx"
    before = hashlib.sha256(baseline.read_bytes()).hexdigest()
    reports = ROOT / "reports/examples"
    negatives = ROOT / "examples/negative"
    reports.mkdir(parents=True, exist_ok=True)
    negatives.mkdir(parents=True, exist_ok=True)
    baseline_report = audit_docx(baseline)
    if baseline_report["issues"]:
        raise ValueError("The reviewed sample must have no detected issues.")
    (reports / "valid_sample.json").write_text(json.dumps(baseline_report, indent=2) + "\n", encoding="utf-8")
    expected = {
        "broken_bookmark": "UNMATCHED_BOOKMARK",
        "missing_ref_target": "MISSING_REF_TARGET",
        "malformed_table": "MALFORMED_EQUATION_TABLE",
        "missing_seq": "MISSING_SEQ",
        "visible_border": "VISIBLE_EQUATION_BORDER",
    }
    index = []
    with TemporaryDirectory() as temporary:
        for name, error in expected.items():
            doc = synthetic()
            if name == "broken_bookmark":
                end = doc.find(".//w:bookmarkEnd", NS)
                end.getparent().remove(end)
            elif name == "missing_ref_target":
                doc.find("w:body/w:p", NS).find(".//w:instrText", NS).text = " REF MissingEquation \\h "
            elif name == "malformed_table":
                table = doc.find(".//w:tbl", NS)
                table.find("w:tblGrid", NS).append(node("w:gridCol", w_w=1310))
                cell = node("w:tc"); cell.append(node("w:p"))
                table.find("w:tr", NS).append(cell)
            elif name == "visible_border":
                doc.find(".//w:tblBorders/w:top", NS).set("{" + NS["w"] + "}val", "single")
            else:
                paragraph = doc.find(".//w:tbl/w:tr", NS).findall("w:tc", NS)[2].find("w:p", NS)
                for run in paragraph.findall("w:r", NS):
                    paragraph.remove(run)
                run = node("w:r"); run.append(node("w:t", "1")); paragraph.append(run)
            generated = write_doc(Path(temporary), doc)
            destination = negatives / (name + ".docx")
            # Normalize ZIP metadata for reproducible synthetic fixtures.
            with ZipFile(generated) as source, ZipFile(destination, "w") as output:
                for part in sorted(source.namelist()):
                    info = ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
                    info.create_system = 3
                    info.compress_type = ZIP_DEFLATED
                    info.external_attr = 0o644 << 16
                    output.writestr(info, source.read(part))
            report = audit_docx(destination)
            if error not in {i["code"] for i in report["issues"]} or not report["input_unchanged"]:
                raise ValueError(f"Expected error not detected for {name}.")
            (reports / (name + ".json")).write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
            index.append({"fixture": destination.relative_to(ROOT).as_posix(), "expected_error": error,
                          "expected_exit_code": 1, "sha256": report["sha256"]})
    (reports / "negative_cases.json").write_text(json.dumps(index, indent=2) + "\n", encoding="utf-8")
    if hashlib.sha256(baseline.read_bytes()).hexdigest() != before:
        raise ValueError("Baseline changed during fixture generation.")
    print(f"Generated {len(expected)} synthetic negative fixtures and their validation reports; baseline unchanged.")


if __name__ == "__main__":
    main()
