"""Read-only OOXML equation audit. Python 3.10-3.13; no Word automation or repair."""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass, field
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import posixpath
import re
import sys
import tempfile
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZipFile

from lxml import etree as ET

VERSION = "0.1.0"
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def q(name: str) -> str:
    prefix, local = name.split(":")
    return f"{{{NS[prefix]}}}{local}"


def attr(element, name: str, default=None):
    return element.get(q(name), default) if element is not None else default


def integer(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def tokens(code: str) -> list[str]:
    # Word backslashes are switches, not shell escape characters.
    return [t.strip('"') for t in re.findall(r'"[^"]*"|[^\s]+', code)]


@dataclass
class Field:
    part: str
    anchor: object
    position: int
    kind: str
    code_fragments: list[str] = field(default_factory=list)
    result_fragments: list[str] = field(default_factory=list)
    in_code: bool = True

    @property
    def code(self):
        return "".join(self.code_fragments).strip()

    @property
    def result(self):
        return "".join(self.result_fragments).strip()


def parse_fields(root, part, issue) -> list[Field]:
    """Parse fldSimple and nested complex fields, including split instructions."""
    fields, stack = [], []
    serial = 0

    def visit(node):
        nonlocal serial
        serial += 1
        if node.tag == q("w:fldSimple"):
            current = Field(part, node, serial, "simple", [attr(node, "w:instr", "")], [], False)
            stack.append(current)
            for child in node:
                visit(child)
            if stack and stack[-1] is current:
                stack.pop()
                fields.append(current)
            else:
                issue("MALFORMED_FIELD", "Unbalanced field inside fldSimple.", part)
            return
        if node.tag == q("w:fldChar"):
            kind = attr(node, "w:fldCharType")
            if kind == "begin":
                stack.append(Field(part, node, serial, "complex"))
            elif kind == "separate":
                if not stack or not stack[-1].in_code:
                    issue("MALFORMED_FIELD", "Unmatched or repeated field separator.", part)
                else:
                    stack[-1].in_code = False
            elif kind == "end":
                if not stack or stack[-1].kind == "simple":
                    issue("MALFORMED_FIELD", "Unmatched complex field end.", part)
                else:
                    fields.append(stack.pop())
            else:
                issue("MALFORMED_FIELD", "Unknown fldCharType.", part)
        elif node.tag == q("w:instrText"):
            if stack and stack[-1].in_code:
                stack[-1].code_fragments.append(node.text or "")
            else:
                issue("MALFORMED_FIELD", "Instruction text outside a field instruction.", part)
        elif node.tag == q("w:t"):
            for current in stack:
                if not current.in_code:
                    current.result_fragments.append(node.text or "")
        for child in node:
            visit(child)

    visit(root)
    for current in stack:
        issue("MALFORMED_FIELD", "Field start has no matching end.", part)
    for current in fields:
        if not current.code:
            issue("MALFORMED_FIELD", "Field instruction is empty.", part)
    return sorted(fields, key=lambda f: f.position)


def nearest(element, tag):
    return next((e for e in element.iterancestors(q(tag))), None)


def field_type(f):
    ts = tokens(f.code)
    return ts[0].upper() if ts else ""


def is_seq_eq(f):
    ts = tokens(f.code)
    return len(ts) >= 2 and ts[0].upper() == "SEQ" and ts[1].casefold() == "eq"


def effective_alignment(paragraph, styles, default_alignment="left", default_style="Normal"):
    direct = paragraph.find("w:pPr/w:jc", NS)
    if direct is not None:
        return attr(direct, "w:val")
    sid = attr(paragraph.find("w:pPr/w:pStyle", NS), "w:val", default_style)
    visited = set()
    while sid and sid not in visited and sid in styles:
        visited.add(sid)
        style = styles[sid]
        jc = style.find("w:pPr/w:jc", NS)
        if jc is not None:
            return attr(jc, "w:val")
        sid = attr(style.find("w:basedOn", NS), "w:val")
    if sid and (sid in visited or (sid not in styles and styles)):
        return None  # Unknown/cyclic styles cannot establish alignment.
    return default_alignment


def equation_borders(table, styles, issue, location):
    """Check explicit border declarations, not Word's full style/conflict cascade."""
    visible_styles = {
        "single", "thick", "double", "dotted", "dashed", "dotdash", "dotdotdash",
        "triple", "thinthicksmallgap", "thickthinsmallgap", "thinthickthinsmallgap",
        "thinthickmediumgap", "thickthinmediumgap", "thinthickthinmediumgap",
        "thinthicklargegap", "thickthinlargegap", "thinthickthinlargegap",
        "wave", "doublewave", "dashsmallgap", "dashdotstroked", "threedemboss",
        "threedengrave", "outset", "inset",
    }
    invisible = {"nil", "none"}
    reasons, visible = [], []
    direct = table.find("w:tblPr/w:tblBorders", NS)
    declarations = [("table", direct)]
    for ri, row in enumerate(table.findall("w:tr", NS), 1):
        declarations.append((f"row[{ri}]", row.find("w:tblPrEx/w:tblBorders", NS)))
        for ci, cell in enumerate(row.findall("w:tc", NS), 1):
            declarations.append((f"row[{ri}]/cell[{ci}]", cell.find("w:tcPr/w:tcBorders", NS)))
    for scope, borders in declarations:
        if borders is None:
            continue
        for border in borders:
            style = attr(border, "w:val", "").casefold()
            side = ET.QName(border).localname
            if style in visible_styles:
                item = dict(scope=scope, side=side, style=style)
                visible.append(item)
                issue("VISIBLE_EQUATION_BORDER", f"Explicit {scope} {side} border declares '{style}'; equation tables must be borderless.", location)
            elif style not in invisible:
                reasons.append(f"Unknown or missing {scope} {side} border value.")
    # Missing properties can inherit. Do not silently treat absence as 'none'.
    sides = {ET.QName(b).localname: attr(b, "w:val", "").casefold() for b in direct} if direct is not None else {}
    if not all(sides.get(s) in invisible for s in ("top", "left", "bottom", "right", "insideH", "insideV")):
        reasons.append("The six physical table borders are not all explicitly nil/none; inheritance may affect appearance.")
    sid = attr(table.find("w:tblPr/w:tblStyle", NS), "w:val")
    if sid is None:
        sid = next((key for key, s in styles.items() if attr(s, "w:type") == "table"
                    and attr(s, "w:default") in ("1", "true", "on")), None)
    style_id, seen = sid, set()
    while sid:
        if sid in seen or sid not in styles:
            reasons.append("The table style chain is missing or cyclic.")
            break
        seen.add(sid)
        style = styles[sid]
        if style.find(".//w:tcBorders", NS) is not None:
            reasons.append("Inherited cell/conditional borders require Word's style and border-conflict resolution.")
        sid = attr(style.find("w:basedOn", NS), "w:val")
    reasons = list(dict.fromkeys(reasons))
    status = "error" if visible else "unverified" if reasons else "pass"
    if status == "unverified":
        issue("BORDER_STYLE_UNVERIFIED", " ".join(reasons), location, "warning")
    return dict(status=status, explicit_visible=visible, style_id=style_id, reasons=reasons)


def audit_docx(path: str | Path) -> dict:
    """Read a ZIP snapshot; return an audit report. Never writes to the DOCX."""
    path = Path(path)
    raw = path.read_bytes()
    digest = hashlib.sha256(raw).hexdigest()
    issues = []

    def issue(code, message, location="", severity="error"):
        issues.append(dict(code=code, severity=severity, location=location, message=message))

    parser = ET.XMLParser(resolve_entities=False, load_dtd=False, no_network=True)
    roots = {}
    with ZipFile(BytesIO(raw)) as archive:
        infos = archive.infolist()
        if len(infos) > 10000 or sum(i.file_size for i in infos) > 100 * 1024 * 1024:
            raise ValueError("Package exceeds the v0.1 inspection size limit (100 MiB / 10,000 parts).")
        names = [i.filename for i in infos]
        if len(names) != len(set(names)):
            raise ValueError("Duplicate ZIP part names are ambiguous.")
        if any(i.flag_bits & 1 for i in infos):
            raise ValueError("Encrypted ZIP parts are not supported.")
        for name in names:
            if name.endswith((".xml", ".rels")):
                root = ET.fromstring(archive.read(name), parser)
                if root.getroottree().docinfo.doctype:
                    raise ValueError("DOCTYPE declarations are not supported in OOXML parts.")
                roots[name] = root
    if "[Content_Types].xml" not in roots or "_rels/.rels" not in roots:
        raise ValueError("Required OPC content types or root relationships are missing.")
    if "word/document.xml" not in roots:
        raise ValueError("word/document.xml is missing.")
    main = roots["word/document.xml"]
    if main.tag != q("w:document"):
        raise ValueError("Only transitional WordprocessingML document namespaces are supported in v0.1.")
    content_type = next((e.get("ContentType") for e in roots["[Content_Types].xml"]
                         if e.get("PartName") == "/word/document.xml"), None)
    if content_type not in {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    }:
        raise ValueError("Expected a macro-free DOCX or DOTX main content type.")
    if main.find("w:body", NS) is None:
        raise ValueError("The Word document has no body element.")
    story_names = ["word/document.xml"] + sorted(
        n for n in roots if re.fullmatch(r"word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml", n)
    )
    fields = []
    starts, ends, positions = defaultdict(list), defaultdict(list), {}
    bookmark_names = defaultdict(list)
    for part in story_names:
        story = roots[part]
        fields.extend(parse_fields(story, part, issue))
        positions[part] = {e: i for i, e in enumerate(story.iter())}
        for node in story.findall(".//w:bookmarkStart", NS):
            key = (part, attr(node, "w:id"))
            starts[key].append(node)
            name = attr(node, "w:name")
            if not name or key[1] is None:
                issue("INVALID_BOOKMARK", "Bookmark start lacks a name or ID.", part)
            bookmark_names[name].append((part, node))
        for node in story.findall(".//w:bookmarkEnd", NS):
            key = (part, attr(node, "w:id"))
            ends[key].append(node)
    paired = {}
    for key in sorted(starts.keys() | ends.keys(), key=lambda k: (k[0], k[1] or "")):
        if len(starts[key]) != 1 or len(ends[key]) != 1:
            issue("UNMATCHED_BOOKMARK", f"Bookmark ID {key[1]} needs exactly one start and end.", key[0])
        elif positions[key[0]][starts[key][0]] >= positions[key[0]][ends[key][0]]:
            issue("UNMATCHED_BOOKMARK", f"Bookmark ID {key[1]} ends before it starts.", key[0])
        else:
            paired[key] = (starts[key][0], ends[key][0])
    for name, entries in bookmark_names.items():
        if len(entries) > 1:
            issue("DUPLICATE_BOOKMARK_NAME", f"Duplicate bookmark name: {name}.")

    sequences = [f for f in fields if is_seq_eq(f)]
    references = [f for f in fields if field_type(f) == "REF"]
    numbering = []
    counters = defaultdict(int)
    for f in sequences:
        ts = tokens(f.code)
        counter = counters[f.part]
        switches = ts[2:]
        supported, repeat, hidden = True, False, False
        controls = 0
        index = 0
        while index < len(switches):
            switch = switches[index].lower()
            if switch in ("\\r", "\\*"):
                index += 1
                argument = switches[index] if index < len(switches) else ""
                if switch == "\\r" and integer(argument) is not None:
                    counter = integer(argument) - 1
                    controls += 1
                elif switch == "\\*" and argument.upper() == "ARABIC":
                    pass
                else:
                    supported = False
            elif switch == "\\c":
                repeat = True
                controls += 1
            elif switch == "\\h":
                hidden = True
            elif switch == "\\n":
                controls += 1
            else:
                supported = False
            index += 1
        expected = counter if repeat else counter + 1 if counter is not None else None
        if not supported or controls > 1:
            expected = None
        counters[f.part] = expected
        cached = integer(f.result)
        status = "unverified"
        if expected is not None and not hidden and cached is not None:
            status = "pass" if cached == expected else "error"
            if status == "error":
                issue("NONCONTINUOUS_SEQUENCE", f"SEQ Eq cached value {cached}; expected {expected}.", f.part)
        elif expected is not None and not hidden and f.result:
            issue("INVALID_SEQUENCE_CACHE", "SEQ Eq has a nonnumeric cached value.", f.part)
            status = "error"
        numbering.append(dict(part=f.part, kind=f.kind, instruction=f.code, cached=f.result,
                              expected=expected, status=status))

    reference_report = []
    for f in references:
        ts = tokens(f.code)
        target = ts[1] if len(ts) >= 2 and not ts[1].startswith("\\") else None
        status = "unverified"
        entries = bookmark_names.get(target, [])
        if target is None:
            issue("MALFORMED_REF", "REF has no valid target name.", f.part)
            status = "error"
        elif not entries:
            issue("MISSING_REF_TARGET", f"REF target does not exist: {target}.", f.part)
            status = "error"
        elif len(entries) == 1:
            part, start = entries[0]
            pair = paired.get((part, attr(start, "w:id")))
            if pair:
                lo, hi = (positions[part][e] for e in pair)
                visible = "".join((e.text or "") for e in roots[part].iter(q("w:t"))
                                  if lo < positions[part][e] < hi).strip()
                # Formatting/relative-position switches can intentionally change the result.
                plain_switches = [s.lower() for s in ts[2:]]
                if all(s in ("\\h", "\\*", "arabic", "mergeformat", "charformat") for s in plain_switches) and visible.isdecimal():
                    if not f.result:
                        status = "unverified"
                    else:
                        status = "pass" if f.result == visible else "error"
                        if status == "error":
                            issue("STALE_REF_CACHE", f"REF {target} cached value differs from its bookmark.", f.part)
        reference_report.append(dict(part=f.part, kind=f.kind, instruction=f.code,
                                     target=target, cached=f.result, cache_status=status))

    style_root = roots.get("word/styles.xml")
    styles = {attr(s, "w:styleId"): s for s in style_root.findall("w:style", NS)} if style_root is not None else {}
    default_alignment = attr(style_root.find("w:docDefaults/w:pPrDefault/w:pPr/w:jc", NS), "w:val", "left") if style_root is not None else "left"
    default_style = next((sid for sid, s in styles.items()
                          if attr(s, "w:type") == "paragraph" and attr(s, "w:default") in ("1", "true", "on")), "Normal")
    main_positions = positions["word/document.xml"]
    section_nodes = main.findall(".//w:sectPr", NS)
    line_numbering = []
    for index, section in enumerate(section_nodes, 1):
        line = section.find("w:lnNumType", NS)
        line_numbering.append(dict(section=index, enabled=line is not None,
                                   attributes={ET.QName(k).localname: v for k, v in line.attrib.items()} if line is not None else {},
                                   visible_range="unverified"))

    def available_width(table):
        if table not in main_positions:
            return None
        section = next((s for s in section_nodes if main_positions[s] > main_positions[table]), None)
        if section is None:
            return None
        cols = section.find("w:cols", NS)
        if integer(attr(cols, "w:num", "1")) != 1:
            return None
        size, margins = section.find("w:pgSz", NS), section.find("w:pgMar", NS)
        settings = roots.get("word/settings.xml")
        gutter_top = settings.find("w:gutterAtTop", NS) if settings is not None else None
        gutter = 0 if gutter_top is not None and attr(gutter_top, "w:val", "true") not in ("0", "false", "off") else integer(attr(margins, "w:gutter", "0"))
        values = [integer(attr(size, "w:w")), integer(attr(margins, "w:left")),
                  integer(attr(margins, "w:right")), gutter]
        return values[0] - sum(values[1:]) if None not in values else None

    table_report = []
    numbered_rows = 0
    for part in story_names:
        for table_index, table in enumerate(roots[part].findall(".//w:tbl", NS), 1):
            local_fields = [f for f in sequences if f.part == part and nearest(f.anchor, "w:tbl") is table]
            if table.find(".//m:oMathPara", NS) is None and not local_fields:
                continue
            location = f"{part}:table[{table_index}]"
            borders = equation_borders(table, styles, issue, location)
            grid = [integer(attr(c, "w:w")) for c in table.findall("w:tblGrid/w:gridCol", NS)]
            width_node = table.find("w:tblPr/w:tblW", NS)
            width = integer(attr(width_node, "w:w")) if attr(width_node, "w:type") == "dxa" else None
            available = available_width(table)
            if len(grid) != 3:
                issue("MALFORMED_EQUATION_TABLE", f"Expected three grid columns; found {len(grid)}.", location)
            if width is not None and available is not None and width > available:
                issue("TABLE_TOO_WIDE", f"Width {width} exceeds text width {available} twips.", location)
            if None in grid or not grid or any(v <= 0 for v in grid if v is not None):
                issue("MALFORMED_EQUATION_TABLE", "Missing or invalid grid widths.", location)
            elif width is not None and sum(grid) != width:
                issue("TABLE_WIDTH_MISMATCH", "Grid sum differs from table width.", location)
            if len(grid) == 3 and grid[0] != grid[2]:
                issue("ASYMMETRIC_SPACERS", "Spacer and number column widths differ; equation center is displaced.", location)
            if table.find("w:tblPr/w:tblpPr", NS) is not None:
                issue("FLOATING_TABLE_UNVERIFIED", "Floating-table positioning requires visual review.", location, "warning")
            rows = []
            for row_index, row in enumerate(table.findall("w:tr", NS), 1):
                cells = row.findall("w:tc", NS)
                row_location = f"{location}:row[{row_index}]"
                row_fields = [f for f in local_fields if nearest(f.anchor, "w:tr") is row]
                math = row.findall(".//m:oMath", NS)
                if row_fields and not math:
                    issue("MISSING_OMML", "Equation number has no native equation in its row.", row_location)
                if math:
                    numbered_rows += 1
                    if len(row_fields) != 1:
                        issue("MISSING_SEQ" if not row_fields else "MULTIPLE_SEQ", "A numbered equation row must contain exactly one SEQ Eq field.", row_location)
                if len(cells) != 3 or any(attr(c.find("w:tcPr/w:gridSpan", NS), "w:val", "1") != "1" or c.find("w:tcPr/w:vMerge", NS) is not None for c in cells):
                    issue("MALFORMED_EQUATION_TABLE", "Expected three unmerged physical cells.", row_location)
                elif len(grid) == 3:
                    cell_widths = [integer(attr(c.find("w:tcPr/w:tcW", NS), "w:w")) for c in cells]
                    if cell_widths != grid:
                        issue("TABLE_WIDTH_MISMATCH", "Cell widths do not match the grid.", row_location)
                    if cells[0].find(".//m:oMath", NS) is not None or "".join(cells[0].itertext()).strip():
                        issue("NONEMPTY_SPACER", "Left spacer contains content.", row_location)
                    for equation in math:
                        if nearest(equation, "w:tc") is not cells[1]:
                            issue("EQUATION_CELL", "Native equation is outside the center cell.", row_location)
                    for f in row_fields:
                        if nearest(f.anchor, "w:tc") is not cells[2]:
                            issue("NUMBER_CELL", "SEQ Eq is outside the right cell.", row_location)
                    for p in cells[1].findall("w:p", NS):
                        alignment = effective_alignment(p, styles, default_alignment, default_style)
                        if p.find(".//m:oMath", NS) is not None and alignment is not None and alignment != "center":
                            issue("EQUATION_ALIGNMENT", "Equation paragraph is not centered.", row_location)
                    for f in row_fields:
                        p = nearest(f.anchor, "w:p")
                        alignment = effective_alignment(p, styles, default_alignment, default_style) if p is not None else None
                        if alignment is not None and alignment != "right":
                            issue("NUMBER_ALIGNMENT", "Number paragraph is not right-aligned.", row_location)
                rows.append(dict(row=row_index, physical_cells=len(cells), math_objects=len(math), seq_eq=len(row_fields)))
            table_report.append(dict(part=part, table=table_index, grid_twips=grid, width_twips=width,
                                     available_text_width_twips=available, borders=borders, rows=rows))

    math_objects = [m for part in story_names for m in roots[part].findall(".//m:oMath", NS)]
    display = [m for m in math_objects if nearest(m, "m:oMathPara") is not None]
    matrices = []
    for part in story_names:
        for matrix in roots[part].findall(".//m:m", NS):
            row_widths = [len(r.findall("m:e", NS)) for r in matrix.findall("m:mr", NS)]
            counts = matrix.findall("m:mPr/m:mcs/m:mc/m:mcPr/m:count", NS)
            declared = sum(integer(attr(c, "m:val", "1")) or 0 for c in counts) if counts else None
            valid = bool(row_widths) and row_widths[0] > 0 and len(set(row_widths)) == 1 and (declared is None or row_widths[0] == declared)
            if not valid:
                issue("MALFORMED_MATRIX", "Matrix row cells and declared columns disagree.", part)
            matrices.append(dict(part=part, rows=len(row_widths), cells_per_row=row_widths,
                                 declared_columns=declared, structural_status="pass" if valid else "error",
                                 rendering="unverified"))

    metadata = {}
    for part in ["docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"]:
        root = roots.get(part)
        metadata[part] = {ET.QName(e).localname: (e.text or "") for e in root if len(e) == 0} if root is not None else None
    for part, root in roots.items():
        if part.endswith(".rels"):
            for rel in root:
                if rel.get("TargetMode") == "External":
                    issue("EXTERNAL_RELATIONSHIP", "External relationship needs a separate privacy review.", part, "warning")
                else:
                    base = posixpath.dirname(posixpath.dirname(part)) if part != "_rels/.rels" else ""
                    uri = urlsplit(rel.get("Target", ""))
                    target = posixpath.normpath(posixpath.join(base, unquote(uri.path))).lstrip("/")
                    if target not in names:
                        issue("MISSING_RELATIONSHIP_PART", f"Relationship target missing: {target}.", part)
        if part in story_names:
            for p in root.findall(".//w:p", NS):
                visible_text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
                if re.search(r"Error!\s*Reference source not found|错误[!！].*引用源", visible_text, re.I):
                    issue("VISIBLE_REFERENCE_ERROR", "Stored visible reference error found.", part)
    unchanged = digest == hashlib.sha256(path.read_bytes()).hexdigest()
    if not unchanged:
        issue("INPUT_CHANGED", "Input changed during audit (possibly by another process).")
    errors = sum(i["severity"] == "error" for i in issues)
    return {
        "tool": "Word Equation Toolkit", "version": VERSION, "file": path.name,
        "sha256": digest, "input_unchanged": unchanged,
        "status": "fail" if errors else "pass", "error_count": errors,
        "warning_count": sum(i["severity"] == "warning" for i in issues),
        "counts": {"omml_objects": len(math_objects), "display_math_objects": len(display),
                   "inline_math_objects": len(math_objects) - len(display), "numbered_equation_rows": numbered_rows,
                   "seq_eq_fields": len(sequences), "ref_fields": len(references),
                   "bookmark_starts": sum(map(len, starts.values())), "bookmark_ends": sum(map(len, ends.values())),
                   "matched_bookmark_pairs": len(paired), "equation_tables": len(table_report)},
        "stored_numbering": {"status": "error" if any(n["status"] == "error" for n in numbering) else
                             "unverified" if not numbering or any(n["status"] == "unverified" for n in numbering) else "pass",
                             "fields": numbering},
        "references": reference_report, "equation_tables": table_report,
        "line_numbering": line_numbering, "matrices": matrices, "metadata": metadata,
        "unverified_checks": ["Live Word field updates", "Rendered page count", "Visible line-number range",
                              "Visual equation/number alignment", "Renderer/font compatibility", "Complete privacy review"],
        "issues": issues,
    }


def summary(report: dict) -> str:
    c = report["counts"]
    label = report['status'].upper() + (" WITH WARNINGS" if report['status'] == "pass" and report['warning_count'] else "")
    lines = [f"Word Equation Toolkit {VERSION}: {label} (static OOXML only)",
             f"Math: {c['omml_objects']} native; {c['display_math_objects']} display; {c['inline_math_objects']} inline.",
             f"Fields: {c['seq_eq_fields']} SEQ Eq; {c['ref_fields']} REF. Bookmarks: {c['matched_bookmark_pairs']} matched pairs.",
             f"Equation tables: {c['equation_tables']}; numbered rows: {c['numbered_equation_rows']}.",
             f"Stored numbering: {report['stored_numbering']['status']}. Input unchanged: {report['input_unchanged']}."]
    for line in report["line_numbering"]:
        lines.append(f"Line numbering, section {line['section']}: {line['attributes'] if line['enabled'] else 'disabled'}; visible range unverified.")
    lines.extend(f"{i['severity'].upper()} {i['code']} [{i['location']}]: {i['message']}" for i in report["issues"])
    lines.append("Unverified: " + "; ".join(report["unverified_checks"]) + ".")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("docx", type=Path)
    ap.add_argument("--json", dest="json_path", metavar="PATH", help="Optional JSON output; '-' writes JSON to stdout.")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)
    try:
        if args.json_path and args.json_path != "-":
            output = Path(args.json_path)
            if output.suffix.lower() in {".docx", ".dotx", ".docm", ".dotm"}:
                raise ValueError("JSON report output must not be a Word document or template.")
            if output.resolve() == args.docx.resolve() or (output.exists() and os.path.samefile(output, args.docx)):
                raise ValueError("JSON report must not overwrite or alias the input DOCX.")
        report = audit_docx(args.docx)
        if args.json_path:
            data = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
            if args.json_path == "-":
                print(data, end="")
            else:
                # Replace a directory entry rather than writing through a symlink/hard link.
                # Even if a report alias is created after the guard, its target is not edited.
                output = Path(args.json_path)
                temporary = None
                try:
                    with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", newline="\n",
                                                     dir=output.parent, prefix=".equation-audit-", delete=False) as stream:
                        temporary = Path(stream.name)
                        stream.write(data)
                    os.replace(temporary, output)
                finally:
                    if temporary is not None:
                        temporary.unlink(missing_ok=True)
        print(summary(report), file=sys.stderr if args.json_path == "-" else sys.stdout)
        return 1 if report["error_count"] else 0
    except (OSError, ValueError, BadZipFile, RuntimeError, NotImplementedError, ET.XMLSyntaxError) as exc:
        print(f"Audit could not run: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
