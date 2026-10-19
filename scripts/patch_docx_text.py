from __future__ import annotations

import argparse
import hashlib
import json
import re
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from ooxml_utils import parse_xml_preserving_prefixes, serialize_xml, write_docx_exclusive


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W}


def qn(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def text_of(node: ET.Element) -> str:
    return "".join((item.text or "") for item in node.findall(".//w:t", NS))


def direct_body_paragraphs(body: ET.Element) -> list[ET.Element]:
    return [child for child in body if child.tag == qn(W, "p")]


def find_paragraph(body: ET.Element, anchor: str, mode: str) -> ET.Element:
    candidates = direct_body_paragraphs(body)
    if mode == "exact":
        matches = [p for p in candidates if text_of(p).strip() == anchor]
    elif mode == "prefix":
        matches = [p for p in candidates if text_of(p).strip().startswith(anchor)]
    else:
        raise ValueError(f"unsupported match mode: {mode}")
    if len(matches) != 1:
        raise ValueError(f"anchor {anchor!r}: expected 1 paragraph, got {len(matches)}")
    return matches[0]


def following_paragraph(body: ET.Element, paragraph: ET.Element) -> ET.Element:
    paragraphs = direct_body_paragraphs(body)
    index = paragraphs.index(paragraph)
    if index + 1 >= len(paragraphs):
        raise ValueError("anchor has no following direct body paragraph")
    return paragraphs[index + 1]


def plain_text_runs(paragraph: ET.Element, require_uniform: bool = False) -> list[ET.Element]:
    """Reject changes that would discard rich content, anchors, fields or run formatting."""
    if paragraph.find(".//w:pPrChange", NS) is not None:
        raise ValueError("tracked paragraph formatting requires a preserving Range/OOXML edit")
    runs = []
    for child in paragraph:
        if child.tag == qn(W, "pPr"):
            continue
        if child.tag != qn(W, "r"):
            raise ValueError("paragraph contains protected structure; use a preserving Range/OOXML edit")
        if child.find(".//w:rPrChange", NS) is not None:
            raise ValueError("tracked character formatting requires a preserving edit")
        if any(item.tag not in {qn(W, "rPr"), qn(W, "t")} for item in child):
            raise ValueError("run contains a field, object or other non-text content; refusing replacement")
        runs.append(child)
    if require_uniform and len({run_format(run) for run in runs}) > 1:
        raise ValueError("whole-paragraph replacement would discard mixed character formatting")
    return runs


def run_format(run: ET.Element) -> bytes:
    properties = run.find("w:rPr", NS)
    return ET.tostring(properties) if properties is not None else b""


def replace_paragraph_text(paragraph: ET.Element, value: str) -> None:
    plain_text_runs(paragraph, require_uniform=True)
    ppr = paragraph.find("w:pPr", NS)
    first_rpr = paragraph.find("w:r/w:rPr", NS)
    ppr_copy = deepcopy(ppr) if ppr is not None else None
    rpr_copy = deepcopy(first_rpr) if first_rpr is not None else None
    for child in list(paragraph):
        paragraph.remove(child)
    if ppr_copy is not None:
        paragraph.append(ppr_copy)
    run = ET.SubElement(paragraph, qn(W, "r"))
    if rpr_copy is not None:
        run.append(rpr_copy)
    text = ET.SubElement(run, qn(W, "t"))
    text.set(qn(XML, "space"), "preserve")
    text.text = value


def replace_across_text_nodes(paragraph: ET.Element, source: str, target: str) -> int:
    if not source or source == target:
        return 0
    full = text_of(paragraph)
    matches = list(re.finditer(re.escape(source), full))
    if not matches:
        return 0
    runs = plain_text_runs(paragraph)
    nodes = [node for run in runs for node in run.findall("w:t", NS)]
    formats = [run_format(run) for run in runs for _ in run.findall("w:t", NS)]
    spans = []
    cursor = 0
    for node in nodes:
        value = node.text or ""
        spans.append((cursor, cursor + len(value)))
        cursor += len(value)
    locations = []
    for match in matches:
        start, end = match.span()
        first = next(i for i, (_, right) in enumerate(spans) if right > start)
        last = next(i for i, (left, right) in enumerate(spans) if left < end <= right)
        if len(set(formats[first:last + 1])) > 1:
            raise ValueError("literal replacement crosses character formats; use a targeted preserving edit")
        locations.append((start, end, first, last))
    # Reverse original matches so replacement text containing the source cannot loop.
    for start, end, first, last in reversed(locations):
        first_value = nodes[first].text or ""
        last_value = nodes[last].text or ""
        first_offset = start - spans[first][0]
        last_offset = end - spans[last][0]
        if first == last:
            nodes[first].text = first_value[:first_offset] + target + first_value[last_offset:]
        else:
            nodes[first].text = first_value[:first_offset] + target + last_value[last_offset:]
            for index in range(first + 1, last + 1):
                nodes[index].text = ""
        nodes[first].set(qn(XML, "space"), "preserve")
    return len(locations)


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch selected plain-text DOCX paragraphs without rebuilding the document.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("plan", type=Path, help="UTF-8 JSON patch plan")
    args = parser.parse_args()

    source = args.source.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if source == output:
        raise ValueError("output must differ from source")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite: {output}")
    plan = json.loads(args.plan.expanduser().resolve().read_text(encoding="utf-8"))
    expected_hash = plan.get("source_sha256")
    if expected_hash and hashlib.sha256(source.read_bytes()).hexdigest() != expected_hash:
        raise ValueError("source hash does not match the reviewed patch plan")

    with ZipFile(source) as package:
        infos = package.infolist()
        members = {info.filename: package.read(info.filename) for info in infos}
    root = parse_xml_preserving_prefixes(members["word/document.xml"])
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("document body not found")

    report = {"paragraphs": [], "literal_replacements": []}
    for item in plan.get("paragraphs", []):
        anchor = find_paragraph(body, item["anchor"], item.get("match", "exact"))
        target = anchor if item.get("target", "self") == "self" else following_paragraph(body, anchor)
        old = text_of(target)
        if "expected_old_text" in item and old != item["expected_old_text"]:
            raise ValueError(f"paragraph at {item['anchor']!r} changed since the plan was prepared")
        replace_paragraph_text(target, item["text"])
        report["paragraphs"].append({"anchor": item["anchor"], "old_length": len(old), "new_length": len(item["text"])})

    for item in plan.get("literal_replacements", []):
        count = 0
        candidates = [find_paragraph(body, item["anchor"], item.get("match", "exact"))] if "anchor" in item else root.findall(".//w:p", NS)
        for paragraph in candidates:
            count += replace_across_text_nodes(paragraph, item["source"], item["target"])
        expected = item.get("expected_count")
        if expected is not None and count != expected:
            raise ValueError(f"replacement {item['source']!r}: expected {expected}, got {count}")
        report["literal_replacements"].append({"source": item["source"], "count": count})

    members["word/document.xml"] = serialize_xml(root)
    write_docx_exclusive(output, infos, members)
    report.update({"source": str(source), "output": str(output)})
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
