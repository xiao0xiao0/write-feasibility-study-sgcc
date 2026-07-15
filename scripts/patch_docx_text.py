from __future__ import annotations

import argparse
import json
import os
from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

from ooxml_utils import parse_xml_preserving_prefixes, serialize_xml


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


def replace_paragraph_text(paragraph: ET.Element, value: str) -> None:
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
    changed = 0
    while True:
        nodes = paragraph.findall(".//w:t", NS)
        full = "".join(node.text or "" for node in nodes)
        start = full.find(source)
        if start < 0:
            return changed
        end = start + len(source)
        spans = []
        cursor = 0
        for node in nodes:
            value = node.text or ""
            spans.append((cursor, cursor + len(value)))
            cursor += len(value)
        first = next(i for i, (_, right) in enumerate(spans) if right > start)
        last = next(i for i, (left, right) in enumerate(spans) if left < end <= right)
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
        changed += 1


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
        replace_paragraph_text(target, item["text"])
        report["paragraphs"].append({"anchor": item["anchor"], "old_length": len(old), "new_length": len(item["text"])})

    for item in plan.get("literal_replacements", []):
        count = 0
        for paragraph in root.findall(".//w:p", NS):
            count += replace_across_text_nodes(paragraph, item["source"], item["target"])
        expected = item.get("expected_count")
        if expected is not None and count != expected:
            raise ValueError(f"replacement {item['source']!r}: expected {expected}, got {count}")
        report["literal_replacements"].append({"source": item["source"], "count": count})

    members["word/document.xml"] = serialize_xml(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(output.suffix + ".tmp")
    with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as package:
        for info in infos:
            package.writestr(info, members[info.filename])
    os.replace(temporary, output)
    report.update({"source": str(source), "output": str(output)})
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
