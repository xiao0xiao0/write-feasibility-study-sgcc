from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from ooxml_utils import parse_xml_preserving_prefixes, serialize_xml, write_docx_exclusive


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]
PPR_INDEX = {name: index for index, name in enumerate(PPR_ORDER)}


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


def text_of(node: ET.Element) -> str:
    return "".join((item.text or "") for item in node.findall(".//w:t", NS)).strip()


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def ensure_ordered_child(parent: ET.Element, name: str) -> ET.Element:
    existing = parent.find(f"w:{name}", NS)
    if existing is not None:
        return existing
    node = ET.Element(qn(name))
    desired = PPR_INDEX[name]
    insert_at = len(parent)
    for index, child in enumerate(parent):
        # XML comments and processing instructions have callable tags. Preserve
        # them in place; only element nodes participate in property ordering.
        if not isinstance(child.tag, str):
            continue
        local = child.tag.split("}")[-1]
        if PPR_INDEX.get(local, len(PPR_ORDER)) > desired:
            insert_at = index
            break
    parent.insert(insert_at, node)
    return node


def set_style(paragraph: ET.Element, style_id: str) -> None:
    ppr = ensure_ppr(paragraph)
    node = ppr.find("w:pStyle", NS)
    if node is None:
        node = ET.Element(qn("pStyle"))
        ppr.insert(0, node)
    node.set(qn("val"), style_id)


def zero_indent(paragraph: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    node = ensure_ordered_child(ppr, "ind")
    node.set(qn("firstLine"), "0")
    node.set(qn("firstLineChars"), "0")
    node.attrib.pop(qn("hanging"), None)
    node.attrib.pop(qn("hangingChars"), None)


def set_alignment(paragraph: ET.Element, value: str) -> None:
    ppr = ensure_ppr(paragraph)
    node = ensure_ordered_child(ppr, "jc")
    node.set(qn("val"), value)


def style_name_map(styles: ET.Element) -> dict[str, str]:
    result = {}
    for style in styles.findall("w:style", NS):
        sid = style.get(qn("styleId"), "")
        name = style.find("w:name", NS)
        if name is not None:
            result[name.get(qn("val"), "")] = sid
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply SGCC heading/caption zero-indent and alignment rules to a DOCX clone.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--verbose", action="store_true", help="Include every changed paragraph in JSON output")
    args = parser.parse_args()
    source = args.source.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if source == output:
        raise ValueError("output must differ from source")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite: {output}")

    with ZipFile(source) as package:
        infos = package.infolist()
        members = {info.filename: package.read(info.filename) for info in infos}
    document = parse_xml_preserving_prefixes(members["word/document.xml"])
    styles = ET.fromstring(members["word/styles.xml"])
    by_name = style_name_map(styles)

    rules = [
        (re.compile(r"^[一二三四五六七八九十]+、"), "SG-章标题", "left"),
        (re.compile(r"^\d+\.\d+\.\d+\s"), "SG-三级标题", "left"),
        (re.compile(r"^\d+\.\d+\s"), "SG-二级标题", "left"),
        (re.compile(r"^图\s*\d+-\d+\s"), "SG-图题", "center"),
        (re.compile(r"^表\s*\d+-\d+\s"), "SG-表题", "center"),
    ]
    changed = []
    body = document.find("w:body", NS)
    if body is None:
        raise ValueError("document body not found")
    for paragraph in [child for child in body if child.tag == qn("p")]:
        value = text_of(paragraph)
        applied = None
        for pattern, style_name, alignment in rules:
            if pattern.match(value):
                style_id = by_name.get(style_name)
                if style_id is None and style_name == "SG-表题":
                    style_id = by_name.get("SG-图题")
                if style_id:
                    set_style(paragraph, style_id)
                zero_indent(paragraph)
                set_alignment(paragraph, alignment)
                applied = style_name
                break
        if applied is None:
            pstyle = paragraph.find("w:pPr/w:pStyle", NS)
            style_id = pstyle.get(qn("val"), "") if pstyle is not None else ""
            style_name = next((name for name, sid in by_name.items() if sid == style_id), "")
            if style_name in {"SG-章标题", "SG-二级标题", "SG-三级标题", "SG-题内标题", "SG-图题", "SG-表题", "MTDisplayEquation"}:
                zero_indent(paragraph)
                if style_name in {"SG-图题", "SG-表题"}:
                    set_alignment(paragraph, "center")
                applied = style_name
            elif paragraph.find(".//w:drawing", NS) is not None and not value:
                zero_indent(paragraph)
                set_alignment(paragraph, "center")
                applied = "drawing-paragraph"
        if applied:
            changed.append({"rule": applied, "text": value[:100]})

    members["word/document.xml"] = serialize_xml(document)
    write_docx_exclusive(output, infos, members)
    report = {"source": str(source), "output": str(output), "changed_paragraphs": len(changed)}
    if args.verbose:
        report["changes"] = changed
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
