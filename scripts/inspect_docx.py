from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import Counter
from pathlib import Path
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS = {"w": W, "r": R, "a": A, "wp": WP, "m": M, "o": O, "cp": CUSTOM, "vt": VT}


def qn(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def node_text(node: ET.Element) -> str:
    out: list[str] = []
    for child in node.iter():
        if child.tag == qn(W, "t") and child.text:
            out.append(child.text)
        elif child.tag == qn(W, "tab"):
            out.append("\t")
        elif child.tag == qn(W, "br"):
            out.append("\n")
    return "".join(out)


def paragraph_style(paragraph: ET.Element) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", NS)
    return node.get(qn(W, "val"), "") if node is not None else ""


def style_names(styles_root: ET.Element) -> dict[str, str]:
    result = {}
    for style in styles_root.findall("w:style", NS):
        sid = style.get(qn(W, "styleId"), "")
        name = style.find("w:name", NS)
        result[sid] = name.get(qn(W, "val"), sid) if name is not None else sid
    return result


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def custom_number_format(members: dict[str, bytes]) -> str | None:
    name = "docProps/custom.xml"
    if name not in members:
        return None
    root = ET.fromstring(members[name])
    for prop in root.findall("cp:property", NS):
        if prop.attrib.get("name") == "MTEquationNumber2":
            value = prop.find("vt:lpwstr", NS)
            return value.text if value is not None else ""
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect an SGCC feasibility-study DOCX without modifying it.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--output", type=Path, help="Optional JSON output path")
    parser.add_argument("--hash-parts", action="store_true", help="Hash media and embedding parts")
    parser.add_argument("--list-parts", action="store_true", help="List media and embedding part names")
    args = parser.parse_args()

    path = args.docx.expanduser().resolve()
    try:
        with ZipFile(path) as package:
            names = package.namelist()
            bad_member = package.testzip()
            members = {name: package.read(name) for name in names}
    except BadZipFile as exc:
        raise SystemExit(f"invalid DOCX ZIP: {path}") from exc

    required = ["[Content_Types].xml", "word/document.xml", "word/styles.xml"]
    missing = [name for name in required if name not in members]
    if missing:
        raise SystemExit(f"missing DOCX parts: {missing}")

    document = ET.fromstring(members["word/document.xml"])
    styles = ET.fromstring(members["word/styles.xml"])
    names_by_style = style_names(styles)
    paragraphs = document.findall(".//w:p", NS)

    style_counts = Counter()
    headings = []
    captions = []
    equations = []
    for paragraph in paragraphs:
        sid = paragraph_style(paragraph)
        style_counts[names_by_style.get(sid, sid or "(none)")] += 1
        text = re.sub(r"\s+", " ", node_text(paragraph)).strip()
        style_name = names_by_style.get(sid, sid)
        if style_name in {"SG-章标题", "SG-二级标题", "SG-三级标题", "SG-题内标题"}:
            headings.append({"style": style_name, "text": text})
        if style_name in {"SG-图题", "SG-表题"} or re.match(r"^[图表]\s*\d+-\d+", text):
            captions.append({"style": style_name, "text": text})
        if sid == "MTDisplayEquation" or style_name == "MTDisplayEquation":
            instruction = "".join(
                node.text or "" for node in paragraph.findall(".//w:instrText", NS)
            )
            static_label = re.search(r"\((\d+-\d+)\)\s*$", node_text(paragraph))
            equations.append({
                "automatic_number": "MACROBUTTON MTPlaceRef" in instruction,
                "static_label": static_label.group(1) if static_label else None,
                "tabs": len(paragraph.findall(".//w:tab", NS)),
                "text": node_text(paragraph)[-160:],
            })

    sections = []
    for section in document.findall(".//w:sectPr", NS):
        page = section.find("w:pgSz", NS)
        margins = section.find("w:pgMar", NS)
        sections.append({
            "page": {k.split("}")[-1]: v for k, v in (page.attrib if page is not None else {}).items()},
            "margins": {k.split("}")[-1]: v for k, v in (margins.attrib if margins is not None else {}).items()},
        })

    all_xml_text = "\n".join(
        data.decode("utf-8", errors="ignore")
        for name, data in members.items()
        if name.endswith(".xml")
    )
    media = sorted(name for name in names if name.startswith("word/media/"))
    embeddings = sorted(name for name in names if name.startswith("word/embeddings/"))
    prog_ids = [node.attrib.get("ProgID", "") for node in document.findall(".//o:OLEObject", NS)]
    instruction_text = "".join(node.text or "" for node in document.findall(".//w:instrText", NS))
    full_text = "\n".join(node_text(paragraph) for paragraph in paragraphs)
    result = {
        "path": str(path),
        "size_bytes": path.stat().st_size,
        "zip_integrity": "pass" if bad_member is None else f"bad member: {bad_member}",
        "paragraphs": len(paragraphs),
        "tables": len(document.findall(".//w:tbl", NS)),
        "drawings": len(document.findall(".//w:drawing", NS)),
        "ole_objects": len(document.findall(".//w:object", NS)),
        "mathtype_ole_objects": sum(value == "Equation.DSMT4" for value in prog_ids),
        "other_ole_progids": [value for value in prog_ids if value != "Equation.DSMT4"],
        "word_omath_objects": len(document.findall(".//m:oMath", NS)),
        "mathtype_number_fields": instruction_text.count("MACROBUTTON MTPlaceRef"),
        "mteqn_field_mentions": instruction_text.count("SEQ MTEqn"),
        "mtchap_field_mentions": instruction_text.count("SEQ MTChap"),
        "mtsec_field_mentions": instruction_text.count("SEQ MTSec"),
        "mathtype_number_format": custom_number_format(members),
        "residual_mathtype_placeholders": re.findall(r"⟦MT[IDN]:.*?⟧", full_text, flags=re.S),
        "residual_display_tex": bool(re.search(r"\\\[.*?\\\]", full_text, flags=re.S)),
        "residual_inline_tex_risk": bool(re.search(r"\$[^$\r\n]{1,200}\$", full_text)),
        "media_parts": len(media),
        "embedding_parts": len(embeddings),
        "header_parts": len([name for name in names if re.fullmatch(r"word/header\d+\.xml", name)]),
        "footer_parts": len([name for name in names if re.fullmatch(r"word/footer\d+\.xml", name)]),
        "has_toc_field": bool(re.search(r"TOC\s+\\", all_xml_text, re.I)),
        "style_counts": dict(style_counts),
        "headings": headings,
        "captions": captions,
        "equations": equations,
        "sections": sections,
    }
    if args.list_parts:
        result["media"] = media
        result["embeddings"] = embeddings
    if args.hash_parts:
        result["part_hashes"] = {
            name: sha256(members[name]) for name in media + embeddings
        }

    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.expanduser().resolve().write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
