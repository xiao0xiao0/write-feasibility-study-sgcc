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
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O = "urn:schemas-microsoft-com:office:office"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS = {"w": W, "m": M, "o": O, "cp": CUSTOM, "vt": VT}


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def node_text(node: ET.Element) -> str:
    out = []
    for item in node.iter():
        if item.tag == qn("t") and item.text:
            out.append(item.text)
        elif item.tag == qn("tab"):
            out.append("\t")
        elif item.tag == qn("br"):
            out.append("\n")
    return "".join(out)


def paragraph_style(paragraph: ET.Element) -> str:
    node = paragraph.find("w:pPr/w:pStyle", NS)
    return node.get(qn("val"), "") if node is not None else ""


def ppr_values(ppr: ET.Element | None) -> dict[str, str]:
    if ppr is None:
        return {}
    result = {}
    jc = ppr.find("w:jc", NS)
    if jc is not None:
        result["jc"] = jc.get(qn("val"), "")
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        for key in ["firstLine", "firstLineChars", "hanging", "hangingChars", "left", "leftChars", "right", "rightChars"]:
            value = ind.get(qn(key))
            if value is not None:
                result[key] = value
    return result


def load_styles(root: ET.Element):
    definitions = {}
    for style in root.findall("w:style", NS):
        sid = style.get(qn("styleId"), "")
        name_node = style.find("w:name", NS)
        based_node = style.find("w:basedOn", NS)
        definitions[sid] = {
            "name": name_node.get(qn("val"), sid) if name_node is not None else sid,
            "based_on": based_node.get(qn("val"), "") if based_node is not None else "",
            "ppr": ppr_values(style.find("w:pPr", NS)),
        }
    return definitions


def effective_style_ppr(style_id: str, definitions: dict, stack=None) -> dict[str, str]:
    if not style_id or style_id not in definitions:
        return {}
    stack = set() if stack is None else set(stack)
    if style_id in stack:
        return {}
    stack.add(style_id)
    current = definitions[style_id]
    result = effective_style_ppr(current["based_on"], definitions, stack)
    result.update(current["ppr"])
    return result


def effective_paragraph_ppr(paragraph: ET.Element, definitions: dict) -> dict[str, str]:
    result = effective_style_ppr(paragraph_style(paragraph), definitions)
    result.update(ppr_values(paragraph.find("w:pPr", NS)))
    return result


def package(path: Path):
    try:
        with ZipFile(path) as z:
            names = z.namelist()
            bad = z.testzip()
            members = {name: z.read(name) for name in names}
    except BadZipFile as exc:
        raise ValueError(f"invalid DOCX ZIP: {path}") from exc
    return names, members, bad


def part_counts(document: ET.Element, names: list[str]) -> dict[str, int]:
    return {
        "drawings": len(document.findall(".//w:drawing", NS)),
        "ole_objects": len(document.findall(".//w:object", NS)),
        "media_parts": len([name for name in names if name.startswith("word/media/")]),
        "embedding_parts": len([name for name in names if name.startswith("word/embeddings/")]),
        "header_parts": len([name for name in names if re.fullmatch(r"word/header\d+\.xml", name)]),
        "footer_parts": len([name for name in names if re.fullmatch(r"word/footer\d+\.xml", name)]),
        "sections": len(document.findall(".//w:sectPr", NS)),
    }


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
    parser = argparse.ArgumentParser(description="Quality-assure an SGCC feasibility-study DOCX.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--config", type=Path, help="Optional UTF-8 JSON configuration")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--allow-toc", action="store_true")
    parser.add_argument("--no-fail", action="store_true", help="Always exit 0 after reporting")
    args = parser.parse_args()

    path = args.docx.expanduser().resolve()
    config = {}
    if args.config:
        config = json.loads(args.config.expanduser().resolve().read_text(encoding="utf-8"))
    allowed_media = set(config.get("allowed_changed_media", []))
    allow_mathtype_changes = bool(config.get("allow_mathtype_object_changes", False))
    expected_mathtype = config.get("expected_mathtype_ole_count")
    expected_numbered = config.get("expected_numbered_equation_count")
    expected_embeddings = config.get("expected_embedding_part_count")
    expected_number_format = config.get("expected_mathtype_number_format")
    banned_phrases = [
        "可为章节组织、技术路线和算例设计提供参考",
    ] + config.get("banned_phrases", [])

    issues = []

    def add(severity: str, code: str, message: str, context: str = ""):
        issues.append({"severity": severity, "code": code, "message": message, "context": context[:240]})

    try:
        names, members, bad = package(path)
    except Exception as exc:
        add("blocker", "invalid-docx", str(exc))
        result = {"path": str(path), "issues": issues, "status": "fail"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if args.no_fail else 2

    required = {"[Content_Types].xml", "word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels"}
    for missing in sorted(required - set(names)):
        add("blocker", "missing-part", f"missing required DOCX part: {missing}")
    if bad:
        add("blocker", "zip-integrity", f"corrupt ZIP member: {bad}")
    if any(issue["severity"] == "blocker" for issue in issues):
        result = {"path": str(path), "issues": issues, "status": "fail"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if args.no_fail else 2

    document = ET.fromstring(members["word/document.xml"])
    styles_root = ET.fromstring(members["word/styles.xml"])
    definitions = load_styles(styles_root)
    style_names = {sid: item["name"] for sid, item in definitions.items()}
    paragraphs = document.findall(".//w:p", NS)
    full_text = "\n".join(node_text(p) for p in paragraphs)
    counts = part_counts(document, names)
    prog_ids = [node.attrib.get("ProgID", "") for node in document.findall(".//o:OLEObject", NS)]
    mathtype_ole_count = sum(value == "Equation.DSMT4" for value in prog_ids)
    other_ole_progids = [value for value in prog_ids if value != "Equation.DSMT4"]
    omath_count = len(document.findall(".//m:oMath", NS))
    instruction_text = "".join(node.text or "" for node in document.findall(".//w:instrText", NS))
    number_field_count = instruction_text.count("MACROBUTTON MTPlaceRef")
    mteqn_mentions = instruction_text.count("SEQ MTEqn")
    mtchap_mentions = instruction_text.count("SEQ MTChap")
    mtsec_mentions = instruction_text.count("SEQ MTSec")
    number_format = custom_number_format(members)
    visible_equation_numbers = re.findall(
        r"(?:\(\s*\d+(?:[-.]\d+){0,2}\s*\)|（\s*\d+(?:[-.]\d+){0,2}\s*）)",
        instruction_text,
    )

    bibliography_paragraph_ids = set()
    body = document.find("w:body", NS)
    in_bibliography = False
    if body is not None:
        for child in body:
            if child.tag != qn("p"):
                continue
            value = node_text(child).strip()
            style_name = style_names.get(paragraph_style(child), paragraph_style(child))
            if value.startswith("参考文献"):
                in_bibliography = True
                bibliography_paragraph_ids.add(id(child))
                continue
            if in_bibliography and style_name == "SG-章标题":
                in_bibliography = False
            if in_bibliography:
                bibliography_paragraph_ids.add(id(child))

    all_xml = "\n".join(data.decode("utf-8", errors="ignore") for name, data in members.items() if name.endswith(".xml"))
    if not (args.allow_toc or config.get("allow_toc", False)) and re.search(r"TOC\s+\\", all_xml, re.I):
        add("major", "unexpected-toc", "document contains a TOC field")

    for phrase in banned_phrases:
        if phrase and phrase in full_text:
            add("major", "banned-phrase", f"prohibited explanatory phrase remains: {phrase}")

    heading_styles = {"SG-章标题", "SG-二级标题", "SG-三级标题", "SG-题内标题"}
    caption_styles = {"SG-图题", "SG-表题"}
    citation_runs = 0
    citation_not_superscript = 0
    for paragraph in paragraphs:
        value = node_text(paragraph).strip()
        sid = paragraph_style(paragraph)
        style_name = style_names.get(sid, sid)
        effective = effective_paragraph_ppr(paragraph, definitions)
        first_line = effective.get("firstLine")
        first_line_chars = effective.get("firstLineChars")

        if style_name in heading_styles:
            if (first_line not in {None, "0"}) or (first_line_chars not in {None, "0"}):
                add("major", "heading-indent", f"{style_name} has nonzero first-line indent", value)
        if style_name in caption_styles or re.match(r"^[图表]\s*\d+-\d+\s", value):
            if effective.get("jc") != "center":
                add("major", "caption-alignment", "caption is not centered", value)
            if (first_line not in {None, "0"}) or (first_line_chars not in {None, "0"}):
                add("major", "caption-indent", "caption has nonzero first-line indent", value)
        if style_name == "SG-正文" and re.match(r"^[ 　]{1,}", node_text(paragraph)):
            add("minor", "manual-body-indent", "body paragraph starts with manual spaces", value)
        if re.match(r"^(式中|其中)[，,:：]", value) or style_name == "MTVariableExplanation":
            left_indent = effective.get("left")
            left_chars = effective.get("leftChars")
            if (first_line not in {None, "0"}) or (first_line_chars not in {None, "0"}) or (left_indent not in {None, "0"}) or (left_chars not in {None, "0"}):
                add("major", "variable-explanation-indent", "variable-explanation paragraph must have zero left and first-line indent", value)
        if sid == "MTDisplayEquation" or style_name == "MTDisplayEquation":
            tabs = len(paragraph.findall(".//w:tab", NS))
            paragraph_instr = "".join(
                node.text or "" for node in paragraph.findall(".//w:instrText", NS)
            )
            is_numbered = "MACROBUTTON MTPlaceRef" in paragraph_instr
            if is_numbered and tabs < 2:
                add("major", "formula-tabs", f"numbered MathType paragraph has {tabs} real tabs; expected at least 2", value)
            static_label = re.search(r"\(\d+-\d+\)\s*$", node_text(paragraph))
            if static_label and not is_numbered:
                add("major", "static-formula-number", "formula number is static text rather than a MathType field", value)
        for run in paragraph.findall("w:r", NS):
            run_text = "".join((node.text or "") for node in run.findall("w:t", NS))
            if id(paragraph) not in bibliography_paragraph_ids and "$" not in run_text and re.search(r"\[\d+(?:[-,—]\d+)*\]", run_text):
                citation_runs += 1
                vert = run.find("w:rPr/w:vertAlign", NS)
                if vert is None or vert.get(qn("val")) != "superscript":
                    citation_not_superscript += 1
                    add("major", "citation-superscript", "numeric citation is not superscript", run_text)

    residual_placeholders = re.findall(r"⟦MT[IDN]:.*?⟧", full_text, flags=re.S)
    residual_display_tex = bool(re.search(r"\\\[.*?\\\]", full_text, flags=re.S))
    residual_inline_tex = bool(re.search(r"\$[^$\r\n]{1,200}\$", full_text))
    if residual_placeholders:
        add("major", "mathtype-placeholder", "unconverted MathType placeholders remain")
    if residual_display_tex or residual_inline_tex:
        add("major", "residual-tex", "unconverted display or inline TeX remains in the final DOCX")
    if omath_count:
        add("major", "word-omath", f"Word OMath objects found: {omath_count}")
    if expected_mathtype is not None and mathtype_ole_count != int(expected_mathtype):
        add("major", "mathtype-count", f"expected {expected_mathtype} Equation.DSMT4 objects, found {mathtype_ole_count}")
    if expected_numbered is not None and number_field_count != int(expected_numbered):
        add("major", "mathtype-number-count", f"expected {expected_numbered} MathType numbers, found {number_field_count}")
    if expected_embeddings is not None and counts["embedding_parts"] != int(expected_embeddings):
        add("blocker", "embedding-count", f"expected {expected_embeddings} embedding parts, found {counts['embedding_parts']}")
    if expected_number_format is not None and number_format != expected_number_format:
        add("major", "mathtype-number-format", f"expected MTEquationNumber2={expected_number_format!r}, found {number_format!r}")

    if args.baseline:
        baseline_path = args.baseline.expanduser().resolve()
        base_names, base_members, base_bad = package(baseline_path)
        if base_bad:
            add("blocker", "baseline-integrity", f"baseline has corrupt member: {base_bad}")
        else:
            base_document = ET.fromstring(base_members["word/document.xml"])
            base_counts = part_counts(base_document, base_names)
            fixed_keys = ["drawings", "header_parts", "footer_parts", "sections"]
            if not allow_mathtype_changes:
                fixed_keys.extend(["ole_objects", "embedding_parts"])
            for key in fixed_keys:
                if counts[key] != base_counts[key]:
                    add("blocker", "package-count-change", f"{key} changed from {base_counts[key]} to {counts[key]}")
            current_media = {name for name in names if name.startswith("word/media/")}
            base_media = {name for name in base_names if name.startswith("word/media/")}
            if current_media != base_media:
                add("blocker", "media-set-change", "media part names changed unexpectedly")
            for name in sorted(current_media & base_media):
                if sha256(members[name]) != sha256(base_members[name]) and name not in allowed_media:
                    add("major", "media-content-change", f"media changed without allowance: {name}")
            current_embeddings = {name for name in names if name.startswith("word/embeddings/")}
            base_embeddings = {name for name in base_names if name.startswith("word/embeddings/")}
            if not allow_mathtype_changes:
                if current_embeddings != base_embeddings:
                    add("blocker", "embedding-set-change", "embedding part names changed")
                else:
                    for name in sorted(current_embeddings):
                        if sha256(members[name]) != sha256(base_members[name]):
                            add("blocker", "embedding-content-change", f"embedding changed: {name}")

    severity_counts = Counter(issue["severity"] for issue in issues)
    status = "fail" if severity_counts["blocker"] or severity_counts["major"] else ("warning" if severity_counts["minor"] else "pass")
    result = {
        "path": str(path),
        "baseline": str(args.baseline.expanduser().resolve()) if args.baseline else None,
        "status": status,
        "counts": counts,
        "mathtype_ole_count": mathtype_ole_count,
        "other_ole_progids": other_ole_progids,
        "omath_count": omath_count,
        "number_field_count": number_field_count,
        "mteqn_field_mentions": mteqn_mentions,
        "mtchap_field_mentions": mtchap_mentions,
        "mtsec_field_mentions": mtsec_mentions,
        "visible_equation_numbers": visible_equation_numbers,
        "mathtype_number_format": number_format,
        "citation_runs_checked": citation_runs,
        "citation_runs_not_superscript": citation_not_superscript,
        "severity_counts": dict(severity_counts),
        "issues": issues,
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.expanduser().resolve().write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    if args.no_fail:
        return 0
    return 2 if status == "fail" else 0


if __name__ == "__main__":
    raise SystemExit(main())
