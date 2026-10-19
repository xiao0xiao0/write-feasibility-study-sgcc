from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import Counter
from pathlib import Path
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET

from docx_integrity_checks import audit_package, audit_citations, compare_protected_parts


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


def emit_report(result: dict, output: Path | None) -> None:
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if output:
        destination = output.expanduser().resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)


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
    toc_options = parser.add_mutually_exclusive_group()
    toc_options.add_argument("--allow-toc", action="store_true", help="Allow TOC fields; existing TOCs are still preserved with a baseline")
    toc_options.add_argument("--forbid-toc", action="store_true", help="Explicitly require no TOC fields")
    parser.add_argument("--no-fail", action="store_true", help="Always exit 0 after reporting")
    args = parser.parse_args()

    path = args.docx.expanduser().resolve()
    config = {}
    if args.config:
        config = json.loads(args.config.expanduser().resolve().read_text(encoding="utf-8"))
    allow_mathtype_changes = bool(config.get("allow_mathtype_object_changes", False))
    expected_mathtype = config.get("expected_mathtype_ole_count")
    expected_numbered = config.get("expected_numbered_equation_count")
    expected_embeddings = config.get("expected_embedding_part_count")
    expected_number_format = config.get("expected_mathtype_number_format")
    # Project-specific rejected terms belong to the project config, not global bans.
    banned_phrases = config.get("banned_phrases", [])

    issues = []

    def add(severity: str, code: str, message: str, context: str = ""):
        issues.append({"severity": severity, "code": code, "message": message, "context": context[:240]})

    try:
        names, members, bad = package(path)
    except Exception as exc:
        add("blocker", "invalid-docx", str(exc))
        result = {"path": str(path), "issues": issues, "status": "fail"}
        emit_report(result, args.output)
        return 0 if args.no_fail else 2

    required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels"}
    for missing in sorted(required - set(names)):
        add("blocker", "missing-part", f"missing required DOCX part: {missing}")
    if bad:
        add("blocker", "zip-integrity", f"corrupt ZIP member: {bad}")
    if len(names) != len(set(names)):
        add("blocker", "duplicate-zip-member", "duplicate DOCX ZIP part names")
    roots = audit_package(members, add)
    base_members = None
    if args.baseline:
        try:
            base_names, base_members, base_bad = package(args.baseline.expanduser().resolve())
            if base_bad or len(base_names) != len(set(base_names)):
                add("blocker", "baseline-integrity", "baseline has corrupt or duplicate ZIP members")
            audit_package(base_members, lambda severity, code, message, context="": add(severity, "baseline-" + code, message, context))
        except Exception as exc:
            add("blocker", "baseline-integrity", str(exc))
    if any(issue["severity"] == "blocker" for issue in issues):
        result = {"path": str(path), "issues": issues, "status": "fail"}
        emit_report(result, args.output)
        return 0 if args.no_fail else 2

    document = roots["word/document.xml"]
    styles_root = roots["word/styles.xml"]
    if document.tag != qn("document"):
        add("blocker", "unsupported-document-namespace", "expected transitional WordprocessingML document root")
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

    reference_audit = audit_citations(document, styles_root, add, config)
    reference_audit.pop("bibliography_paragraph_ids")

    def toc_count(parts):
        root = ET.fromstring(parts["word/document.xml"])
        codes = [item.text or "" for item in root.findall(".//w:instrText", NS)]
        codes += [item.get(qn("instr"), "") for item in root.findall(".//w:fldSimple", NS)]
        return sum(bool(re.search(r"\bTOC\b", code, re.I)) for code in codes)

    toc_policy = config.get("toc_policy", "preserve" if base_members is not None else "allow")
    if args.forbid_toc or config.get("allow_toc") is False:
        toc_policy = "forbid"
    elif args.allow_toc and base_members is None:
        toc_policy = "allow"
    if toc_policy not in {"preserve", "allow", "forbid"}:
        raise ValueError("toc_policy must be preserve, allow, or forbid")
    current_toc_count = toc_count(members)
    if toc_policy == "forbid" and current_toc_count:
        add("major", "unexpected-toc", "TOC field is explicitly forbidden for this task")
    if toc_policy == "preserve" and base_members is not None and current_toc_count != toc_count(base_members):
        add("major", "toc-not-preserved", "existing TOC field count changed without an explicit policy change")

    for phrase in banned_phrases:
        if phrase and phrase in full_text:
            add("major", "banned-phrase", f"prohibited explanatory phrase remains: {phrase}")

    heading_styles = {"SG-章标题", "SG-二级标题", "SG-三级标题", "SG-题内标题"}
    caption_styles = {"SG-图题", "SG-表题"}
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
            paragraph_instr = "".join(
                node.text or "" for node in paragraph.findall(".//w:instrText", NS)
            )
            is_numbered = "MACROBUTTON MTPlaceRef" in paragraph_instr
            # Formula -> actual Tab -> number and column width belong to the MathType validator.
            static_label = re.search(r"\(\d+-\d+\)\s*$", node_text(paragraph))
            if static_label and not is_numbered:
                add("major", "static-formula-number", "formula number is static text rather than a MathType field", value)

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

    part_changes = {"changed_parts": [], "added_parts": [], "removed_parts": []}
    if base_members is not None:
        base_document = ET.fromstring(base_members["word/document.xml"])
        base_counts = part_counts(base_document, base_names)
        expected_counts = config.get("expected_counts", {}).copy()
        if allow_mathtype_changes and expected_mathtype is not None and expected_embeddings is not None:
            original_mathtype = sum(node.get("ProgID") == "Equation.DSMT4" for node in base_document.findall(".//o:OLEObject", NS))
            expected_counts.setdefault("ole_objects", base_counts["ole_objects"] + int(expected_mathtype) - original_mathtype)
            expected_counts.setdefault("embedding_parts", int(expected_embeddings))
        for key in ["drawings", "header_parts", "footer_parts", "sections", "ole_objects", "embedding_parts"]:
            expected = int(expected_counts.get(key, base_counts[key]))
            if counts[key] != expected:
                add("blocker", "package-count-change", f"{key}: expected {expected}, found {counts[key]}")
        part_changes = compare_protected_parts(base_members, members, config, add)

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
        "citation_runs_checked": reference_audit["citation_groups_checked"],
        "citation_runs_not_superscript": reference_audit["citation_groups_not_superscript"],
        "references": reference_audit,
        "toc_policy": toc_policy,
        "toc_fields": current_toc_count,
        "part_changes": part_changes,
        "word_mathtype_com": "not_run",
        "mathtype_layout_validation": "external_required" if number_field_count else "not_applicable",
        "severity_counts": dict(severity_counts),
        "issues": issues,
    }
    emit_report(result, args.output)
    if args.no_fail:
        return 0
    return 2 if status == "fail" else 0


if __name__ == "__main__":
    raise SystemExit(main())
