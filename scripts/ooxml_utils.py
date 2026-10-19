from __future__ import annotations

import io
import os
from copy import copy
from pathlib import Path
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML = "http://www.w3.org/XML/1998/namespace"
XSI = "http://www.w3.org/2001/XMLSchema-instance"


def _namespace_uri(name: str) -> str | None:
    if name.startswith("{") and "}" in name:
        return name[1:].split("}", 1)[0]
    return None


def _declarations(node: ET.Element) -> dict[str, str]:
    result = {}
    for name, value in node.attrib.items():
        if name == "xmlns":
            result[""] = value
        elif name.startswith("xmlns:"):
            result[name[6:]] = value
    return result


def _validate_prefix_values(node: ET.Element, scope: dict[str, str]) -> None:
    prefix_lists = {f"{{{MC}}}Ignorable", f"{{{MC}}}MustUnderstand"}
    qname_lists = {f"{{{MC}}}ProcessContent", f"{{{MC}}}PreserveElements", f"{{{MC}}}PreserveAttributes", f"{{{XSI}}}type"}
    for name, value in node.attrib.items():
        if name in prefix_lists or node.tag == f"{{{MC}}}Choice" and name == "Requires":
            for prefix in value.split():
                if prefix not in scope or not scope[prefix]:
                    raise ValueError(f"undeclared compatibility prefix {prefix!r} in {name}")
        elif name in qname_lists:
            for token in value.split():
                if ":" in token:
                    prefix = token.split(":", 1)[0]
                    if prefix not in scope or not scope[prefix]:
                        raise ValueError(f"undeclared QName prefix {prefix!r} in {name}")


def parse_xml_preserving_prefixes(data: bytes) -> ET.Element:
    # XML namespace declarations may be used ONLY by string-valued MC attributes.
    # Retain every declaration on its original element, including aliases and
    # nested/default scopes. Literal xmlns attributes also survive deepcopy.
    # No process-global register_namespace calls: documents can use the same
    # prefix for different URIs without changing one another's serialization.
    if b"<!DOCTYPE" in data.upper():
        raise ValueError("DTD declarations are not supported in Office XML parts")
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
    events = ET.iterparse(io.BytesIO(data), events=("start-ns", "start"), parser=parser)
    pending = []
    for event, item in events:
        if event == "start-ns":
            pending.append(item)
        else:
            for prefix, uri in pending:
                item.set("xmlns:" + prefix if prefix else "xmlns", uri)
            pending.clear()
    return events.root


def _lexical_clone(node: ET.Element, inherited: dict[str, str]) -> ET.Element:
    if node.tag in {ET.Comment, ET.ProcessingInstruction}:
        result = ET.Element(node.tag)
        result.text, result.tail = node.text, node.tail
        return result
    scope = dict(inherited)
    declarations = _declarations(node)
    scope.update(declarations)
    _validate_prefix_values(node, scope)

    def lexical(name, attribute=False):
        if isinstance(name, ET.QName):
            name = str(name)
        uri = _namespace_uri(name)
        if uri is None:
            if not attribute and scope.get(""):
                declarations[""] = ""
                scope[""] = ""
            return name
        local = name.split("}", 1)[1]
        if uri == XML:
            return "xml:" + local
        if not attribute and scope.get("") == uri:
            return local
        for prefix, bound_uri in scope.items():
            if prefix and bound_uri == uri:
                return prefix + ":" + local
        counter = 1
        while "sgccns" + str(counter) in scope:
            counter += 1
        prefix = "sgccns" + str(counter)
        scope[prefix] = uri
        declarations[prefix] = uri
        return prefix + ":" + local

    result = ET.Element(lexical(node.tag))
    for name, value in node.attrib.items():
        if name == "xmlns" or name.startswith("xmlns:"):
            continue
        result.set(lexical(name, attribute=True), value)
    for prefix, uri in declarations.items():
        result.set("xmlns:" + prefix if prefix else "xmlns", uri)
    result.text, result.tail = node.text, node.tail
    for child in node:
        result.append(_lexical_clone(child, scope))
    return result


def serialize_xml(root: ET.Element) -> bytes:
    # All qualified names are rendered lexically before ET writes the XML, so ET
    # cannot drop value-only namespace declarations or create duplicate xmlns.
    clone = _lexical_clone(root, {"xml": XML})
    payload = ET.tostring(clone, encoding="utf-8", xml_declaration=True)
    ET.fromstring(payload)  # fail closed before any DOCX output is published
    return payload


def _publish_new_file(temporary: Path, output: Path) -> None:
    if os.name == "nt":
        # Windows rename refuses an existing destination, including a file
        # created after the caller's initial existence check.
        os.rename(temporary, output)
    else:
        # POSIX rename replaces destinations; a hard link provides no-clobber
        # publication instead. Both paths use the same directory/filesystem.
        os.link(temporary, output)
        temporary.unlink()


def write_docx_exclusive(output: Path, infos, members: dict[str, bytes]) -> None:
    """Publish a completed ZIP once, with a unique temp and no replacement."""
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output.parent, prefix=".sgcc-", suffix=".tmp", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as package:
            for info in infos:
                package.writestr(copy(info), members[info.filename])
        with ZipFile(temporary) as package:
            if package.testzip():
                raise ValueError("generated DOCX failed ZIP integrity verification")
        _publish_new_file(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
