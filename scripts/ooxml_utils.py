from __future__ import annotations

import io
from xml.etree import ElementTree as ET


MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _namespace_uri(name: str) -> str | None:
    if name.startswith("{") and "}" in name:
        return name[1:].split("}", 1)[0]
    return None


def parse_xml_preserving_prefixes(data: bytes) -> ET.Element:
    prefix_to_uri: dict[str, str] = {}
    for _, item in ET.iterparse(io.BytesIO(data), events=("start-ns",)):
        prefix, uri = item
        prefix = prefix or ""
        prefix_to_uri.setdefault(prefix, uri)
    for prefix, uri in prefix_to_uri.items():
        if prefix in {"xml", "xmlns"}:
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            # Original Office prefixes are valid. Ignore only reserved/internal prefixes.
            pass

    root = ET.fromstring(data)
    used_uris = set()
    for node in root.iter():
        tag_uri = _namespace_uri(node.tag)
        if tag_uri:
            used_uris.add(tag_uri)
        for attribute in node.attrib:
            attr_uri = _namespace_uri(attribute)
            if attr_uri:
                used_uris.add(attr_uri)

    ignorable_key = f"{{{MC}}}Ignorable"
    ignorable = root.get(ignorable_key)
    if ignorable:
        kept = [
            prefix for prefix in ignorable.split()
            if prefix in prefix_to_uri and prefix_to_uri[prefix] in used_uris
        ]
        if kept:
            root.set(ignorable_key, " ".join(kept))
        else:
            root.attrib.pop(ignorable_key, None)
    return root


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
