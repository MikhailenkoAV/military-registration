from __future__ import annotations

import copy
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
TOKEN = re.compile(r"\{\{[A-Z0-9_]+\}\}")


def normalize_xml(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
        if not text_nodes:
            continue
        combined = "".join(node.text or "" for node in text_nodes)
        matches = list(TOKEN.finditer(combined))
        for match in reversed(matches):
            positions: list[tuple[etree._Element, int, int]] = []
            cursor = 0
            for node in text_nodes:
                value = node.text or ""
                positions.append((node, cursor, cursor + len(value)))
                cursor += len(value)
            covered = [(node, start, end) for node, start, end in positions if end > match.start() and start < match.end()]
            if len(covered) <= 1:
                continue
            first, first_start, _ = covered[0]
            last, last_start, _ = covered[-1]
            before = (first.text or "")[: match.start() - first_start]
            after = (last.text or "")[match.end() - last_start :]

            # Copy formatting from the run containing most of the token name.
            style_node = max(covered, key=lambda item: len((item[0].text or "").strip("{}")))[0]
            first_run = first.getparent()
            style_run = style_node.getparent()
            style_properties = style_run.find(f"{{{W}}}rPr")
            current_properties = first_run.find(f"{{{W}}}rPr")
            if current_properties is not None:
                first_run.remove(current_properties)
            if style_properties is not None:
                first_run.insert(0, copy.deepcopy(style_properties))

            first.text = before + match.group() + after
            first.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for node, _, _ in covered[1:]:
                node.text = ""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                data = normalize_xml(data)
            dst.writestr(info, data)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: normalize_message_template.py input.docx output.docx")
    main(Path(sys.argv[1]), Path(sys.argv[2]))
