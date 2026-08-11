from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

REPLACEMENTS = {
    "ФИЛИАЛ АО ЦА «СОЛЯРИС»": "{{F2_BRANCH_TITLE}}",
    "(город Москва)": "{{F2_BRANCH_CITY}}",
}


def replace_paragraph_text(root: etree._Element) -> None:
    found: set[str] = set()
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        nodes = paragraph.xpath(".//w:t", namespaces=NS)
        combined = "".join(node.text or "" for node in nodes).strip()
        replacement = REPLACEMENTS.get(combined)
        if replacement is None or not nodes:
            continue
        nodes[0].text = replacement
        nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for node in nodes[1:]:
            node.text = ""
        found.add(combined)
    missing = set(REPLACEMENTS) - found
    if missing:
        raise RuntimeError(f"Не найдены строки шаблона: {sorted(missing)}")

    # Длинный обязательный московский адрес занимает две строки. Уменьшаем
    # только избыточный отступ перед заголовком, сохраняя шрифты и сетку формы.
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()
        if text != "СВЕДЕНИЯ":
            continue
        properties = paragraph.find(f"{{{W}}}pPr")
        if properties is None:
            properties = etree.Element(f"{{{W}}}pPr")
            paragraph.insert(0, properties)
        spacing = properties.find(f"{{{W}}}spacing")
        if spacing is None:
            spacing = etree.SubElement(properties, f"{{{W}}}spacing")
        spacing.set(f"{{{W}}}before", "360")


def prepare(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                root = etree.fromstring(data)
                replace_paragraph_text(root)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            dst.writestr(info, data)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: prepare_f2_template_v2.py input.docx output.docx")
    prepare(Path(sys.argv[1]), Path(sys.argv[2]))
