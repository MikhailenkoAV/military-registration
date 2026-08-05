from copy import copy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


SOURCE = Path("/workspace/scratch/b9fd23b3471c/upload/Документ Microsoft Word(1).docx")
TARGET = Path(__file__).parents[1] / "public" / "documents" / "message-sheet-template.docx"
TEMPORARY = TARGET.with_suffix(".tmp.docx")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def red_runs(root):
    result = []
    for run in root.xpath(".//w:r", namespaces=NS):
        color = run.find("w:rPr/w:color", NS)
        if color is not None and color.get(f"{{{W}}}val") == "EE0000":
            result.append(run)
    return result


with ZipFile(SOURCE, "r") as source, ZipFile(TEMPORARY, "w", ZIP_DEFLATED) as target:
    for item in source.infolist():
        content = source.read(item.filename)
        if item.filename == "word/document.xml":
            root = etree.fromstring(content)
            runs = red_runs(root)
            if len(runs) != 20:
                raise RuntimeError(f"Expected 20 red runs, found {len(runs)}")

            replacements = {
                0: "{{LAST_NAME}}",
                1: "{{FIRST_NAME}}",
                2: "",
                3: "{{LAST_NAME}}",
                4: "{{MIDDLE_NAME}}",
                5: "",
                6: "{{FIRST_NAME}}",
                7: "{{MIDDLE_NAME}}",
                8: "{{BIRTH_YEAR}}",
                9: "{{MILITARY_RANK}}",
                10: "{{VUS}}",
                11: "{{EDUCATION}}",
                12: "{{ORG_NAME}}",
                13: "",
                14: "{{POSITION}}",
                15: "{{HEALTH_STATUS}}",
                16: "{{DEPARTMENT}}",
                17: "",
                18: "{{MARITAL_STATUS}}",
                19: "{{REGISTRATION_ADDRESS}}",
            }

            for index, run in enumerate(runs):
                for text in run.findall("w:t", NS):
                    run.remove(text)
                text = etree.SubElement(run, f"{{{W}}}t")
                text.text = replacements[index]
                color = run.find("w:rPr/w:color", NS)
                if color is not None:
                    color.set(f"{{{W}}}val", "000000")

            # Уплотняем только масштаб шрифта и поля страницы: структура и
            # размеры созданных пользователем ячеек остаются без изменений.
            for size in root.xpath(".//w:sz | .//w:szCs", namespaces=NS):
                value = size.get(f"{{{W}}}val")
                if value and value.isdigit():
                    size.set(f"{{{W}}}val", str(max(14, round(int(value) * 0.88))))

            section = root.find(".//w:sectPr", NS)
            if section is not None:
                margins = section.find("w:pgMar", NS)
                if margins is not None:
                    margins.set(f"{{{W}}}top", "400")
                    margins.set(f"{{{W}}}bottom", "400")

            # В исходнике после таблицы оставлены три пустых абзаца. При
            # плотной табличной верстке Word переносит их на пустую страницу.
            body = root.find("w:body", NS)
            if body is not None:
                for child in list(body):
                    if child.tag == f"{{{W}}}p" and not child.xpath(".//w:t[normalize-space()]", namespaces=NS):
                        body.remove(child)

            content = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        target.writestr(copy(item), content)

TEMPORARY.replace(TARGET)
