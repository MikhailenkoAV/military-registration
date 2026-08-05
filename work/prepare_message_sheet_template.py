from copy import copy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


SOURCE = Path("/workspace/scratch/b9fd23b3471c/upload/Листок сообщений.docx")
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
            if len(runs) != 27:
                raise RuntimeError(f"Expected 27 red runs, found {len(runs)}")

            replacements = {
                0: "{{LAST_NAME}}",
                1: "{{FIRST_NAME}}",
                2: "{{MIDDLE_NAME}}",
                3: "{{BIRTH_YEAR}}",
                4: "",
                5: "",
                6: "{{MILITARY_RANK}}",
                7: "{{VUS}}",
                8: "{{EDUCATION}}",
                9: "{{POSITION}}",
                10: "{{MARITAL_STATUS}}",
                11: ")",
                12: " {{REGISTRATION_ADDRESS}}",
                13: "",
                14: "",
                15: "",
                16: "",
                17: "",
                18: "{{LAST_NAME}}",
                19: "{{FIRST_NAME}}",
                20: "",
                21: "{{MIDDLE_NAME}}",
                22: "",
                23: "",
                24: "",
                25: "{{DEPARTMENT}}",
                26: "",
            }

            for index, run in enumerate(runs):
                texts = run.findall("w:t", NS)
                for text in texts:
                    run.remove(text)
                text = etree.SubElement(run, f"{{{W}}}t")
                text.text = replacements[index]
                if text.text and (text.text.startswith(" ") or text.text.endswith(" ")):
                    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                color = run.find("w:rPr/w:color", NS)
                if color is not None:
                    color.set(f"{{{W}}}val", "000000")

            content = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        target.writestr(copy(item), content)

TEMPORARY.replace(TARGET)
