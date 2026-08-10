import fs from "node:fs";
import path from "node:path";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = path.resolve(import.meta.dirname, "..");
const source = path.join(root, "public/documents/message-sheet-template.docx");
const outputDirectory = path.join(root, "work/rendered-message-sheet");
const target = path.join(outputDirectory, "message-sheet-filled.docx");
fs.mkdirSync(outputDirectory, { recursive: true });

const values = {
  ORG_NAME: "Акционерное общество Центр авиации «Солярис»",
  LAST_NAME: "Пронин",
  FIRST_NAME: "Александр",
  MIDDLE_NAME: "Константинович",
  BIRTH_YEAR: "1968",
  MILITARY_RANK: "майор",
  VUS: "616003",
  EDUCATION: "высшее",
  POSITION: "пилот-инструктор",
  HEALTH_STATUS: "здоров",
  MARITAL_STATUS: "женат",
  WIFE_DETAILS: "Пронина Елена Сергеевна, 14.03.1971",
  CHILDREN_DETAILS: "Сын: Пронин Михаил Александрович, 22.07.1998; Дочь: Пронина Анна Александровна, 05.11.2003",
  MOTHER_DETAILS: "",
  FATHER_DETAILS: "",
  REGISTRATION_ADDRESS: "г. Сочи, ул. Бамбуковая, д. 42, кв. 47",
  DEPARTMENT: "Лётная служба",
};

function xmlValue(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

const archive = unzipSync(new Uint8Array(fs.readFileSync(source)));
for (const [entry, content] of Object.entries(archive)) {
  if (!entry.startsWith("word/") || !entry.endsWith(".xml")) continue;
  let xml = strFromU8(content);
  for (const [key, value] of Object.entries(values)) {
    xml = xml.replaceAll(`{{${key}}}`, xmlValue(value));
  }
  xml = xml.replace(/\{\{[A-Z0-9_]+\}\}/g, "");
  archive[entry] = strToU8(xml);
}
fs.writeFileSync(target, zipSync(archive, { level: 6 }));
