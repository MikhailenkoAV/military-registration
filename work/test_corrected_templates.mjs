import fs from "node:fs";
import path from "node:path";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = path.resolve(import.meta.dirname, "..");
const output = path.join(root, "work/rendered-corrections");
fs.mkdirSync(output, { recursive: true });

const values = {
  ORG_NAME: "Акционерное общество Центр авиации «Солярис»",
  ORG_SHORT_NAME: "АО ЦА «Солярис»",
  ORG_ADDRESS: "г. Москва",
  LAST_NAME: "Пилипчук",
  FIRST_NAME: "Владимир",
  MIDDLE_NAME: "Анатольевич",
  FULL_NAME: "Пилипчук Владимир Анатольевич",
  EMPLOYEE_NAME: "Пилипчук Владимир Анатольевич",
  BIRTH_DATE: "16.12.1993",
  BIRTH_PLACE: "г. Москва",
  EDUCATION: "высшее",
  PROFESSION: "пилот",
  MARITAL_STATUS: "женат",
  PASSPORT_SERIES: "4510",
  PASSPORT_NUMBER: "123456",
  PASSPORT: "4510 123456",
  PASSPORT_ISSUE_DATE: "10.10.2010",
  PASSPORT_ISSUED_BY: "МВД России",
  REGISTRATION_ADDRESS: "г. Москва",
  MOBILE_PHONE: "+7 900 000-00-00",
  RESERVE_CATEGORY: "1",
  COMMISSARIAT: "Жуковский горвоенкомат Московской обл.",
  COMMISSARIAT_SHORT: "Жуковский горвоенкомат Московской обл.",
  MILITARY_RANK: "Матрос",
  COMPOSITION_PROFILE: "Солдаты",
  VUS: "104916",
  FITNESS_CATEGORY: "Б-3",
  MILITARY_DOCUMENT: "Военный билет, АС №4245960",
  ORDER_DETAILS: "01.08.2026 № 15-к",
  POSITION: "пилот",
  RESPONSIBLE_POSITION: "Специалист по ведению воинского учёта",
  RESPONSIBLE_NAME: "Михайленко А.В.",
  SNILS: "123-456-789 00",
  EVENT_HIRE: "принят (поступил) на работу",
  EVENT_DISMISSAL: "уволен с работы (отчислен из образовательной организации)",
  HIRE_STRIKE: "0",
  DISMISSAL_STRIKE: "1",
  ORDER_NUMBER: "15-к",
  ORDER_DAY: "01",
  ORDER_MONTH: "августа",
  ORDER_YEAR: "26",
  DIRECTOR_POSITION: "Генеральный директор",
  DIRECTOR_NAME: "Петров П.П.",
  HEADER_AUTHORITY_TYPE: "военного",
  HEADER_AUTHORITY_NAME: "комиссариата города Москвы",
  F2_BRANCH_TITLE: "ФИЛИАЛ АО ЦА «СОЛЯРИС»",
  F2_BRANCH_CITY: "(город Москва)",
};

function xmlValue(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function generate(templateName, outputName, overrides = {}) {
  const source = path.join(root, "public/documents", templateName);
  const archive = unzipSync(new Uint8Array(fs.readFileSync(source)));
  const documentValues = { ...values, ...overrides };
  for (const [entry, content] of Object.entries(archive)) {
    if (!entry.startsWith("word/") || !entry.endsWith(".xml")) continue;
    let xml = strFromU8(content);
    for (const [key, value] of Object.entries(documentValues)) {
      xml = xml.replaceAll(`{{${key}}}`, xmlValue(value));
    }
    xml = xml.replace(/\{\{[A-Z0-9_]+\}\}/g, "");
    archive[entry] = strToU8(xml);
  }
  fs.writeFileSync(path.join(output, outputName), zipSync(archive, { level: 6 }));
}

generate("form10-template.docx", "form10-corrected.docx");
generate("f2-template.docx", "f2-moscow.docx", {
  REGISTRATION_ADDRESS: "г. Москва, ул. Заморенова, д. 12, стр. 1",
  ORG_ADDRESS: "Российская Федерация, город Москва, вн. тер. г. муниципальный округ Пресненский, наб. Пресненская, д. 12, помещ. 20/80",
  ORG_NAME: 'Акционерное общество Центр авиации "Солярис"',
  ORG_SHORT_NAME: 'Филиал АО ЦА "Солярис" (город Москва)',
});
generate("f2-template.docx", "f2-sochi.docx", {
  HEADER_AUTHORITY_TYPE: "городского округа",
  HEADER_AUTHORITY_NAME: "город-курорт Сочи Краснодарского края",
  F2_BRANCH_TITLE: "",
  F2_BRANCH_CITY: "",
  REGISTRATION_ADDRESS: "г. Сочи, ул. Бамбуковая, д. 42, кв. 47",
  ORG_ADDRESS: "Краснодарский край, город Сочи, ул. Авиационная, д. 1",
  ORG_NAME: "Акционерное общество Центр авиации «Солярис»",
  ORG_SHORT_NAME: "АО ЦА «Солярис»",
});
