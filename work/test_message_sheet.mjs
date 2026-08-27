import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { webcrypto } from "node:crypto";
import ts from "typescript";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = path.resolve(import.meta.dirname, "..");
const outputDirectory = path.join(root, "work/rendered-message-sheet");
// Exercise the actual application functions, not a copy of the selection logic.
const source = fs.readFileSync(path.join(root, "app/page.tsx"), "utf8");
const ast = ts.createSourceFile("page.tsx", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
const functions = new Set(["makeId", "emptyEmployee", "normalizeEmployee", "parseLocalDate", "formatDate", "familyRows", "maritalStatusKind", "documentMaritalStatus", "documentValues", "assetUrl", "xmlValue", "buildDocx"]);
const constants = new Set(["APP_VERSION", "MOSCOW_WORKPLACE"]);
const extracted = ast.statements.filter((node) =>
  (ts.isFunctionDeclaration(node) && functions.has(node.name?.text)) ||
  (ts.isVariableStatement(node) && node.declarationList.declarations.some((item) => constants.has(item.name.getText(ast)))),
).map((node) => node.getText(ast)).join("\n");
const context = vm.createContext({
  Blob, Uint8Array, Date, Intl, crypto: webcrypto, strFromU8, strToU8, unzipSync, zipSync,
  process: { env: { NEXT_PUBLIC_BASE_PATH: "/military-registration" } },
  fetch: async (url, options) => {
    assert.equal(options.cache, "no-store");
    const parsed = new URL(url, "https://example.test");
    assert.ok(parsed.searchParams.has("v"));
    const bytes = fs.readFileSync(path.join(root, "public/documents", path.basename(parsed.pathname)));
    return { ok: true, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) };
  },
});
vm.runInContext(ts.transpileModule(extracted, { compilerOptions: { target: ts.ScriptTarget.ES2022 } }).outputText, context);
const settings = { organizationName: "Тестовая организация", shortName: "Тест", defaultCommissariat: "" };
const parentRows = "Мать | Тестова Елена Викторовна | 1972-07-10\nОтец | Тестов Роман Николаевич | 1971-03-14";
const childRows = "Сын | Тестов Михаил Романович | 2010-02-03\nДочь | Тестова Анна Романовна | 2012-05-06";
const wifeRow = "Жена | Тестова Ольга Петровна | 1980-04-16";
const expectedParents = ["Мать: Тестова Елена Викторовна, 10.07.1972", "Отец: Тестов Роман Николаевич, 14.03.1971"];
function employee(overrides = {}) {
  return context.normalizeEmployee({ fullName: "Тестов Роман Петрович", lastName: "Тестов", firstName: "Роман", middleName: "Петрович", birthDate: "1980-01-02", department: "Лётная служба", position: "пилот", education: "высшее", healthStatus: "здоров", militaryRank: "майор", vus: "616003", registrationAddress: "г. Сочи", ...overrides }, settings);
}
function values(person) { return context.documentValues(person, settings, "hire", "", "", "moscow"); }
function assertParents(actual, expected = expectedParents) {
  assert.deepEqual([actual.MOTHER_DETAILS, actual.FATHER_DETAILS], expected);
}
for (const status of ["Разведён (разведена)", "Разведен", "Вдовец (вдова)", "Вдова"]) {
  for (const sex of ["male", "female"]) {
    const actual = values(employee({ maritalStatus: status, sex, familyMembers: parentRows }));
    assertParents(actual);
    assert.equal(actual.MARITAL_STATUS, status);
    assert.equal(actual.WIFE_DETAILS, "");
    assert.equal(actual.CHILDREN_DETAILS, "");
    const withChildren = values(employee({ maritalStatus: status, sex, familyMembers: `${parentRows}\n${childRows}` }));
    assertParents(withChildren, ["", ""]);
    assert.match(withChildren.CHILDREN_DETAILS, /^Сын: .*03\.02\.2010; Дочь: .*06\.05\.2012$/);
  }
}
for (const sex of ["male", "female"]) {
  const actual = values(employee({ maritalStatus: "Не состоит в браке (не замужем / не женат)", sex, familyMembers: parentRows }));
  assertParents(actual);
  assert.equal(actual.MARITAL_STATUS, sex === "male" ? "холост" : "не замужем");
  assert.equal(actual.WIFE_DETAILS, "");
}
const married = employee({ maritalStatus: "Состоит в зарегистрированном браке (женат / замужем)", familyMembers: `${wifeRow}\n${childRows}\n${parentRows}` });
assertParents(values(married), ["", ""]);
assert.equal(values(married).WIFE_DETAILS, "Жена: Тестова Ольга Петровна, 16.04.1980");
assert.match(values(married).CHILDREN_DETAILS, /Сын: .*; Дочь: /);
for (const familyMembers of ["", "Мать | Тестова Елена Викторовна |", "Отец | Тестов Роман Николаевич | 1971-03-14"]) {
  assertParents(values(employee({ maritalStatus: "Разведён (разведена)", familyMembers })), [familyMembers.startsWith("Мать") ? "Мать: Тестова Елена Викторовна" : "", familyMembers.startsWith("Отец") ? expectedParents[1] : ""]);
}
async function generate(outputName, person, organization = settings, location = "moscow") {
  const blob = await context.buildDocx(person, organization, "messageSheet", "hire", "", "", location);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]);
  assert.doesNotMatch(xml, /\{\{[A-Z0-9_]+\}\}/);
  const actual = context.documentValues(person, organization, "hire", "", "", location);
  for (const key of ["MARITAL_STATUS", "MOTHER_DETAILS", "FATHER_DETAILS", "WIFE_DETAILS", "CHILDREN_DETAILS"]) {
    if (actual[key]) assert.ok(xml.includes(context.xmlValue(actual[key])), `${key} missing from DOCX`);
  }
  fs.mkdirSync(outputDirectory, { recursive: true });
  fs.writeFileSync(path.join(outputDirectory, outputName), bytes);
}
await generate("message-sheet-married.docx", married);
await generate("message-sheet-single.docx", employee({ maritalStatus: "Не состоит в браке (не замужем / не женат)", familyMembers: parentRows }));
await generate("message-sheet-divorced.docx", employee({ maritalStatus: "Разведён (разведена)", familyMembers: parentRows }));
// Optional local-only verification: no backup records are copied into source or fixtures.
if (process.argv[2]) {
  const backup = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
  const matches = backup.employees.filter((item) => item.lastName === process.argv[3]);
  assert.equal(matches.length, 1, "Expected one employee in the supplied backup");
  const person = context.normalizeEmployee(matches[0], backup.settings);
  for (const location of ["moscow", "sochi"]) {
    const actual = context.documentValues(person, backup.settings, "hire", "", "", location);
    for (const relation of ["Мать", "Отец"]) {
      const parent = context.familyRows(person.familyMembers).find((item) => item.relation === relation);
      assert.ok(parent, `Missing ${relation} in backup`);
      assert.equal(actual[relation === "Мать" ? "MOTHER_DETAILS" : "FATHER_DETAILS"], `${relation}: ${parent.fullName}, ${context.formatDate(parent.birthDate, "")}`);
    }
    await generate(`message-sheet-backup-${location}.docx`, person, backup.settings, location);
  }
}
console.log("PASS: family selection regression cases and application DOCX generation");
