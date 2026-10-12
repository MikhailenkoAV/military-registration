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
const functions = new Set(["emptyOrganizationProfile", "defaultSettings", "normalizeOrganizationSettings", "makeId", "emptyEmployee", "normalizeEmployee", "parseLocalDate", "formatDate", "familyRows", "maritalStatusKind", "documentMaritalStatus", "f2HeaderAuthority", "documentValues", "assetUrl", "xmlValue", "resizeDocxPlaceholderRun", "buildDocx", "buildChangesDocx", "buildPersonnelListDocx"]);
const constants = new Set(["APP_VERSION"]);
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
const settings = context.defaultSettings();
const realisticSettings = context.defaultSettings();
Object.assign(realisticSettings.profiles.solar_moscow, {
  directorName: "Иванов И.И.",
  responsibleName: "Михайленко А.В.",
  responsiblePhone: "+7 900 000-00-00",
});
const migrated = context.normalizeOrganizationSettings({
  organizationName: "Старое наименование",
  shortName: "Старое краткое",
  organizationAddress: "Старый адрес",
  responsibleName: "Старый ответственный",
  extraHolidays: "2026-09-01",
});
assert.equal(migrated.profiles.solar_sochi.workplaceName, "Старое наименование");
assert.equal(migrated.profiles.solar_sochi.organizationName, "Старое наименование");
assert.equal(migrated.profiles.solar_sochi.organizationAddress, "Старый адрес");
assert.equal(migrated.profiles.solar_sochi.responsibleName, "Старый ответственный");
assert.equal(migrated.profiles.solar_sochi.directorPosition, "Генеральный директор");
assert.match(migrated.profiles.solar_moscow.workplaceName, /город Москва/);
assert.equal(migrated.profiles.airline_solaris.shortName, "АО Авиакомпания «Солярис»");
assert.equal(migrated.extraHolidays, "2026-09-01");
for (const [id, profile] of Object.entries(settings.profiles)) {
  Object.assign(profile, {
    workplaceName: `Место работы ${id}`,
    organizationName: `Полное наименование ${id}`,
    shortName: `Краткое наименование ${id}`,
    organizationAddress: `Адрес ${id}`,
    headerAuthorityType: `Тип органа ${id}`,
    headerAuthorityName: `Наименование органа ${id}`,
    f2BranchTitle: `Логотип ${id}`,
    f2BranchCity: `Город ${id}`,
    responsiblePosition: `Ответственный ${id}`,
    responsibleName: `ФИО ответственного ${id}`,
    responsiblePhone: `Телефон ${id}`,
    directorPosition: `Руководитель ${id}`,
    directorName: `ФИО руководителя ${id}`,
    defaultCommissariat: `Военкомат ${id}`,
  });
}
const parentRows = "Мать | Тестова Елена Викторовна | 1972-07-10\nОтец | Тестов Роман Николаевич | 1971-03-14";
const childRows = "Сын | Тестов Михаил Романович | 2010-02-03\nДочь | Тестова Анна Романовна | 2012-05-06";
const wifeRow = "Жена | Тестова Ольга Петровна | 1980-04-16";
const expectedParents = ["Мать: Тестова Елена Викторовна, 10.07.1972", "Отец: Тестов Роман Николаевич, 14.03.1971"];
function employee(overrides = {}) {
  return context.normalizeEmployee({ fullName: "Тестов Роман Петрович", lastName: "Тестов", firstName: "Роман", middleName: "Петрович", birthDate: "1980-01-02", department: "Лётная служба", position: "пилот", education: "высшее", healthStatus: "здоров", militaryRank: "майор", vus: "616003", registrationAddress: "г. Сочи", ...overrides });
}
function values(person) { return context.documentValues(person, settings, "hire", "", "", "solar_moscow"); }
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
async function generate(outputName, person, organization = settings, profileId = "solar_moscow") {
  const blob = await context.buildDocx(person, organization, "messageSheet", "hire", "", "", profileId);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const xml = strFromU8(unzipSync(bytes)["word/document.xml"]);
  assert.doesNotMatch(xml, /\{\{[A-Z0-9_]+\}\}/);
  assert.equal((xml.match(/М\.П\./g) ?? []).length, 2, "Both M.P. labels must remain in the message sheet");
  const actual = context.documentValues(person, organization, "hire", "", "", profileId);
  for (const key of ["MARITAL_STATUS", "MOTHER_DETAILS", "FATHER_DETAILS", "WIFE_DETAILS", "CHILDREN_DETAILS"]) {
    if (actual[key]) assert.ok(xml.includes(context.xmlValue(actual[key])), `${key} missing from DOCX`);
  }
  fs.mkdirSync(outputDirectory, { recursive: true });
  fs.writeFileSync(path.join(outputDirectory, outputName), bytes);
}
await generate("message-sheet-married.docx", married);
await generate("message-sheet-single.docx", employee({ maritalStatus: "Не состоит в браке (не замужем / не женат)", familyMembers: parentRows }));
await generate("message-sheet-divorced.docx", employee({ maritalStatus: "Разведён (разведена)", familyMembers: parentRows }));
const denseSettings = context.defaultSettings();
denseSettings.profiles.solar_moscow.workplaceName = "Акционерное общество Центр авиации «Солярис» Филиал АО ЦА «Солярис» (город Москва)";
const denseFamilyRows = [
  "Жена | Александрова-Романовская Екатерина Константиновна | 1984-11-19",
  "Сын | Александров-Романовский Константин Владимирович | 2008-01-17",
  "Дочь | Александрова-Романовская Елизавета Владимировна | 2011-06-24",
  "Сын | Александров-Романовский Михаил Владимирович | 2015-09-02",
].join("\n");
await generate("message-sheet-dense.docx", employee({
  fullName: "Александров-Романовский Владимир Константинович",
  lastName: "Александров-Романовский",
  firstName: "Владимир",
  middleName: "Константинович",
  department: "Управление организации лётной деятельности и производственного контроля",
  position: "Ведущий инженер по техническому обслуживанию авиационной техники",
  education: "высшее профессиональное, магистратура и профессиональная переподготовка",
  healthStatus: "годен к военной службе с незначительными ограничениями",
  maritalStatus: "Состоит в зарегистрированном браке (женат / замужем)",
  familyMembers: denseFamilyRows,
  registrationAddress: "Российская Федерация, Московская область, городской округ Воскресенск, улица Центральная, дом 125, корпус 3, квартира 147",
}), denseSettings);

const organizationOutput = path.join(root, "work/rendered-organization-profiles");
fs.mkdirSync(organizationOutput, { recursive: true });
const profileEmployee = employee({
  id: "profile-test-employee",
  birthPlace: "г. Сочи",
  passportSeries: "0312",
  passportNumber: "123456",
  snils: "123-456-789 00",
  composition: "Солдаты",
  reserveCategory: "1",
  fitnessCategory: "А",
  accountType: "general",
  orderNumber: "15-к",
  orderDate: "2026-08-01",
  militaryDocIssuedBy: "Военный комиссариат города Москвы",
  militaryCommissariat: "Военный комиссариат г. Воскресенска Московской обл.",
});
async function saveAndInspect(blob, filename, expected) {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const archive = unzipSync(bytes);
  const xml = Object.entries(archive)
    .filter(([name]) => name.startsWith("word/") && name.endsWith(".xml"))
    .map(([, content]) => strFromU8(content)).join("\n");
  assert.doesNotMatch(xml, /\{\{[A-Z0-9_]+\}\}/);
  for (const value of expected) assert.ok(xml.includes(context.xmlValue(value)), `${filename}: missing ${value}`);
  fs.writeFileSync(path.join(organizationOutput, filename), bytes);
}
const longHeaderEmployee = {
  ...profileEmployee,
  militaryCommissariat: "Военный комиссариат Московской области по городам Шатура, Егорьевск, Рошаль, Шатурскому и Егорьевскому районам",
};
await saveAndInspect(
  await context.buildDocx(longHeaderEmployee, realisticSettings, "f2", "hire", "15-к", "2026-08-01", "solar_moscow"),
  "f2-header-long-realistic.docx",
  ["комиссариата Московской области по городам Шатура, Егорьевск, Рошаль, Шатурскому и Егорьевскому районам"],
);
for (const profileId of ["solar_sochi", "solar_moscow", "airline_solaris"]) {
  const profile = settings.profiles[profileId];
  const actual = context.documentValues(profileEmployee, settings, "hire", "15-к", "2026-08-01", profileId);
  assert.equal(actual.ORG_NAME, profile.workplaceName);
  assert.equal(actual.ORG_SHORT_NAME, profile.shortName);
  assert.equal(actual.ORG_ADDRESS, profile.organizationAddress);
  assert.equal(actual.HEADER_AUTHORITY_TYPE, profile.headerAuthorityType);
  assert.equal(actual.HEADER_AUTHORITY_NAME, profile.headerAuthorityName);
  assert.equal(actual.F2_BRANCH_TITLE, profile.f2BranchTitle);
  assert.equal(actual.F2_BRANCH_CITY, profile.f2BranchCity);
  assert.equal(actual.RESPONSIBLE_NAME, profile.responsibleName);
  const f2Actual = context.documentValues(profileEmployee, settings, "hire", "15-к", "2026-08-01", profileId, "f2");
  assert.equal(f2Actual.HEADER_AUTHORITY_TYPE, "военного");
  assert.equal(f2Actual.HEADER_AUTHORITY_NAME, "комиссариата г. Воскресенска Московской обл.");
  for (const type of ["form10", "f2", "messageSheet", "employmentNotice"]) {
    const expectedName = type === "form10" || type === "messageSheet" ? profile.workplaceName : profile.organizationName;
    const expected = [expectedName];
    if (type === "f2") expected.push(profile.shortName, profile.f2BranchTitle, profile.f2BranchCity);
    if (type === "f2") expected.push(profile.organizationAddress, f2Actual.HEADER_AUTHORITY_NAME);
    if (type === "employmentNotice") expected.push(profile.organizationAddress, profile.headerAuthorityName);
    await saveAndInspect(
      await context.buildDocx(profileEmployee, settings, type, "hire", "15-к", "2026-08-01", profileId),
      `${profileId}-${type}.docx`,
      expected,
    );
  }
  await saveAndInspect(
    await context.buildChangesDocx([{ employeeId: profileEmployee.id, content: "Изменение сведений" }], [profileEmployee], settings, profileId),
    `${profileId}-changes.docx`,
    [profile.organizationName, profile.organizationAddress, profile.headerAuthorityName],
  );
  for (const type of ["officerList", "enlistedList"]) {
    await saveAndInspect(
      await context.buildPersonnelListDocx(type, [profileEmployee.id], [profileEmployee], settings, profileId),
      `${profileId}-${type}.docx`,
      [profile.organizationName, profile.organizationAddress, profile.headerAuthorityName],
    );
  }
}
const abbreviatedF2 = context.documentValues(
  employee({ militaryDocIssuedBy: "Военный комиссариат города Москвы", militaryCommissariat: "ВК Амурской обл. по г. Благовещенск и Благовещенскому р-ну" }),
  settings,
  "hire",
  "15-к",
  "2026-08-01",
  "solar_sochi",
  "f2",
);
assert.equal(abbreviatedF2.HEADER_AUTHORITY_TYPE, "военного");
assert.equal(abbreviatedF2.HEADER_AUTHORITY_NAME, "комиссариата Амурской обл. по г. Благовещенск и Благовещенскому р-ну");
const shortF2Employee = employee({
  militaryCommissariat: "Военный комиссариат города Москвы",
  orderNumber: "15-к",
  orderDate: "2026-08-01",
});
await saveAndInspect(
  await context.buildDocx(shortF2Employee, settings, "f2", "hire", "15-к", "2026-08-01", "solar_moscow"),
  "f2-header-short.docx",
  ["военного", "комиссариата города Москвы"],
);
const abbreviatedF2Employee = employee({
  militaryDocIssuedBy: "Военный комиссариат города Москвы",
  militaryCommissariat: "ВК Амурской обл. по г. Благовещенск и Благовещенскому р-ну",
  orderNumber: "15-к",
  orderDate: "2026-08-01",
});
await saveAndInspect(
  await context.buildDocx(abbreviatedF2Employee, settings, "f2", "hire", "15-к", "2026-08-01", "solar_sochi"),
  "f2-header-abbreviated.docx",
  [abbreviatedF2.HEADER_AUTHORITY_NAME],
);
await saveAndInspect(
  await context.buildDocx(abbreviatedF2Employee, realisticSettings, "f2", "hire", "15-к", "2026-08-01", "solar_moscow"),
  "f2-header-abbreviated-realistic.docx",
  [abbreviatedF2.HEADER_AUTHORITY_NAME],
);
const ignoresIssuerF2 = context.documentValues(
  employee({ militaryDocIssuedBy: "Военный комиссариат г. Воскресенска Московской обл.", militaryCommissariat: "Военный комиссариат города Москвы" }),
  settings,
  "hire",
  "15-к",
  "2026-08-01",
  "solar_moscow",
  "f2",
);
assert.equal(ignoresIssuerF2.HEADER_AUTHORITY_NAME, "комиссариата города Москвы");
// Optional local-only verification: no backup records are copied into source or fixtures.
if (process.argv[2]) {
  const backup = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
  const matches = backup.employees.filter((item) => item.lastName === process.argv[3]);
  assert.equal(matches.length, 1, "Expected one employee in the supplied backup");
  const restoredSettings = context.normalizeOrganizationSettings(backup.settings);
  const person = context.normalizeEmployee(matches[0]);
  if (process.argv[4] === "f2") {
    const actual = context.documentValues(person, restoredSettings, "hire", person.orderNumber, person.orderDate, "solar_moscow", "f2");
    const expected = context.f2HeaderAuthority(person.militaryCommissariat);
    assert.equal(actual.HEADER_AUTHORITY_NAME, expected.name);
    await saveAndInspect(
      await context.buildDocx(person, restoredSettings, "f2", "hire", person.orderNumber, person.orderDate, "solar_moscow"),
      `f2-backup-${person.lastName}.docx`,
      [expected.name],
    );
  } else {
    for (const profileId of ["solar_moscow", "solar_sochi", "airline_solaris"]) {
      const actual = context.documentValues(person, restoredSettings, "hire", "", "", profileId);
      for (const relation of ["Мать", "Отец"]) {
        const parent = context.familyRows(person.familyMembers).find((item) => item.relation === relation);
        assert.ok(parent, `Missing ${relation} in backup`);
        assert.equal(actual[relation === "Мать" ? "MOTHER_DETAILS" : "FATHER_DETAILS"], `${relation}: ${parent.fullName}, ${context.formatDate(parent.birthDate, "")}`);
      }
      await generate(`message-sheet-backup-${profileId}.docx`, person, restoredSettings, profileId);
    }
  }
}
console.log("PASS: family selection regression cases and application DOCX generation");
