"use client";

import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import * as XLSX from "xlsx";
import { LIST_TEMPLATE_BASE64 } from "./generated-list-templates";
import {
  AlertCircle,
  ArchiveRestore,
  Bell,
  BookOpen,
  CalendarCheck,
  Check,
  ChevronDown,
  ChevronRight,
  CircleAlert,
  Clock3,
  Database,
  Download,
  FileDown,
  FileSpreadsheet,
  FileText,
  Home,
  Info,
  ListChecks,
  Menu,
  Pencil,
  Plus,
  RotateCcw,
  Search,
  Settings,
  ShieldCheck,
  Trash2,
  Upload,
  UserRound,
  Users,
  X,
} from "lucide-react";

type View =
  | "dashboard"
  | "employees"
  | "notifications"
  | "workPlan"
  | "documents"
  | "reconciliations"
  | "settings";

type DocumentType = "form10" | "f2" | "messageSheet" | "changes" | "employmentNotice" | "officerList" | "enlistedList";
type DocumentHeaderLocation = "moscow" | "sochi";

type ChangeDocumentEntry = {
  employeeId: string;
  content: string;
};

type Employee = {
  id: string;
  fullName: string;
  lastName: string;
  firstName: string;
  middleName: string;
  sex: "male" | "female";
  active: boolean;
  department: string;
  position: string;
  hireDate: string;
  dismissalDate: string;
  orderNumber: string;
  orderDate: string;
  birthDate: string;
  birthPlace: string;
  passportSeries: string;
  passportNumber: string;
  passportIssueDate: string;
  passportIssuedBy: string;
  registrationAddress: string;
  registrationDate: string;
  actualAddress: string;
  actualAddressDate: string;
  phone: string;
  workPhone: string;
  snils: string;
  inn: string;
  education: string;
  profession: string;
  languages: string;
  driverLicense: string;
  maritalStatus: string;
  familyMembers: string;
  militaryDocType: string;
  militaryDocNumber: string;
  militaryDocIssueDate: string;
  militaryDocIssuedBy: string;
  militaryRank: string;
  composition: string;
  profile: string;
  vus: string;
  reserveCategory: string;
  fitnessCategory: string;
  healthStatus: string;
  militaryCommissariat: string;
  militaryCommissariatAddress: string;
  accountType: "general" | "special" | "";
  teamNumber: string;
  specialAccountNumber: string;
  lastEmployeeVerification: string;
  lastCommissariatVerification: string;
  notes: string;
};

type Notice = {
  id: string;
  employeeId: string;
  ruleId: string;
  eventDate: string;
  dueDate: string;
  createdAt: string;
  completedAt: string;
  outgoingNumber: string;
  note: string;
  completedByDocumentId?: string;
};

type DocumentRecord = {
  id: string;
  employeeId: string;
  type: DocumentType;
  createdAt: string;
  title: string;
  status?: "formed" | "signed" | "sent";
  headerLocation?: DocumentHeaderLocation;
  eventType?: "hire" | "dismissal";
  orderNumber?: string;
  orderDate?: string;
  employeeIds?: string[];
  changeEntries?: ChangeDocumentEntry[];
  employeeCount?: number;
  outgoingNumber?: string;
  sentAt?: string;
  noticeIds?: string[];
};

type EmployeeFieldChange = {
  key: keyof Employee;
  label: string;
  oldValue: string | boolean;
  newValue: string | boolean;
};

type EmployeeChangeRecord = {
  id: string;
  employeeId: string;
  employeeName: string;
  createdAt: string;
  changes: EmployeeFieldChange[];
  undoneAt?: string;
};

type OrganizationSettings = {
  organizationName: string;
  shortName: string;
  organizationAddress: string;
  directorPosition: string;
  directorName: string;
  responsiblePosition: string;
  responsibleName: string;
  responsiblePhone: string;
  defaultCommissariat: string;
  defaultCommissariatAddress: string;
  extraHolidays: string;
};

type StoredState = {
  employees: Employee[];
  notices: Notice[];
  documents: DocumentRecord[];
  employeeChanges: EmployeeChangeRecord[];
  settings: OrganizationSettings;
};

type Rule = {
  id: string;
  title: string;
  shortTitle: string;
  days: number | null;
  workingDays: boolean;
  source: string;
  sourceUrl: string;
  help: string;
  documentHint: string;
};

const STORAGE_KEY = "voinskiy-uchet-v1";
const LAST_BACKUP_KEY = "voinskiy-uchet-last-backup";
const APP_VERSION = "19.8";

const RULES: Rule[] = [
  {
    id: "hire",
    title: "Направить сведения о приёме",
    shortTitle: "Приём сотрудника",
    days: 5,
    workingDays: false,
    source: "Постановление Правительства РФ № 719, п. 32 «а»",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_64215/674f7a847f97b0d98c437823abac7703aa4f5dee/",
    help: "Сведения направляются в соответствующий военный комиссариат в течение 5 дней со дня приёма.",
    documentHint: "Справка Ф-2 — сведения о приёме",
  },
  {
    id: "dismissal",
    title: "Направить сведения об увольнении",
    shortTitle: "Увольнение сотрудника",
    days: 5,
    workingDays: false,
    source: "Постановление Правительства РФ № 719, п. 32 «а»",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_64215/674f7a847f97b0d98c437823abac7703aa4f5dee/",
    help: "Сведения направляются в соответствующий военный комиссариат в течение 5 дней со дня увольнения.",
    documentHint: "Справка Ф-2 — сведения об увольнении",
  },
  {
    id: "change",
    title: "Сообщить об изменении учётных сведений",
    shortTitle: "Изменение сведений",
    days: 5,
    workingDays: false,
    source: "Постановление Правительства РФ № 719, п. 32 «е»",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_64215/674f7a847f97b0d98c437823abac7703aa4f5dee/",
    help:
      "Учитываются изменения семейного положения, образования, подразделения, должности, адреса и состояния здоровья.",
    documentHint: "Сообщение об изменении сведений",
  },
  {
    id: "unregistered",
    title: "Сообщить о гражданине, не состоящем на учёте",
    shortTitle: "Не состоит на учёте",
    days: 3,
    workingDays: true,
    source: "Федеральный закон № 53-ФЗ, ст. 4",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_18260/600d790053e43f565b39b6391205e2b91a2f8318/",
    help:
      "Сведения о выявленном гражданине, который обязан состоять на воинском учёте, направляются в течение 3 рабочих дней.",
    documentHint: "Сообщение и направление для постановки на учёт",
  },
  {
    id: "document_issue",
    title: "Сообщить об ошибках в документах",
    shortTitle: "Ошибка в документе",
    days: 5,
    workingDays: true,
    source: "Постановление Правительства РФ № 719, п. 30 «г»",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_64215/674f7a847f97b0d98c437823abac7703aa4f5dee/",
    help:
      "Сообщаются неоговорённые исправления, неточности, признаки подделки, неполные листы и нарушения обязанностей.",
    documentHint: "Информационное письмо в военный комиссариат",
  },
  {
    id: "request",
    title: "Ответить на запрос военного комиссариата",
    shortTitle: "Запрос военкомата",
    days: 14,
    workingDays: false,
    source: "Постановление Правительства РФ № 719, п. 32 «б»",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_64215/674f7a847f97b0d98c437823abac7703aa4f5dee/",
    help: "Необходимые сведения по запросу направляются в двухнедельный срок.",
    documentHint: "Ответ на запрос военного комиссариата",
  },
  {
    id: "summons",
    title: "Оповестить сотрудника о повестке",
    shortTitle: "Получена повестка",
    days: null,
    workingDays: false,
    source: "Федеральный закон № 53-ФЗ, ст. 4",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_18260/600d790053e43f565b39b6391205e2b91a2f8318/",
    help:
      "Зафиксируйте дату поступления, дату явки, вручение или отказ от вручения и обеспечьте возможность своевременной явки.",
    documentHint: "Журнал оповещения и отметка о вручении",
  },
  {
    id: "employee_verification",
    title: "Провести сверку с документами сотрудника",
    shortTitle: "Сверка с сотрудником",
    days: 365,
    workingDays: false,
    source: "Постановление Правительства РФ № 719, п. 32 «г»",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_64215/674f7a847f97b0d98c437823abac7703aa4f5dee/",
    help: "Сведения карточки сверяются с документами воинского учёта не реже одного раза в год.",
    documentHint: "Отметка о сверке в Форме № 10",
  },
  {
    id: "commissariat_verification",
    title: "Провести сверку с военным комиссариатом",
    shortTitle: "Сверка с военкоматом",
    days: 365,
    workingDays: false,
    source: "Приказ Министра обороны РФ № 700, п. 36",
    sourceUrl:
      "https://www.consultant.ru/document/cons_doc_LAW_405064/0ae5a1ee2359bed35fbe57ae19c83c920095c8d4/",
    help:
      "Сверка карточек со сведениями военного комиссариата проводится не реже одного раза в год.",
    documentHint: "Список для сверки и отметка в Форме № 10",
  },
];

const emptySettings: OrganizationSettings = {
  organizationName: "",
  shortName: "",
  organizationAddress: "",
  directorPosition: "Генеральный директор",
  directorName: "",
  responsiblePosition: "Специалист по ведению воинского учёта",
  responsibleName: "Михайленко А.В.",
  responsiblePhone: "",
  defaultCommissariat: "",
  defaultCommissariatAddress: "",
  extraHolidays: "",
};

const navItems: { id: View; label: string; icon: typeof Home }[] = [
  { id: "dashboard", label: "Главная", icon: Home },
  { id: "employees", label: "Сотрудники", icon: Users },
  { id: "notifications", label: "Уведомления", icon: Bell },
  { id: "workPlan", label: "Рабочий план", icon: CalendarCheck },
  { id: "documents", label: "Документы", icon: FileText },
  { id: "reconciliations", label: "Сверки", icon: ListChecks },
  { id: "settings", label: "Настройки", icon: Settings },
];

const requiredFields: { key: keyof Employee; label: string }[] = [
  { key: "fullName", label: "Ф.И.О." },
  { key: "birthDate", label: "дата рождения" },
  { key: "birthPlace", label: "место рождения" },
  { key: "passportSeries", label: "серия паспорта" },
  { key: "passportNumber", label: "номер паспорта" },
  { key: "registrationAddress", label: "адрес регистрации" },
  { key: "militaryDocNumber", label: "документ воинского учёта" },
  { key: "militaryRank", label: "воинское звание" },
  { key: "composition", label: "состав" },
  { key: "vus", label: "ВУС" },
  { key: "reserveCategory", label: "категория запаса" },
  { key: "fitnessCategory", label: "категория годности" },
  { key: "militaryCommissariat", label: "военный комиссариат" },
];

const employeeHistoryFields: { key: keyof Employee; label: string }[] = [
  { key: "fullName", label: "Ф.И.О." }, { key: "sex", label: "пол" },
  { key: "active", label: "статус" }, { key: "department", label: "подразделение" },
  { key: "position", label: "должность" }, { key: "hireDate", label: "дата приёма" },
  { key: "dismissalDate", label: "дата увольнения" }, { key: "orderNumber", label: "номер приказа" },
  { key: "orderDate", label: "дата приказа" }, { key: "birthDate", label: "дата рождения" },
  { key: "birthPlace", label: "место рождения" }, { key: "passportSeries", label: "серия паспорта" },
  { key: "passportNumber", label: "номер паспорта" }, { key: "passportIssueDate", label: "дата выдачи паспорта" },
  { key: "passportIssuedBy", label: "кем выдан паспорт" }, { key: "registrationAddress", label: "адрес регистрации" },
  { key: "registrationDate", label: "дата регистрации" }, { key: "actualAddress", label: "фактический адрес" },
  { key: "phone", label: "телефон" }, { key: "snils", label: "СНИЛС" }, { key: "inn", label: "ИНН" },
  { key: "education", label: "образование" }, { key: "profession", label: "профессия" },
  { key: "languages", label: "иностранные языки" }, { key: "driverLicense", label: "водительское удостоверение" },
  { key: "maritalStatus", label: "семейное положение" }, { key: "familyMembers", label: "состав семьи" },
  { key: "militaryDocType", label: "вид документа воинского учёта" }, { key: "militaryDocNumber", label: "номер документа воинского учёта" },
  { key: "militaryDocIssueDate", label: "дата выдачи документа воинского учёта" }, { key: "militaryDocIssuedBy", label: "кем выдан документ воинского учёта" },
  { key: "militaryRank", label: "воинское звание" }, { key: "composition", label: "состав" },
  { key: "profile", label: "профиль" }, { key: "vus", label: "ВУС" },
  { key: "reserveCategory", label: "категория запаса" }, { key: "fitnessCategory", label: "категория годности" },
  { key: "healthStatus", label: "состояние здоровья" }, { key: "militaryCommissariat", label: "военный комиссариат" },
  { key: "militaryCommissariatAddress", label: "адрес военного комиссариата" },
  { key: "lastEmployeeVerification", label: "сверка с документами сотрудника" },
  { key: "lastCommissariatVerification", label: "сверка с военкоматом" }, { key: "notes", label: "дополнительные сведения" },
];

const enlistedRankOrder = [
  "рядовой", "ефрейтор", "матрос", "старший матрос",
  "младший сержант", "сержант", "старший сержант", "старшина",
  "старшина 2 статьи", "старшина 1 статьи", "главный старшина", "главный корабельный старшина",
  "прапорщик", "старший прапорщик", "мичман", "старший мичман",
];

const officerRankOrder = [
  "младший лейтенант", "лейтенант", "старший лейтенант", "капитан", "капитан лейтенант",
  "майор", "подполковник", "полковник", "капитан 3 ранга", "капитан 2 ранга", "капитан 1 ранга",
  "генерал майор", "генерал лейтенант", "генерал полковник", "генерал армии",
  "маршал российской федерации", "контр адмирал", "вице адмирал", "адмирал", "адмирал флота",
];

const enlistedRanks = new Set(enlistedRankOrder);
const officerRanks = new Set(officerRankOrder);

function normalizedMilitaryRank(value: string) {
  return value
    .toLocaleLowerCase("ru-RU")
    .replaceAll("ё", "е")
    .replace(/\bгвардии\b/g, "")
    .replace(/\bзапаса\b/g, "")
    .replace(/\bв отставке\b/g, "")
    .replace(/\bмедицинской службы\b/g, "")
    .replace(/\bюстиции\b/g, "")
    .replace(/[^а-я0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function militaryRankGroup(rank: string): "enlisted" | "officer" | "unknown" {
  const normalized = normalizedMilitaryRank(rank);
  if (enlistedRanks.has(normalized)) return "enlisted";
  if (officerRanks.has(normalized)) return "officer";
  return "unknown";
}

function militaryRankOrder(rank: string) {
  const normalized = normalizedMilitaryRank(rank);
  const enlistedIndex = enlistedRankOrder.indexOf(normalized);
  if (enlistedIndex >= 0) return enlistedIndex;
  const officerIndex = officerRankOrder.indexOf(normalized);
  return officerIndex >= 0 ? 100 + officerIndex : 999;
}

const federalHolidays2026 = new Set([
  "2026-01-01",
  "2026-01-02",
  "2026-01-03",
  "2026-01-04",
  "2026-01-05",
  "2026-01-06",
  "2026-01-07",
  "2026-01-08",
  "2026-01-09",
  "2026-02-23",
  "2026-03-09",
  "2026-05-01",
  "2026-05-11",
  "2026-06-12",
  "2026-11-04",
  "2026-12-31",
]);

function makeId() {
  return typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
    d.getDate(),
  ).padStart(2, "0")}`;
}

function emptyEmployee(settings: OrganizationSettings): Employee {
  return {
    id: makeId(),
    fullName: "",
    lastName: "",
    firstName: "",
    middleName: "",
    sex: "male",
    active: true,
    department: "",
    position: "",
    hireDate: "",
    dismissalDate: "",
    orderNumber: "",
    orderDate: "",
    birthDate: "",
    birthPlace: "",
    passportSeries: "",
    passportNumber: "",
    passportIssueDate: "",
    passportIssuedBy: "",
    registrationAddress: "",
    registrationDate: "",
    actualAddress: "",
    actualAddressDate: "",
    phone: "",
    workPhone: "",
    snils: "",
    inn: "",
    education: "",
    profession: "",
    languages: "",
    driverLicense: "",
    maritalStatus: "",
    familyMembers: "",
    militaryDocType: "Военный билет",
    militaryDocNumber: "",
    militaryDocIssueDate: "",
    militaryDocIssuedBy: "",
    militaryRank: "",
    composition: "",
    profile: "",
    vus: "",
    reserveCategory: "",
    fitnessCategory: "",
    healthStatus: "",
    militaryCommissariat: settings.defaultCommissariat,
    militaryCommissariatAddress: settings.defaultCommissariatAddress,
    accountType: "",
    teamNumber: "",
    specialAccountNumber: "",
    lastEmployeeVerification: "",
    lastCommissariatVerification: "",
    notes: "",
  };
}

function parseLocalDate(value: string) {
  if (!value) return null;
  const [year, month, day] = value.slice(0, 10).split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

function dateToIso(date: Date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
    date.getDate(),
  ).padStart(2, "0")}`;
}

function formatDate(value: string, fallback = "—") {
  const date = parseLocalDate(value);
  return date
    ? new Intl.DateTimeFormat("ru-RU", { day: "2-digit", month: "2-digit", year: "numeric" }).format(
        date,
      )
    : fallback;
}

function formatLongDate(date = new Date()) {
  return new Intl.DateTimeFormat("ru-RU", {
    day: "numeric",
    month: "long",
    year: "numeric",
    weekday: "long",
  }).format(date);
}

function cleanCell(value: unknown) {
  if (value === null || value === undefined) return "";
  const text = String(value).trim();
  return text === "-" || text === "_" ? "" : text;
}

function excelDateToIso(value: unknown) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateToIso(value);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const text = cleanCell(value);
  const match = text.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})/);
  if (match) return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  return "";
}

function splitName(fullName: string) {
  const parts = fullName.trim().split(/\s+/);
  return {
    lastName: parts[0] ?? "",
    firstName: parts[1] ?? "",
    middleName: parts.slice(2).join(" "),
  };
}

function addCalendarDays(value: string, days: number) {
  const date = parseLocalDate(value) ?? new Date();
  date.setDate(date.getDate() + days);
  return dateToIso(date);
}

function addWorkingDays(value: string, days: number, extraHolidays: string) {
  const date = parseLocalDate(value) ?? new Date();
  const custom = new Set(
    extraHolidays
      .split(/[\s,;]+/)
      .map((item) => item.trim())
      .filter(Boolean),
  );
  let added = 0;
  while (added < days) {
    date.setDate(date.getDate() + 1);
    const iso = dateToIso(date);
    const weekday = date.getDay();
    if (weekday !== 0 && weekday !== 6 && !federalHolidays2026.has(iso) && !custom.has(iso)) {
      added += 1;
    }
  }
  return dateToIso(date);
}

function dueForRule(rule: Rule, eventDate: string, settings: OrganizationSettings) {
  if (rule.days === null) return eventDate;
  return rule.workingDays
    ? addWorkingDays(eventDate, rule.days, settings.extraHolidays)
    : addCalendarDays(eventDate, rule.days);
}

function daysFromToday(value: string) {
  const date = parseLocalDate(value);
  const today = parseLocalDate(todayIso());
  if (!date || !today) return null;
  return Math.round((date.getTime() - today.getTime()) / 86400000);
}

function noticeStatus(notice: Notice) {
  if (notice.completedAt) return "completed";
  const days = daysFromToday(notice.dueDate);
  if (days === null) return "upcoming";
  if (days < 0) return "overdue";
  if (days === 0) return "today";
  return "upcoming";
}

function statusLabel(status: string, dueDate = "") {
  if (status === "completed") return "Исполнено";
  if (status === "overdue") return "Просрочено";
  if (status === "today") return "Срок сегодня";
  const days = daysFromToday(dueDate);
  return days === 1 ? "1 день" : days && days > 1 ? `${days} дн.` : "Предстоит";
}

function getMissingFields(employee: Employee) {
  return requiredFields.filter(({ key }) => !String(employee[key] ?? "").trim()).map(({ label }) => label);
}

function cardCompleteness(employee: Employee) {
  const completed = requiredFields.length - getMissingFields(employee).length;
  return Math.round((completed / requiredFields.length) * 100);
}

function displayChangeValue(change: EmployeeFieldChange, value: string | boolean) {
  if (change.key === "active") return value ? "Работает" : "Уволен";
  if (change.key === "sex") return value === "male" ? "Мужской" : "Женский";
  if (change.key === "familyMembers" && typeof value === "string") {
    const members = familyRows(value);
    return members.length ? members.map((member) => [member.relation, member.fullName, formatDate(member.birthDate, "")].filter(Boolean).join(" — ")).join("; ") : "не указано";
  }
  return String(value || "не указано");
}

type DocumentCheck = { critical: string[]; warnings: string[] };

function checkDocumentEmployee(employee: Employee, type: DocumentType): DocumentCheck {
  const checks: Partial<Record<keyof Employee, string>> = {
    fullName: "Ф.И.О.", birthDate: "дата рождения", birthPlace: "место рождения",
    militaryRank: "воинское звание", reserveCategory: "категория запаса",
    composition: "состав", vus: "ВУС", fitnessCategory: "категория годности",
    militaryDocNumber: "документ воинского учёта", militaryCommissariat: "военный комиссариат",
    registrationAddress: "адрес регистрации", education: "образование",
    maritalStatus: "семейное положение", position: "должность", department: "подразделение",
  };
  const requiredByType: Partial<Record<DocumentType, (keyof Employee)[]>> = {
    form10: ["fullName", "birthDate", "birthPlace", "militaryRank", "reserveCategory", "composition", "vus", "fitnessCategory", "militaryDocNumber", "militaryCommissariat"],
    f2: ["fullName", "birthDate", "birthPlace", "militaryRank", "vus", "militaryDocNumber", "position"],
    employmentNotice: ["fullName", "birthDate", "birthPlace", "militaryRank", "vus", "militaryDocNumber", "position"],
    messageSheet: ["fullName", "birthDate", "registrationAddress", "militaryDocNumber", "militaryCommissariat"],
    officerList: ["fullName", "militaryRank", "reserveCategory", "composition", "vus", "fitnessCategory", "birthDate", "birthPlace", "education", "maritalStatus", "position", "department"],
    enlistedList: ["fullName", "militaryRank", "reserveCategory", "composition", "vus", "fitnessCategory", "birthDate", "birthPlace", "education", "maritalStatus", "position", "department"],
  };
  const warningByType: Partial<Record<DocumentType, (keyof Employee)[]>> = {
    form10: ["registrationAddress", "education", "maritalStatus", "position", "department"],
    f2: ["registrationAddress", "militaryCommissariat"],
    employmentNotice: ["registrationAddress", "militaryCommissariat"],
    messageSheet: ["maritalStatus", "militaryRank", "vus"],
  };
  const missing = (keys: (keyof Employee)[]) => keys
    .filter((key) => !String(employee[key] ?? "").trim())
    .map((key) => checks[key] ?? String(key));
  return {
    critical: missing(requiredByType[type] ?? []),
    warnings: missing(warningByType[type] ?? []),
  };
}

function uniqueValues(employees: Employee[], key: keyof Employee) {
  return Array.from(
    new Set(employees.map((employee) => String(employee[key] ?? "").trim()).filter(Boolean)),
  ).sort((a, b) => a.localeCompare(b, "ru"));
}

type FamilyMember = {
  relation: string;
  fullName: string;
  birthDate: string;
};

function familyRows(value: string): FamilyMember[] {
  return value
    .split(/\r?\n/)
    .map((item) => {
      const parts = item.split("|").map((part) => part.trim());
      const [relation = "", fullName = "", birthDate = ""] =
        parts.length === 1 ? ["", parts[0], ""] : parts;
      return { relation, fullName, birthDate };
    })
    .filter((item) => item.relation || item.fullName || item.birthDate);
}

function serializeFamilyRows(rows: FamilyMember[]) {
  return rows
    .filter((item) => item.relation || item.fullName || item.birthDate)
    .map((item) => [item.relation, item.fullName, item.birthDate].join(" | "))
    .join("\n");
}

function assetUrl(name: string) {
  return `${process.env.NEXT_PUBLIC_BASE_PATH ?? ""}/documents/${name}?v=${encodeURIComponent(APP_VERSION)}`;
}

function bytesFromBase64(value: string) {
  const binary = window.atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function xmlValue(value: unknown) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\r\n", "\n")
    .replaceAll("\r", "\n")
    .replaceAll("\n", '</w:t><w:br/><w:t xml:space="preserve">');
}

function documentMaritalStatus(employee: Employee) {
  const isMarried = /состоит в зарегистрированном браке|женат|замужем/i.test(employee.maritalStatus);
  if (!isMarried) return employee.maritalStatus;
  return employee.sex === "female" ? "замужем" : "женат";
}

function documentValues(
  employee: Employee,
  settings: OrganizationSettings,
  eventType: "hire" | "dismissal",
  f2OrderNumber: string,
  f2OrderDate: string,
  headerLocation: DocumentHeaderLocation,
) {
  const family = familyRows(employee.familyMembers);
  const isMarried = /состоит в зарегистрированном браке|женат|замужем/i.test(employee.maritalStatus);
  const isNotMarried = /не состоит в браке|не замужем|не женат/i.test(employee.maritalStatus);
  const isDivorcedOrWidowed = /развед|вдов/i.test(employee.maritalStatus);
  const wife = isMarried
    ? family.find((item) => /^(жена|супруга)$/i.test(item.relation.trim()))
    : undefined;
  const mother = isNotMarried
    ? family.find((item) => /^мать$/i.test(item.relation.trim()))
    : undefined;
  const father = isNotMarried
    ? family.find((item) => /^отец$/i.test(item.relation.trim()))
    : undefined;
  const children = isMarried || isDivorcedOrWidowed
    ? family.filter((item) => /^(сын|дочь)$/i.test(item.relation.trim()))
    : [];
  const familyDetail = (item?: FamilyMember) => item
    ? [item.fullName, formatDate(item.birthDate, "")].filter(Boolean).join(", ")
    : "";
  const languages = employee.languages.split(/[;,]/).map((item) => item.trim()).filter(Boolean);
  const orderDate = parseLocalDate(f2OrderDate || employee.orderDate);
  const monthNames = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
  ];
  const militaryDocument = [
    employee.militaryDocType,
    employee.militaryDocNumber,
    formatDate(employee.militaryDocIssueDate, ""),
    employee.militaryDocIssuedBy,
  ].filter(Boolean).join(", ");
  return {
    ORG_NAME: settings.organizationName,
    ORG_SHORT_NAME: settings.shortName,
    ORG_ADDRESS: settings.organizationAddress,
    LAST_NAME: employee.lastName,
    FIRST_NAME: employee.firstName,
    MIDDLE_NAME: employee.middleName,
    FULL_NAME: employee.fullName,
    EMPLOYEE_NAME: employee.fullName,
    BIRTH_DATE: formatDate(employee.birthDate, ""),
    BIRTH_YEAR: employee.birthDate ? employee.birthDate.slice(0, 4) : "",
    BIRTH_PLACE: employee.birthPlace,
    EDUCATION: employee.education,
    EDUCATION_ORG: "",
    QUALIFICATION: "",
    SPECIALITY: "",
    DIPLOMA: "",
    GRADUATION_YEAR: "",
    PROFESSION: employee.profession,
    ADDITIONAL_PROFESSION: "",
    MARITAL_STATUS: documentMaritalStatus(employee),
    WIFE_DETAILS: familyDetail(wife),
    MOTHER_DETAILS: familyDetail(mother),
    FATHER_DETAILS: familyDetail(father),
    CHILDREN_DETAILS: children.map((item) => `${item.relation}: ${familyDetail(item)}`).join("; "),
    FAMILY_REL_1: family[0]?.relation ?? "",
    FAMILY_PERSON_1: family[0] ? [family[0].fullName, formatDate(family[0].birthDate, "")].filter(Boolean).join(", ") : "",
    FAMILY_REL_2: family[1]?.relation ?? "",
    FAMILY_PERSON_2: family[1] ? [family[1].fullName, formatDate(family[1].birthDate, "")].filter(Boolean).join(", ") : "",
    FAMILY_REL_3: family[2]?.relation ?? "",
    FAMILY_PERSON_3: family[2] ? [family[2].fullName, formatDate(family[2].birthDate, "")].filter(Boolean).join(", ") : "",
    FAMILY_REL_4: family[3]?.relation ?? "",
    FAMILY_PERSON_4: family[3] ? [family[3].fullName, formatDate(family[3].birthDate, "")].filter(Boolean).join(", ") : "",
    LANGUAGE_1: languages[0] ?? "",
    LANGUAGE_LEVEL_1: languages.length === 1 ? employee.languages : "",
    LANGUAGE_2: languages[1] ?? "",
    LANGUAGE_LEVEL_2: "",
    PASSPORT_SERIES: employee.passportSeries,
    PASSPORT_NUMBER: employee.passportNumber,
    PASSPORT: `${employee.passportSeries} ${employee.passportNumber}`.trim(),
    PASSPORT_ISSUE_DATE: formatDate(employee.passportIssueDate, ""),
    PASSPORT_ISSUED_BY: employee.passportIssuedBy,
    DRIVER_SERIES: "",
    DRIVER_NUMBER: employee.driverLicense,
    DRIVER_CATEGORIES: "",
    DRIVER_ISSUE_DATE: "",
    REGISTRATION_ADDRESS: employee.registrationAddress,
    REGISTRATION_DATE: formatDate(employee.registrationDate, ""),
    ACTUAL_ADDRESS: employee.actualAddress,
    ACTUAL_ADDRESS_DATE: "",
    WORK_PHONE: "",
    MOBILE_PHONE: employee.phone,
    RESERVE_CATEGORY: employee.reserveCategory,
    COMMISSARIAT: employee.militaryCommissariat || settings.defaultCommissariat,
    COMMISSARIAT_SHORT: employee.militaryCommissariat || settings.defaultCommissariat,
    HEADER_AUTHORITY_TYPE: headerLocation === "sochi" ? "городского округа" : "военного",
    HEADER_AUTHORITY_NAME: headerLocation === "sochi"
      ? "город-курорт Сочи Краснодарского края"
      : "комиссариата города Москвы",
    F2_BRANCH_TITLE: headerLocation === "moscow" ? "ФИЛИАЛ АО ЦА «СОЛЯРИС»" : "",
    F2_BRANCH_CITY: headerLocation === "moscow" ? "(город Москва)" : "",
    MILITARY_RANK: employee.militaryRank,
    COMPOSITION_PROFILE: [employee.composition, employee.profile].filter(Boolean).join(" / "),
    TEAM_NUMBER: "",
    SPECIAL_ACCOUNT_NUMBER: "",
    VUS: employee.vus,
    FITNESS_CATEGORY: employee.fitnessCategory,
    HEALTH_STATUS: employee.healthStatus,
    MILITARY_DOCUMENT: militaryDocument,
    ORDER_DETAILS: `${formatDate(employee.orderDate, "")}${employee.orderNumber ? ` № ${employee.orderNumber}` : ""}`,
    POSITION: employee.position,
    DEPARTMENT: employee.department,
    OUTGOING_DETAILS: "",
    RESPONSIBLE_POSITION: settings.responsiblePosition,
    RESPONSIBLE_NAME: settings.responsibleName,
    RESPONSIBLE_PHONE: settings.responsiblePhone,
    SNILS: employee.snils,
    EVENT_HIRE: "принят (поступил) на работу",
    EVENT_DISMISSAL: "уволен с работы (отчислен из образовательной организации)",
    NOTICE_EVENT_HIRE: "принят (поступил)",
    NOTICE_EVENT_DISMISSAL: "уволен с работы (отчислен из образовательной организации)",
    HIRE_STRIKE: eventType === "hire" ? "0" : "1",
    DISMISSAL_STRIKE: eventType === "dismissal" ? "0" : "1",
    ORDER_NUMBER: f2OrderNumber || employee.orderNumber,
    ORDER_DAY: orderDate ? String(orderDate.getDate()).padStart(2, "0") : "",
    ORDER_MONTH: orderDate ? monthNames[orderDate.getMonth()] : "",
    ORDER_YEAR: orderDate ? String(orderDate.getFullYear()).slice(-2) : "",
    DIRECTOR_POSITION: settings.directorPosition,
    DIRECTOR_NAME: settings.directorName,
  };
}

async function buildDocx(
  employee: Employee,
  settings: OrganizationSettings,
  type: DocumentType,
  eventType: "hire" | "dismissal",
  f2OrderNumber: string,
  f2OrderDate: string,
  headerLocation: DocumentHeaderLocation,
) {
  const templateName = type === "form10"
    ? "form10-template.docx"
    : type === "f2"
      ? "f2-template.docx"
      : type === "employmentNotice"
        ? "employment-notice-template.docx"
      : "message-sheet-template.docx";
  const response = await fetch(assetUrl(templateName), { cache: "no-store" });
  if (!response.ok) throw new Error("Не удалось загрузить шаблон Word");
  const archive = unzipSync(new Uint8Array(await response.arrayBuffer()));
  const values = documentValues(employee, settings, eventType, f2OrderNumber, f2OrderDate, headerLocation);
  if (type === "f2" && headerLocation === "moscow") {
    Object.assign(values, {
      ORG_ADDRESS: "Российская Федерация, город Москва, вн. тер. г. муниципальный округ Пресненский, наб. Пресненская, д. 12, помещ. 20/80",
      ORG_NAME: 'Акционерное общество Центр авиации "Солярис"',
      ORG_SHORT_NAME: 'Филиал АО ЦА "Солярис" (город Москва)',
    });
  }
  for (const [path, content] of Object.entries(archive)) {
    if (!path.startsWith("word/") || !path.endsWith(".xml")) continue;
    let xml = strFromU8(content);
    for (const [key, value] of Object.entries(values)) {
      xml = xml.replaceAll(`{{${key}}}`, xmlValue(value));
    }
    xml = xml.replace(/\{\{[A-Z0-9_]+\}\}/g, "");
    archive[path] = strToU8(xml);
  }
  return new Blob([zipSync(archive, { level: 6 }) as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function buildChangesDocx(
  entries: ChangeDocumentEntry[],
  employees: Employee[],
  settings: OrganizationSettings,
  headerLocation: DocumentHeaderLocation,
) {
  const response = await fetch(assetUrl("changes-template.docx"), { cache: "no-store" });
  if (!response.ok) throw new Error("Не удалось загрузить шаблон Word");
  const archive = unzipSync(new Uint8Array(await response.arrayBuffer()));
  const documentPath = "word/document.xml";
  let xml = strFromU8(archive[documentPath]);
  const marker = "{{CHANGE_ROW_NUMBER}}";
  const markerIndex = xml.indexOf(marker);
  if (markerIndex < 0) throw new Error("В шаблоне отсутствует строка сведений");
  const rowMatches = Array.from(xml.slice(0, markerIndex).matchAll(/<w:tr(?=[\s>])/g));
  const rowStart = rowMatches.at(-1)?.index ?? -1;
  const rowEnd = xml.indexOf("</w:tr>", markerIndex) + "</w:tr>".length;
  if (rowStart < 0 || rowEnd < rowStart) throw new Error("Не удалось прочитать строку сведений");
  const rowTemplate = xml.slice(rowStart, rowEnd);
  const rows = entries.map((entry, index) => {
    const employee = employees.find((item) => item.id === entry.employeeId);
    if (!employee) return "";
    const rowValues: Record<string, string> = {
      CHANGE_ROW_NUMBER: String(index + 1),
      CHANGE_FULL_NAME: employee.fullName,
      CHANGE_MILITARY_RANK: employee.militaryRank,
      CHANGE_BIRTH_YEAR: employee.birthDate ? employee.birthDate.slice(0, 4) : "",
      CHANGE_ACCOUNT: "Общий",
      CHANGE_CONTENT: entry.content,
      CHANGE_NOTE: "",
    };
    let row = rowTemplate;
    for (const [key, value] of Object.entries(rowValues)) {
      row = row.replaceAll(`{{${key}}}`, xmlValue(value));
    }
    return row;
  }).join("");
  xml = `${xml.slice(0, rowStart)}${rows}${xml.slice(rowEnd)}`;

  const firstEmployee = employees.find((item) => item.id === entries[0]?.employeeId);
  if (!firstEmployee) throw new Error("Не выбран сотрудник");
  const values = documentValues(firstEmployee, settings, "hire", "", "", headerLocation);
  for (const [key, value] of Object.entries(values)) {
    xml = xml.replaceAll(`{{${key}}}`, xmlValue(value));
  }
  xml = xml.replace(/\{\{[A-Z0-9_]+\}\}/g, "");
  archive[documentPath] = strToU8(xml);
  return new Blob([zipSync(archive, { level: 6 }) as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function buildPersonnelListDocx(
  type: "officerList" | "enlistedList",
  employeeIds: string[],
  employees: Employee[],
  settings: OrganizationSettings,
  headerLocation: DocumentHeaderLocation,
) {
  const templateName = type === "officerList" ? "officer-list-template.docx" : "enlisted-list-template.docx";
  let archive: ReturnType<typeof unzipSync>;
  try {
    const response = await fetch(assetUrl(templateName), { cache: "no-store" });
    if (!response.ok) throw new Error(`Шаблон недоступен: ${response.status}`);
    archive = unzipSync(new Uint8Array(await response.arrayBuffer()));
  } catch {
    archive = unzipSync(bytesFromBase64(LIST_TEMPLATE_BASE64[type]));
  }
  const documentPath = "word/document.xml";
  let xml = strFromU8(archive[documentPath]);
  const marker = "{{LIST_ROW_NUMBER}}";
  const markerIndex = xml.indexOf(marker);
  if (markerIndex < 0) throw new Error("В шаблоне отсутствует строка списка");
  const rowMatches = Array.from(xml.slice(0, markerIndex).matchAll(/<w:tr(?=[\s>])/g));
  const rowStart = rowMatches.at(-1)?.index ?? -1;
  const rowEnd = xml.indexOf("</w:tr>", markerIndex) + "</w:tr>".length;
  if (rowStart < 0 || rowEnd < rowStart) throw new Error("Не удалось прочитать строку списка");
  const rowTemplate = xml.slice(rowStart, rowEnd);
  const selectedEmployees = employeeIds
    .map((id) => employees.find((employee) => employee.id === id))
    .filter((employee): employee is Employee => Boolean(employee));
  const rows = selectedEmployees.map((employee, index) => {
    const rowValues: Record<string, string> = {
      LIST_ROW_NUMBER: String(index + 1),
      LIST_FULL_NAME: employee.fullName,
      LIST_MILITARY_RANK: employee.militaryRank,
      LIST_RESERVE_CATEGORY: employee.reserveCategory,
      LIST_COMPOSITION_PROFILE: [employee.composition, employee.profile].filter(Boolean).join(" / "),
      LIST_VUS: employee.vus,
      LIST_FITNESS_CATEGORY: employee.fitnessCategory,
      LIST_ACCOUNT: employee.accountType === "special" ? "Специальный" : "Общий",
      LIST_BIRTH_DETAILS: [formatDate(employee.birthDate, ""), employee.birthPlace].filter(Boolean).join(", "),
      LIST_EDUCATION: employee.education,
      LIST_ADDRESS: employee.actualAddress || employee.registrationAddress,
      LIST_MARITAL_STATUS: documentMaritalStatus(employee),
      LIST_POSITION_DEPARTMENT: [employee.position, employee.department ? `(${employee.department})` : ""].filter(Boolean).join(" "),
    };
    let row = rowTemplate;
    for (const [key, value] of Object.entries(rowValues)) row = row.replaceAll(`{{${key}}}`, xmlValue(value));
    return row;
  }).join("");
  xml = `${xml.slice(0, rowStart)}${rows}${xml.slice(rowEnd)}`;
  const firstEmployee = selectedEmployees[0];
  if (!firstEmployee) throw new Error("Не выбран сотрудник");
  const values = documentValues(firstEmployee, settings, "hire", "", "", headerLocation);
  for (const [key, value] of Object.entries(values)) xml = xml.replaceAll(`{{${key}}}`, xmlValue(value));
  xml = xml.replace(/\{\{[A-Z0-9_]+\}\}/g, "");
  archive[documentPath] = strToU8(xml);
  return new Blob([zipSync(archive, { level: 6 }) as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

function downloadBlob(blob: Blob, filename: string) {
  const link = window.document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  window.document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(link.href), 1500);
}

function Field({
  label,
  value,
  onChange,
  type = "text",
  placeholder,
  help,
  required,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  type?: string;
  placeholder?: string;
  help?: string;
  required?: boolean;
}) {
  return (
    <label className="field">
      <span>
        {label}
        {required ? " *" : ""}
        {help ? <span className="field-help" aria-label={help} title={help}><Info size={14} /></span> : null}
      </span>
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function FamilyMembersEditor({
  value,
  onChange,
}: {
  value: string;
  onChange: (value: string) => void;
}) {
  const relationshipOptions = ["Отец", "Мать", "Жена", "Сын", "Дочь", "Сестра", "Брат"];
  const [rows, setRows] = useState<FamilyMember[]>(() => {
    const parsed = familyRows(value);
    return parsed.length ? parsed : [{ relation: "", fullName: "", birthDate: "" }];
  });

  // Only reset the row editor when its serialized value changes externally.
  useEffect(() => {
    if (serializeFamilyRows(rows) === value) return;
    const parsed = familyRows(value);
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setRows(parsed.length ? parsed : [{ relation: "", fullName: "", birthDate: "" }]);
    // The serialized value is the source of truth when another employee is opened.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  function updateRow(index: number, patch: Partial<FamilyMember>) {
    const next = rows.map((row, rowIndex) => rowIndex === index ? { ...row, ...patch } : row);
    setRows(next);
    onChange(serializeFamilyRows(next));
  }

  function removeRow(index: number) {
    const next = rows.filter((_, rowIndex) => rowIndex !== index);
    const normalized = next.length ? next : [{ relation: "", fullName: "", birthDate: "" }];
    setRows(normalized);
    onChange(serializeFamilyRows(normalized));
  }

  return (
    <div className="family-editor">
      <div className="family-editor-heading">
        <span>Состав семьи</span>
        <small>Степень родства, Ф.И.О. и дата рождения</small>
      </div>
      <div className="family-row family-row-head" aria-hidden="true">
        <span>Степень родства</span><span>Ф.И.О.</span><span>Дата рождения</span><span />
      </div>
      {rows.map((row, index) => (
        <div className="family-row" key={index}>
          <select value={row.relation} aria-label="Степень родства" onChange={(event) => updateRow(index, { relation: event.target.value })}>
            <option value="">Выберите</option>
            {row.relation && !relationshipOptions.includes(row.relation) ? <option value={row.relation}>{row.relation}</option> : null}
            {relationshipOptions.map((relation) => <option value={relation} key={relation}>{relation}</option>)}
          </select>
          <input value={row.fullName} placeholder="Фамилия Имя Отчество" onChange={(event) => updateRow(index, { fullName: event.target.value })} />
          <input type="date" value={row.birthDate} onChange={(event) => updateRow(index, { birthDate: event.target.value })} />
          <button type="button" className="icon-button" aria-label="Удалить члена семьи" onClick={() => removeRow(index)}><Trash2 size={16} /></button>
        </div>
      ))}
      <button type="button" className="button secondary small family-add" onClick={() => setRows([...rows, { relation: "", fullName: "", birthDate: "" }])}><Plus size={15} /> Добавить члена семьи</button>
    </div>
  );
}

export default function HomePage() {
  const [view, setView] = useState<View>("dashboard");
  const [menuOpen, setMenuOpen] = useState(false);
  const [hydrated, setHydrated] = useState(false);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [notices, setNotices] = useState<Notice[]>([]);
  const [documents, setDocuments] = useState<DocumentRecord[]>([]);
  const [employeeChanges, setEmployeeChanges] = useState<EmployeeChangeRecord[]>([]);
  const [organization, setOrganization] = useState<OrganizationSettings>(emptySettings);
  const [employeeModal, setEmployeeModal] = useState<Employee | null>(null);
  const [eventModal, setEventModal] = useState(false);
  const [importPreview, setImportPreview] = useState<{
    filename: string;
    employees: Employee[];
    warnings: string[];
  } | null>(null);
  const [toast, setToast] = useState("");
  const [employeeSearch, setEmployeeSearch] = useState("");
  const [departmentFilter, setDepartmentFilter] = useState("");
  const [attentionFilter, setAttentionFilter] = useState("");
  const [activeFilter, setActiveFilter] = useState("active");
  const [sortField, setSortField] = useState<keyof Employee>("fullName");
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");
  const [noticeSearch, setNoticeSearch] = useState("");
  const [noticeFilter, setNoticeFilter] = useState("open");
  const [workPlanMonth, setWorkPlanMonth] = useState(todayIso().slice(0, 7));
  const [workPlanStatus, setWorkPlanStatus] = useState("all");
  const [documentEmployeeId, setDocumentEmployeeId] = useState("");
  const [documentType, setDocumentType] = useState<DocumentType>("form10");
  const [f2Event, setF2Event] = useState<"hire" | "dismissal">("hire");
  const [f2OrderNumber, setF2OrderNumber] = useState("");
  const [f2OrderDate, setF2OrderDate] = useState("");
  const [documentHeaderLocation, setDocumentHeaderLocation] = useState<DocumentHeaderLocation>("moscow");
  const [lastBackupAt, setLastBackupAt] = useState("");
  const [changeEmployeeToAdd, setChangeEmployeeToAdd] = useState("");
  const [changeDocumentEntries, setChangeDocumentEntries] = useState<ChangeDocumentEntry[]>([]);
  const [documentHistoryType, setDocumentHistoryType] = useState<DocumentType | "">("");
  const [documentHistoryEmployee, setDocumentHistoryEmployee] = useState("");
  const [documentHistoryStatus, setDocumentHistoryStatus] = useState("");
  const [reconciliationSearch, setReconciliationSearch] = useState("");
  const [reconciliationFilter, setReconciliationFilter] = useState("attention");
  const excelInput = useRef<HTMLInputElement>(null);
  const backupInput = useRef<HTMLInputElement>(null);

  const [eventDraft, setEventDraft] = useState({
    employeeId: "",
    ruleId: "hire",
    eventDate: todayIso(),
    dueDate: "",
    note: "",
  });

  /* Local storage is the application's external persistence layer. Loading it
     once after mount intentionally hydrates all four related state slices. */
  /* eslint-disable react-hooks/set-state-in-effect */
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const stored = JSON.parse(raw) as Partial<StoredState>;
        setEmployees(Array.isArray(stored.employees) ? stored.employees : []);
        setNotices(Array.isArray(stored.notices) ? stored.notices : []);
        setDocuments(Array.isArray(stored.documents) ? stored.documents : []);
        setEmployeeChanges(Array.isArray(stored.employeeChanges) ? stored.employeeChanges : []);
        setOrganization({ ...emptySettings, ...(stored.settings ?? {}) });
      }
      setLastBackupAt(localStorage.getItem(LAST_BACKUP_KEY) ?? "");
    } catch {
      setToast("Не удалось прочитать сохранённые данные. Можно восстановить резервную копию.");
    } finally {
      setHydrated(true);
    }
  }, []);
  /* eslint-enable react-hooks/set-state-in-effect */

  useEffect(() => {
    if (!hydrated) return;
    const state: StoredState = { employees, notices, documents, employeeChanges, settings: organization };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [employees, notices, documents, employeeChanges, organization, hydrated]);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 3800);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const automaticListEntries = useMemo<ChangeDocumentEntry[]>(() => {
    if (documentType !== "officerList" && documentType !== "enlistedList") return [];
    const targetGroup = documentType === "officerList" ? "officer" : "enlisted";
    return employees
      .filter((employee) => employee.active && militaryRankGroup(employee.militaryRank) === targetGroup)
      .sort((a, b) => militaryRankOrder(a.militaryRank) - militaryRankOrder(b.militaryRank) || a.fullName.localeCompare(b.fullName, "ru"))
      .map((employee) => ({ employeeId: employee.id, content: "" }));
  }, [documentType, employees]);

  const employeesWithUnknownRank = useMemo(
    () => employees.filter((employee) => employee.active && militaryRankGroup(employee.militaryRank) === "unknown"),
    [employees],
  );

  const selectedDocumentEmployee = employees.find((employee) => employee.id === documentEmployeeId);
  const selectedDocumentCheck = selectedDocumentEmployee
    ? checkDocumentEmployee(selectedDocumentEmployee, documentType)
    : { critical: [], warnings: [] };
  const listEmployeesWithGaps = automaticListEntries
    .map((entry) => employees.find((employee) => employee.id === entry.employeeId))
    .filter((employee): employee is Employee => Boolean(employee))
    .map((employee) => ({ employee, check: checkDocumentEmployee(employee, documentType) }))
    .filter(({ check }) => check.critical.length > 0);
  const listRankDistribution = automaticListEntries.reduce<Record<string, number>>((result, entry) => {
    const rank = employees.find((employee) => employee.id === entry.employeeId)?.militaryRank || "Не указано";
    result[rank] = (result[rank] ?? 0) + 1;
    return result;
  }, {});
  const documentSettingsMissing = [
    !organization.organizationName && "наименование организации",
    !organization.organizationAddress && "адрес организации",
    !organization.responsiblePosition && "должность ответственного",
    !organization.responsibleName && "Ф.И.О. ответственного",
  ].filter((value): value is string => Boolean(value));
  const documentParameterMissing = documentType === "f2" || documentType === "employmentNotice"
    ? [!f2OrderNumber.trim() && "номер приказа", !f2OrderDate && "дата приказа"].filter((value): value is string => Boolean(value))
    : [];
  const hasCriticalDocumentGaps = documentType === "officerList" || documentType === "enlistedList"
    ? listEmployeesWithGaps.length > 0 || documentSettingsMissing.length > 0
    : documentType === "changes"
      ? documentSettingsMissing.length > 0
      : selectedDocumentCheck.critical.length > 0 || documentSettingsMissing.length > 0 || documentParameterMissing.length > 0;
  const filteredDocumentHistory = documents.filter((record) => {
    if (documentHistoryType && record.type !== documentHistoryType) return false;
    if (documentHistoryStatus && (record.status ?? "formed") !== documentHistoryStatus) return false;
    if (documentHistoryEmployee && record.employeeId !== documentHistoryEmployee && !record.employeeIds?.includes(documentHistoryEmployee)) return false;
    return true;
  });

  const backupAgeDays = lastBackupAt
    ? Math.floor((Date.now() - new Date(lastBackupAt).getTime()) / 86_400_000)
    : null;
  const backupIsDue = employees.length > 0 && (backupAgeDays === null || backupAgeDays >= 7);

  const pendingNotices = useMemo(
    () => notices.filter((notice) => !notice.completedAt),
    [notices],
  );
  const incompleteEmployees = useMemo(
    () => employees.filter((employee) => employee.active && getMissingFields(employee).length > 0),
    [employees],
  );
  const overdueCount = pendingNotices.filter((notice) => noticeStatus(notice) === "overdue").length;
  const todayCount = pendingNotices.filter((notice) => noticeStatus(notice) === "today").length;
  const attentionCount = pendingNotices.length + incompleteEmployees.length;

  const filteredEmployees = useMemo(() => {
    const search = employeeSearch.trim().toLowerCase();
    return employees
      .filter((employee) => {
        if (activeFilter === "active" && !employee.active) return false;
        if (activeFilter === "dismissed" && employee.active) return false;
        if (departmentFilter && employee.department !== departmentFilter) return false;
        if (attentionFilter === "missing" && getMissingFields(employee).length === 0) return false;
        if (
          search &&
          ![
            employee.fullName,
            employee.snils,
            employee.passportSeries,
            employee.passportNumber,
            employee.vus,
          ]
            .join(" ")
            .toLowerCase()
            .includes(search)
        )
          return false;
        return true;
      })
      .sort((a, b) => {
        const left = String(a[sortField] ?? "");
        const right = String(b[sortField] ?? "");
        const result = left.localeCompare(right, "ru", { numeric: true });
        return sortDirection === "asc" ? result : -result;
      });
  }, [
    employees,
    employeeSearch,
    activeFilter,
    departmentFilter,
    attentionFilter,
    sortField,
    sortDirection,
  ]);

  const filteredNotices = useMemo(() => {
    const search = noticeSearch.trim().toLowerCase();
    return notices
      .filter((notice) => {
        const employee = employees.find((item) => item.id === notice.employeeId);
        const rule = RULES.find((item) => item.id === notice.ruleId);
        const status = noticeStatus(notice);
        if (noticeFilter === "open" && status === "completed") return false;
        if (noticeFilter !== "all" && noticeFilter !== "open" && status !== noticeFilter) return false;
        if (search && !`${employee?.fullName ?? ""} ${rule?.title ?? ""}`.toLowerCase().includes(search))
          return false;
        return true;
      })
      .sort((a, b) => {
        if (a.completedAt && !b.completedAt) return 1;
        if (!a.completedAt && b.completedAt) return -1;
        return a.dueDate.localeCompare(b.dueDate);
      });
  }, [notices, employees, noticeSearch, noticeFilter]);

  const workPlanRows = useMemo(() => notices
    .filter((notice) => notice.dueDate.startsWith(workPlanMonth))
    .filter((notice) => workPlanStatus === "all" || (workPlanStatus === "completed" ? Boolean(notice.completedAt) : !notice.completedAt))
    .map((notice) => {
      const employee = employees.find((item) => item.id === notice.employeeId);
      const rule = RULES.find((item) => item.id === notice.ruleId);
      const document = documents.find((item) => item.noticeIds?.includes(notice.id));
      return { notice, employee, rule, document, status: noticeStatus(notice) };
    })
    .sort((a, b) => a.notice.dueDate.localeCompare(b.notice.dueDate) || (a.employee?.fullName ?? "").localeCompare(b.employee?.fullName ?? "", "ru")),
  [notices, employees, documents, workPlanMonth, workPlanStatus]);

  const workPlanSummary = {
    total: workPlanRows.length,
    open: workPlanRows.filter((row) => !row.notice.completedAt).length,
    overdue: workPlanRows.filter((row) => row.status === "overdue").length,
    completed: workPlanRows.filter((row) => Boolean(row.notice.completedAt)).length,
  };

  const duplicateEmployeeGroups = useMemo(() => {
    const groups = new Map<string, Employee[]>();
    employees.forEach((employee) => {
      const key = `${employee.fullName.trim().toLocaleLowerCase("ru-RU")}|${employee.birthDate}`;
      if (!employee.fullName.trim()) return;
      groups.set(key, [...(groups.get(key) ?? []), employee]);
    });
    return [...groups.values()].filter((group) => group.length > 1);
  }, [employees]);
  const orphanNotices = notices.filter((notice) => !employees.some((employee) => employee.id === notice.employeeId));
  const orphanDocuments = documents.filter((record) => {
    const ids = record.employeeIds?.length ? record.employeeIds : [record.employeeId];
    return ids.some((id) => !employees.some((employee) => employee.id === id));
  });
  const incompleteActiveEmployees = employees.filter((employee) => employee.active && getMissingFields(employee).length);
  const diagnosticIssues: { level: "error" | "warning"; title: string; detail: string; view?: View }[] = [];
  if (orphanNotices.length) diagnosticIssues.push({ level: "error", title: "Задачи без карточки сотрудника", detail: `${orphanNotices.length} шт.`, view: "notifications" });
  if (orphanDocuments.length) diagnosticIssues.push({ level: "error", title: "Документы с отсутствующей карточкой", detail: `${orphanDocuments.length} шт.`, view: "documents" });
  if (duplicateEmployeeGroups.length) diagnosticIssues.push({ level: "error", title: "Возможные дубли сотрудников", detail: `${duplicateEmployeeGroups.length} совпадений`, view: "employees" });
  if (incompleteActiveEmployees.length) diagnosticIssues.push({ level: "warning", title: "Неполные карточки работающих", detail: `${incompleteActiveEmployees.length} шт.`, view: "employees" });
  if (overdueCount) diagnosticIssues.push({ level: "warning", title: "Просроченные задачи", detail: `${overdueCount} шт.`, view: "notifications" });
  if (employeesWithUnknownRank.length) diagnosticIssues.push({ level: "warning", title: "Нераспознанные воинские звания", detail: `${employeesWithUnknownRank.length} шт.`, view: "employees" });
  if (documentSettingsMissing.length) diagnosticIssues.push({ level: "warning", title: "Не заполнены настройки документов", detail: documentSettingsMissing.join(", "), view: "settings" });
  if (backupIsDue) diagnosticIssues.push({ level: "warning", title: "Требуется резервная копия", detail: lastBackupAt ? `Последняя создана ${backupAgeDays} дн. назад` : "Резервная копия не создавалась", view: "settings" });
  const diagnosticErrorCount = diagnosticIssues.filter((item) => item.level === "error").length;

  const dashboardItems = useMemo(() => {
    const legal = pendingNotices
      .map((notice) => ({
        id: notice.id,
        employee: employees.find((employee) => employee.id === notice.employeeId),
        title: RULES.find((rule) => rule.id === notice.ruleId)?.title ?? "Задача",
        dueDate: notice.dueDate,
        status: noticeStatus(notice),
        type: "notice" as const,
      }))
      .sort((a, b) => a.dueDate.localeCompare(b.dueDate));
    const completeness = incompleteEmployees.map((employee) => ({
      id: `missing-${employee.id}`,
      employee,
      title: `Заполнить карточку: отсутствует ${getMissingFields(employee).slice(0, 2).join(", ")}`,
      dueDate: "",
      status: "missing",
      type: "missing" as const,
    }));
    return [...legal, ...completeness].slice(0, 10);
  }, [pendingNotices, employees, incompleteEmployees]);

  function navigate(next: View) {
    setView(next);
    setMenuOpen(false);
  }

  function openNewEmployee() {
    setEmployeeModal(emptyEmployee(organization));
  }

  function saveEmployee(event: FormEvent) {
    event.preventDefault();
    if (!employeeModal?.fullName.trim()) {
      setToast("Укажите Ф.И.О. сотрудника.");
      return;
    }
    const parts = splitName(employeeModal.fullName);
    const saved = { ...employeeModal, ...parts };
    const previous = employees.find((employee) => employee.id === saved.id);
    const fieldChanges: EmployeeFieldChange[] = previous ? employeeHistoryFields
      .filter(({ key }) => previous[key] !== saved[key])
      .map(({ key, label }) => ({
        key,
        label,
        oldValue: previous[key] as string | boolean,
        newValue: saved[key] as string | boolean,
      })) : [];
    setEmployees((current) => {
      const exists = current.some((employee) => employee.id === saved.id);
      return exists
        ? current.map((employee) => (employee.id === saved.id ? saved : employee))
        : [saved, ...current];
    });
    if (fieldChanges.length) {
      setEmployeeChanges((current) => [{
        id: makeId(),
        employeeId: saved.id,
        employeeName: saved.fullName,
        createdAt: new Date().toISOString(),
        changes: fieldChanges,
      }, ...current]);
    }
    setEmployeeModal(null);
    const trackedChanges: { key: keyof Employee; label: string }[] = [
      { key: "maritalStatus", label: "семейное положение" },
      { key: "education", label: "образование" },
      { key: "department", label: "подразделение" },
      { key: "position", label: "должность" },
      { key: "registrationAddress", label: "адрес регистрации" },
      { key: "actualAddress", label: "фактический адрес" },
      { key: "healthStatus", label: "состояние здоровья" },
    ];
    const changedLabels = trackedChanges.filter(({ key }) => fieldChanges.some((change) => change.key === key)).map(({ label }) => label);
    const suggestedRule = !previous
      ? "hire"
      : previous.active && !saved.active
        ? "dismissal"
        : changedLabels.length
          ? "change"
          : "";
    if (suggestedRule) {
      const rule = RULES.find((item) => item.id === suggestedRule) ?? RULES[0];
      const eventDate = suggestedRule === "hire"
        ? saved.hireDate || todayIso()
        : suggestedRule === "dismissal"
          ? saved.dismissalDate || todayIso()
          : todayIso();
      const reason = suggestedRule === "change" ? ` Изменено: ${changedLabels.join(", ")}.` : "";
      if (window.confirm(`Карточка сохранена.${reason}\n\nСоздать задачу «${rule.title}»?`)) {
        setEventDraft({
          employeeId: saved.id,
          ruleId: suggestedRule,
          eventDate,
          dueDate: dueForRule(rule, eventDate, organization),
          note: suggestedRule === "change" ? `Изменено: ${changedLabels.join(", ")}` : "",
        });
        setEventModal(true);
        return;
      }
    }
    setToast("Карточка сотрудника сохранена.");
  }

  function undoEmployeeChange(record: EmployeeChangeRecord) {
    const latest = employeeChanges.find((item) => item.employeeId === record.employeeId && !item.undoneAt);
    if (!latest || latest.id !== record.id) {
      setToast("Отменить можно только последнее активное изменение карточки.");
      return;
    }
    if (!window.confirm(`Отменить последнее изменение карточки «${record.employeeName}»?`)) return;
    setEmployees((current) => current.map((employee) => {
      if (employee.id !== record.employeeId) return employee;
      const restored = { ...employee } as Employee;
      const writable = restored as unknown as Record<string, string | boolean>;
      record.changes.forEach((change) => { writable[change.key] = change.oldValue; });
      if (record.changes.some((change) => change.key === "fullName")) Object.assign(restored, splitName(restored.fullName));
      return restored;
    }));
    setEmployeeChanges((current) => current.map((item) => item.id === record.id ? { ...item, undoneAt: new Date().toISOString() } : item));
    setEmployeeModal((current) => {
      if (!current || current.id !== record.employeeId) return current;
      const restored = { ...current } as Employee;
      const writable = restored as unknown as Record<string, string | boolean>;
      record.changes.forEach((change) => { writable[change.key] = change.oldValue; });
      if (record.changes.some((change) => change.key === "fullName")) Object.assign(restored, splitName(restored.fullName));
      return restored;
    });
    setToast("Последнее изменение отменено.");
  }

  function prepareChangeDocument(record: EmployeeChangeRecord) {
    const employee = employees.find((item) => item.id === record.employeeId);
    if (!employee) {
      setToast("Карточка сотрудника не найдена.");
      return;
    }
    const content = record.changes.map((change) => `${change.label}: «${displayChangeValue(change, change.oldValue)}» → «${displayChangeValue(change, change.newValue)}»`).join("; ");
    setChangeDocumentEntries([{ employeeId: record.employeeId, content }]);
    setDocumentType("changes");
    setDocumentEmployeeId(record.employeeId);
    setEmployeeModal(null);
    setView("documents");
    window.scrollTo({ top: 0, behavior: "smooth" });
    setToast("Изменения перенесены в документ. Проверьте формулировку перед выгрузкой.");
  }

  function deleteEmployee(employee: Employee) {
    if (!window.confirm(`Удалить карточку «${employee.fullName}» и связанные задачи?`)) return;
    setEmployees((current) => current.filter((item) => item.id !== employee.id));
    setNotices((current) => current.filter((notice) => notice.employeeId !== employee.id));
    setDocuments((current) => current.filter((document) => document.employeeId !== employee.id));
    setEmployeeChanges((current) => current.filter((record) => record.employeeId !== employee.id));
    setEmployeeModal(null);
    setToast("Карточка удалена.");
  }

  function toggleSort(field: keyof Employee) {
    if (sortField === field) setSortDirection((current) => (current === "asc" ? "desc" : "asc"));
    else {
      setSortField(field);
      setSortDirection("asc");
    }
  }

  function resetEmployeeFilters() {
    setEmployeeSearch("");
    setDepartmentFilter("");
    setAttentionFilter("");
    setActiveFilter("active");
  }

  function rowValue(row: Record<string, unknown>, ...needles: string[]) {
    const entries = Object.entries(row);
    const match = entries.find(([key]) => {
      const normalized = key.toLowerCase().replaceAll("ё", "е").trim();
      return needles.some((needle) => normalized.includes(needle));
    });
    return match?.[1];
  }

  function mapExcelRow(row: Record<string, unknown>) {
    const fullName = cleanCell(rowValue(row, "фамилия, имя", "фио", "ф.и.о"));
    if (!fullName) return null;
    const name = splitName(fullName);
    const orderText = cleanCell(rowValue(row, "приказ о приеме", "приказ о приёме"));
    const orderNumber = orderText.match(/№\s*([^\s]+)/)?.[1] ?? "";
    const orderDate = excelDateToIso(orderText);
    const passport = cleanCell(rowValue(row, "паспорт"));
    const passportMatch = passport.match(/(\d{4})\s+(\d{6})/);
    const militaryDoc = cleanCell(rowValue(row, "военный билет"));
    return {
      ...emptyEmployee(organization),
      ...name,
      fullName,
      department: cleanCell(rowValue(row, "подразделение")),
      position: cleanCell(rowValue(row, "должность")),
      orderNumber,
      orderDate,
      hireDate: orderDate,
      birthDate: excelDateToIso(rowValue(row, "дата рождения")),
      birthPlace: cleanCell(rowValue(row, "место рождения")),
      passportSeries: passportMatch?.[1] ?? "",
      passportNumber: passportMatch?.[2] ?? "",
      passportIssueDate: excelDateToIso(passport.replace(passportMatch?.[0] ?? "", "")),
      passportIssuedBy: passport
        .replace(passportMatch?.[0] ?? "", "")
        .replace(/\bот\b\s*\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4}/i, "")
        .trim(),
      registrationAddress: cleanCell(rowValue(row, "прописка", "адрес регистрации")),
      militaryDocNumber: militaryDoc,
      militaryRank: cleanCell(rowValue(row, "звание")),
      composition: cleanCell(rowValue(row, "состав")),
      vus: cleanCell(rowValue(row, "вус")),
      profile: cleanCell(rowValue(row, "профиль")),
      reserveCategory: cleanCell(rowValue(row, "запас")),
      fitnessCategory: cleanCell(rowValue(row, "категория годности")),
      healthStatus: cleanCell(rowValue(row, "состояние здоровья", "здоровье")),
      militaryCommissariat: cleanCell(rowValue(row, "военный комиссариат")),
      militaryCommissariatAddress: cleanCell(rowValue(row, "адрес вк")),
      inn: cleanCell(rowValue(row, "инн")),
      snils: cleanCell(rowValue(row, "снилс")),
      maritalStatus: cleanCell(rowValue(row, "семейное")),
      familyMembers: cleanCell(rowValue(row, "дети", "состав семьи")),
    } satisfies Employee;
  }

  async function importExcel(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
      const mapped = rows.map(mapExcelRow).filter((item): item is Employee => Boolean(item));
      if (!mapped.length) {
        setToast("В таблице не найден столбец с Ф.И.О. сотрудников.");
        return;
      }
      const warnings: string[] = [];
      const incomplete = mapped.filter((employee) => getMissingFields(employee).length > 0).length;
      if (incomplete) warnings.push(`Неполные карточки: ${incomplete}`);
      const duplicateNames = mapped.length - new Set(mapped.map((item) => `${item.fullName}|${item.birthDate}`)).size;
      if (duplicateNames) warnings.push(`Повторы внутри файла: ${duplicateNames}`);
      setImportPreview({ filename: file.name, employees: mapped, warnings });
    } catch {
      setToast("Не удалось прочитать файл. Используйте XLSX, XLS или CSV.");
    }
  }

  function commitImport(mode: "merge" | "replace") {
    if (!importPreview) return;
    if (employees.length) exportBackup();
    if (mode === "replace") {
      setEmployees(importPreview.employees);
      setNotices([]);
      setDocuments([]);
      setEmployeeChanges([]);
    } else {
      setEmployees((current) => {
        const next = [...current];
        importPreview.employees.forEach((incoming) => {
          const index = next.findIndex(
            (employee) =>
              (incoming.snils && employee.snils === incoming.snils) ||
              (employee.fullName.toLowerCase() === incoming.fullName.toLowerCase() &&
                employee.birthDate === incoming.birthDate),
          );
          if (index >= 0) next[index] = { ...next[index], ...incoming, id: next[index].id };
          else next.push(incoming);
        });
        return next;
      });
    }
    setToast(`Импортировано карточек: ${importPreview.employees.length}. Исторические уведомления не создавались.`);
    setImportPreview(null);
    setView("employees");
  }

  function openEventModal(employeeId = "", ruleId = "hire") {
    const rule = RULES.find((item) => item.id === ruleId) ?? RULES[0];
    const eventDate = todayIso();
    setEventDraft({
      employeeId: employeeId || employees[0]?.id || "",
      ruleId: rule.id,
      eventDate,
      dueDate: dueForRule(rule, eventDate, organization),
      note: "",
    });
    setEventModal(true);
  }

  function changeEventRule(ruleId: string) {
    const rule = RULES.find((item) => item.id === ruleId) ?? RULES[0];
    setEventDraft((current) => ({
      ...current,
      ruleId,
      dueDate: dueForRule(rule, current.eventDate, organization),
    }));
  }

  function changeEventDate(eventDate: string) {
    const rule = RULES.find((item) => item.id === eventDraft.ruleId) ?? RULES[0];
    setEventDraft((current) => ({
      ...current,
      eventDate,
      dueDate: dueForRule(rule, eventDate, organization),
    }));
  }

  function saveEvent(event: FormEvent) {
    event.preventDefault();
    if (!eventDraft.employeeId || !eventDraft.eventDate || !eventDraft.dueDate) {
      setToast("Укажите сотрудника, дату события и срок.");
      return;
    }
    const notice: Notice = {
      id: makeId(),
      ...eventDraft,
      createdAt: new Date().toISOString(),
      completedAt: "",
      outgoingNumber: "",
    };
    setNotices((current) => [notice, ...current]);
    const eventEmployee = employees.find((employee) => employee.id === eventDraft.employeeId);
    if (eventDraft.ruleId === "hire") {
      setEmployees((current) =>
        current.map((employee) =>
          employee.id === eventDraft.employeeId
            ? { ...employee, active: true, hireDate: eventDraft.eventDate }
            : employee,
        ),
      );
      if (eventEmployee) {
        const changes: EmployeeFieldChange[] = [];
        if (!eventEmployee.active) changes.push({ key: "active", label: "статус", oldValue: false, newValue: true });
        if (eventEmployee.hireDate !== eventDraft.eventDate) changes.push({ key: "hireDate", label: "дата приёма", oldValue: eventEmployee.hireDate, newValue: eventDraft.eventDate });
        if (changes.length) setEmployeeChanges((current) => [{ id: makeId(), employeeId: eventEmployee.id, employeeName: eventEmployee.fullName, createdAt: new Date().toISOString(), changes }, ...current]);
      }
    }
    if (eventDraft.ruleId === "dismissal") {
      setEmployees((current) =>
        current.map((employee) =>
          employee.id === eventDraft.employeeId
            ? { ...employee, active: false, dismissalDate: eventDraft.eventDate }
            : employee,
        ),
      );
      if (eventEmployee) {
        const changes: EmployeeFieldChange[] = [];
        if (eventEmployee.active) changes.push({ key: "active", label: "статус", oldValue: true, newValue: false });
        if (eventEmployee.dismissalDate !== eventDraft.eventDate) changes.push({ key: "dismissalDate", label: "дата увольнения", oldValue: eventEmployee.dismissalDate, newValue: eventDraft.eventDate });
        if (changes.length) setEmployeeChanges((current) => [{ id: makeId(), employeeId: eventEmployee.id, employeeName: eventEmployee.fullName, createdAt: new Date().toISOString(), changes }, ...current]);
      }
    }
    setEventModal(false);
    setToast("Задача создана, срок рассчитан.");
  }

  function completeNotice(notice: Notice) {
    const number = window.prompt("Исходящий номер или примечание об исполнении:", notice.outgoingNumber);
    if (number === null) return;
    setNotices((current) =>
      current.map((item) =>
        item.id === notice.id
          ? { ...item, completedAt: new Date().toISOString(), outgoingNumber: number }
          : item,
      ),
    );
    setToast("Задача отмечена исполненной.");
  }

  function reopenNotice(notice: Notice) {
    setNotices((current) =>
      current.map((item) =>
        item.id === notice.id ? { ...item, completedAt: "", outgoingNumber: "" } : item,
      ),
    );
  }

  function prepareDocumentForNotice(notice: Notice) {
    const employee = employees.find((item) => item.id === notice.employeeId);
    if (!employee) {
      setToast("Карточка сотрудника не найдена.");
      return;
    }
    if (notice.ruleId === "hire" || notice.ruleId === "dismissal") {
      setDocumentType("f2");
      setDocumentEmployeeId(employee.id);
      setF2Event(notice.ruleId === "hire" ? "hire" : "dismissal");
      setF2OrderNumber(employee.orderNumber);
      setF2OrderDate(employee.orderDate || notice.eventDate);
    } else if (notice.ruleId === "change") {
      setDocumentType("changes");
      setDocumentEmployeeId(employee.id);
      setChangeDocumentEntries([{ employeeId: employee.id, content: notice.note }]);
    } else {
      setDocumentEmployeeId(employee.id);
    }
    navigate("documents");
    window.scrollTo({ top: 0, behavior: "smooth" });
    setToast("Параметры документа заполнены из задачи.");
  }

  function matchingNoticeIds(type: DocumentType, employeeIds: string[]) {
    const ruleId = type === "f2" || type === "employmentNotice"
      ? (f2Event === "hire" ? "hire" : "dismissal")
      : type === "changes"
        ? "change"
        : "";
    if (!ruleId) return [];
    return notices
      .filter((notice) => !notice.completedAt && notice.ruleId === ruleId && employeeIds.includes(notice.employeeId))
      .map((notice) => notice.id);
  }

  function recordDocument(employee: Employee, type: DocumentType, employeeIds = [employee.id], entries?: ChangeDocumentEntry[]) {
    const id = makeId();
    const noticeIds = matchingNoticeIds(type, employeeIds);
    setDocuments((current) => [
      {
        id,
        employeeId: employee.id,
        type,
        createdAt: new Date().toISOString(),
        status: "formed",
        headerLocation: (["f2", "employmentNotice", "changes", "officerList", "enlistedList"] as DocumentType[]).includes(type) ? documentHeaderLocation : undefined,
        eventType: type === "f2" || type === "employmentNotice" ? f2Event : undefined,
        orderNumber: type === "f2" || type === "employmentNotice" ? f2OrderNumber : undefined,
        orderDate: type === "f2" || type === "employmentNotice" ? f2OrderDate : undefined,
        employeeIds,
        employeeCount: employeeIds.length,
        changeEntries: entries?.map((entry) => ({ ...entry })),
        outgoingNumber: "",
        sentAt: "",
        noticeIds,
        title: type === "form10"
          ? "Форма № 10"
          : type === "f2"
            ? `Справка Ф-2 — ${f2Event === "hire" ? "приём" : "увольнение"}`
            : type === "messageSheet"
              ? "Листок сообщений"
              : type === "changes"
                ? "Сведения об изменениях"
                : type === "officerList"
                  ? "Список — офицеры"
                  : type === "enlistedList"
                    ? "Список — рядовые и сержанты"
                : `Сведения о принятых/уволенных — ${f2Event === "hire" ? "приём" : "увольнение"}`,
      },
      ...current,
    ]);
  }

  function setDocumentRecordStatus(record: DocumentRecord, status: "formed" | "signed" | "sent") {
    if (status === "sent") {
      const outgoingNumber = window.prompt("Исходящий номер:", record.outgoingNumber ?? "");
      if (outgoingNumber === null) return;
      const sentAt = window.prompt("Дата отправки (ГГГГ-ММ-ДД):", record.sentAt || todayIso());
      if (!sentAt) return;
      setDocuments((current) => current.map((item) => item.id === record.id ? { ...item, status, outgoingNumber, sentAt } : item));
      const linkedIds = new Set(record.noticeIds ?? []);
      if (linkedIds.size) {
        setNotices((current) => current.map((notice) => linkedIds.has(notice.id)
          ? { ...notice, completedAt: `${sentAt}T12:00:00`, outgoingNumber, completedByDocumentId: record.id }
          : notice));
      }
      setToast(linkedIds.size
        ? `Документ отправлен. Связанных задач закрыто: ${linkedIds.size}.`
        : "Документ отмечен отправленным.");
      return;
    }
    if ((record.status ?? "formed") === "sent") {
      setNotices((current) => current.map((notice) => notice.completedByDocumentId === record.id
        ? { ...notice, completedAt: "", outgoingNumber: "", completedByDocumentId: undefined }
        : notice));
    }
    setDocuments((current) => current.map((item) => item.id === record.id ? { ...item, status } : item));
    setToast(status === "signed" ? "Документ отмечен подписанным." : "Статус возвращён: сформирован.");
  }

  function repeatDocument(record: DocumentRecord) {
    setDocumentType(record.type);
    setDocumentHeaderLocation(record.headerLocation ?? "moscow");
    setF2Event(record.eventType ?? "hire");
    setF2OrderNumber(record.orderNumber ?? "");
    setF2OrderDate(record.orderDate ?? "");
    if (record.type === "changes") {
      setChangeDocumentEntries((record.changeEntries ?? []).filter((entry) => employees.some((employee) => employee.id === entry.employeeId)));
    } else if (record.type !== "officerList" && record.type !== "enlistedList") {
      setDocumentEmployeeId(record.employeeId);
    }
    setView("documents");
    window.scrollTo({ top: 0, behavior: "smooth" });
    setToast("Параметры документа восстановлены. Проверьте данные и скачайте Word.");
  }

  async function downloadWord() {
    if (documentType === "officerList" || documentType === "enlistedList") {
      if (!automaticListEntries.length) {
        setToast("В списке нет действующих сотрудников с подходящими воинскими званиями.");
        return;
      }
      if ((listEmployeesWithGaps.length || documentSettingsMissing.length) && !window.confirm(
        `Документ содержит пропуски.\n\nСотрудников с незаполненными графами: ${listEmployeesWithGaps.length}.${documentSettingsMissing.length ? `\nНастройки: ${documentSettingsMissing.join(", ")}.` : ""}\n\nСкачать Word с пропусками?`,
      )) return;
      try {
        setToast("Формируется список сотрудников…");
        const blob = await buildPersonnelListDocx(
          documentType,
          automaticListEntries.map((entry) => entry.employeeId),
          employees,
          organization,
          documentHeaderLocation,
        );
        downloadBlob(blob, documentType === "officerList" ? "Список_офицеры.docx" : "Список_рядовые.docx");
        const firstEmployee = employees.find((item) => item.id === automaticListEntries[0].employeeId);
        if (firstEmployee) recordDocument(firstEmployee, documentType, automaticListEntries.map((entry) => entry.employeeId));
        setToast("Редактируемый Word подготовлен.");
      } catch {
        setToast("Не удалось сформировать Word. Проверьте шаблон и повторите попытку.");
      }
      return;
    }
    if (documentType === "changes") {
      if (!changeDocumentEntries.length) {
        setToast("Добавьте хотя бы одного сотрудника.");
        return;
      }
      if (changeDocumentEntries.some((entry) => !entry.content.trim())) {
        setToast("Заполните содержание изменений для каждого сотрудника.");
        return;
      }
      if (documentSettingsMissing.length && !window.confirm(`В настройках не заполнено: ${documentSettingsMissing.join(", ")}.\n\nСкачать Word с пропусками?`)) return;
      try {
        setToast("Формируются сведения об изменениях…");
        const blob = await buildChangesDocx(changeDocumentEntries, employees, organization, documentHeaderLocation);
        downloadBlob(blob, "Сведения_об_изменениях.docx");
        const firstEmployee = employees.find((item) => item.id === changeDocumentEntries[0].employeeId);
        if (firstEmployee) recordDocument(firstEmployee, documentType, changeDocumentEntries.map((entry) => entry.employeeId), changeDocumentEntries);
        setToast("Редактируемый Word подготовлен.");
      } catch {
        setToast("Не удалось сформировать Word. Проверьте шаблон и повторите попытку.");
      }
      return;
    }
    const employee = employees.find((item) => item.id === documentEmployeeId);
    if (!employee) {
      setToast("Выберите сотрудника.");
      return;
    }
    const allDocumentGaps = [...selectedDocumentCheck.critical, ...documentParameterMissing, ...documentSettingsMissing];
    if (allDocumentGaps.length && !window.confirm(`Для выбранного документа не заполнено: ${allDocumentGaps.join(", ")}.\n\nСкачать Word с пропусками?`)) return;
    try {
      setToast("Формируется редактируемый Word…");
      const blob = await buildDocx(
        employee,
        organization,
        documentType,
        f2Event,
        f2OrderNumber,
        f2OrderDate,
        documentHeaderLocation,
      );
      downloadBlob(
        blob,
        `${documentType === "form10" ? "Форма-10" : documentType === "f2" ? "Справка-Ф-2" : documentType === "messageSheet" ? "Листок-сообщений" : "Сведения-о-принятых-уволенных"}_${employee.fullName.replaceAll(" ", "_")}.docx`,
      );
      recordDocument(employee, documentType);
      setToast("Редактируемый Word подготовлен.");
    } catch {
      setToast("Не удалось сформировать Word. Проверьте шаблоны и повторите попытку.");
    }
  }

  function updateVerification(employee: Employee, type: "employee" | "commissariat") {
    const promptTitle =
      type === "employee" ? "Дата сверки с документами сотрудника:" : "Дата сверки с военным комиссариатом:";
    const date = window.prompt(
      promptTitle,
      type === "employee"
        ? employee.lastEmployeeVerification || todayIso()
        : employee.lastCommissariatVerification || todayIso(),
    );
    if (!date) return;
    const key: keyof Employee = type === "employee" ? "lastEmployeeVerification" : "lastCommissariatVerification";
    const oldValue = employee[key] as string;
    setEmployees((current) =>
      current.map((item) =>
        item.id === employee.id
          ? {
              ...item,
              [type === "employee" ? "lastEmployeeVerification" : "lastCommissariatVerification"]:
                date,
            }
          : item,
      ),
    );
    if (oldValue !== date) setEmployeeChanges((current) => [{
      id: makeId(), employeeId: employee.id, employeeName: employee.fullName, createdAt: new Date().toISOString(),
      changes: [{ key, label: type === "employee" ? "сверка с документами сотрудника" : "сверка с военкоматом", oldValue, newValue: date }],
    }, ...current]);
    setToast("Дата сверки сохранена. Следующий срок рассчитан на год вперёд.");
  }

  function exportBackup() {
    const state: StoredState = { employees, notices, documents, employeeChanges, settings: organization };
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const link = window.document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `voinskiy-uchet-backup-${todayIso()}.json`;
    link.click();
    URL.revokeObjectURL(link.href);
    const createdAt = new Date().toISOString();
    localStorage.setItem(LAST_BACKUP_KEY, createdAt);
    setLastBackupAt(createdAt);
    setToast("Резервная копия создана.");
  }

  function exportWorkPlan() {
    if (!workPlanRows.length) {
      setToast("В выбранном месяце нет задач для выгрузки.");
      return;
    }
    const rows = workPlanRows.map(({ notice, employee, rule, document, status }, index) => ({
      "№": index + 1,
      "Срок": formatDate(notice.dueDate),
      "Сотрудник": employee?.fullName ?? "Карточка удалена",
      "Подразделение": employee?.department ?? "",
      "Задача": rule?.title ?? "Задача",
      "Дата события": formatDate(notice.eventDate),
      "Статус": statusLabel(status, notice.dueDate),
      "Документ": document?.title ?? "Не сформирован",
      "Статус документа": document ? (document.status === "sent" ? "Отправлен" : document.status === "signed" ? "Подписан" : "Сформирован") : "",
      "Дата отправки": formatDate(document?.sentAt ?? notice.completedAt, ""),
      "Исходящий №": document?.outgoingNumber || notice.outgoingNumber,
      "Примечание": notice.note,
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet["!cols"] = [5, 13, 32, 24, 42, 15, 18, 34, 20, 17, 18, 36].map((wch) => ({ wch }));
    worksheet["!autofilter"] = { ref: worksheet["!ref"] ?? "A1:L1" };
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Рабочий план");
    XLSX.writeFile(workbook, `Рабочий_план_${workPlanMonth}.xlsx`);
    setToast("Рабочий план выгружен в Excel.");
  }

  function exportControlReport() {
    const workbook = XLSX.utils.book_new();
    const summary = XLSX.utils.json_to_sheet([
      { "Показатель": "Версия программы", "Значение": APP_VERSION },
      { "Показатель": "Дата проверки", "Значение": formatDate(todayIso()) },
      { "Показатель": "Всего сотрудников", "Значение": employees.length },
      { "Показатель": "Работающих", "Значение": employees.filter((employee) => employee.active).length },
      { "Показатель": "Открытых задач", "Значение": pendingNotices.length },
      { "Показатель": "Просроченных задач", "Значение": overdueCount },
      { "Показатель": "Документов в истории", "Значение": documents.length },
      { "Показатель": "Ошибок целостности", "Значение": diagnosticErrorCount },
      { "Показатель": "Всего замечаний", "Значение": diagnosticIssues.length },
    ]);
    summary["!cols"] = [{ wch: 32 }, { wch: 28 }];
    XLSX.utils.book_append_sheet(workbook, summary, "Итоги");

    const gaps = XLSX.utils.json_to_sheet(incompleteActiveEmployees.map((employee) => ({
      "Сотрудник": employee.fullName,
      "Подразделение": employee.department,
      "Заполнено, %": cardCompleteness(employee),
      "Не заполнено": getMissingFields(employee).join(", "),
    })));
    gaps["!cols"] = [{ wch: 34 }, { wch: 25 }, { wch: 15 }, { wch: 70 }];
    XLSX.utils.book_append_sheet(workbook, gaps, "Неполные карточки");

    const tasks = XLSX.utils.json_to_sheet(pendingNotices.map((notice) => {
      const employee = employees.find((item) => item.id === notice.employeeId);
      const rule = RULES.find((item) => item.id === notice.ruleId);
      return {
        "Срок": formatDate(notice.dueDate),
        "Статус": statusLabel(noticeStatus(notice), notice.dueDate),
        "Сотрудник": employee?.fullName ?? "Карточка отсутствует",
        "Задача": rule?.title ?? notice.ruleId,
        "Примечание": notice.note,
      };
    }));
    tasks["!cols"] = [{ wch: 14 }, { wch: 18 }, { wch: 34 }, { wch: 48 }, { wch: 45 }];
    XLSX.utils.book_append_sheet(workbook, tasks, "Открытые задачи");

    const integrityRows = [
      ...duplicateEmployeeGroups.map((group) => ({ "Тип": "Возможный дубль", "Объект": group.map((employee) => employee.fullName).join("; "), "Описание": `Дата рождения: ${group[0].birthDate || "не указана"}` })),
      ...orphanNotices.map((notice) => ({ "Тип": "Задача без карточки", "Объект": notice.id, "Описание": RULES.find((rule) => rule.id === notice.ruleId)?.title ?? notice.ruleId })),
      ...orphanDocuments.map((record) => ({ "Тип": "Документ без карточки", "Объект": record.title, "Описание": formatDate(record.createdAt) })),
      ...employeesWithUnknownRank.map((employee) => ({ "Тип": "Не распознано звание", "Объект": employee.fullName, "Описание": employee.militaryRank || "не указано" })),
    ];
    const integrity = XLSX.utils.json_to_sheet(integrityRows.length ? integrityRows : [{ "Тип": "Ошибок не найдено", "Объект": "—", "Описание": "Связи и воинские звания проверены" }]);
    integrity["!cols"] = [{ wch: 28 }, { wch: 42 }, { wch: 55 }];
    XLSX.utils.book_append_sheet(workbook, integrity, "Диагностика");
    XLSX.writeFile(workbook, `Контрольный_отчет_${todayIso()}.xlsx`);
    setToast("Контрольный отчёт выгружен в Excel.");
  }

  async function restoreBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const data = JSON.parse(await file.text()) as StoredState;
      if (!Array.isArray(data.employees) || !Array.isArray(data.notices)) throw new Error();
      if (!window.confirm("Заменить текущие данные содержимым резервной копии?")) return;
      setEmployees(data.employees);
      setNotices(data.notices);
      setDocuments(Array.isArray(data.documents) ? data.documents : []);
      setEmployeeChanges(Array.isArray(data.employeeChanges) ? data.employeeChanges : []);
      setOrganization({ ...emptySettings, ...(data.settings ?? {}) });
      const restoredAt = new Date().toISOString();
      localStorage.setItem(LAST_BACKUP_KEY, restoredAt);
      setLastBackupAt(restoredAt);
      setToast("Резервная копия восстановлена.");
    } catch {
      setToast("Файл не является корректной резервной копией.");
    }
  }

  function clearAllData() {
    if (!window.confirm("Удалить все карточки, задачи, документы и настройки с этого устройства?")) return;
    if (!window.confirm("Действие необратимо без резервной копии. Продолжить?")) return;
    setEmployees([]);
    setNotices([]);
    setDocuments([]);
    setEmployeeChanges([]);
    setOrganization(emptySettings);
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(LAST_BACKUP_KEY);
    setLastBackupAt("");
    setToast("Все локальные данные удалены.");
  }

  const pageTitle =
    navItems.find((item) => item.id === view)?.label ?? "Воинский учёт";

  if (!hydrated) {
    return (
      <main className="loading-screen">
        <ShieldCheck size={34} />
        <p>Открываем рабочий контур…</p>
      </main>
    );
  }

  return (
    <div className="app-shell">
      <aside className={`sidebar ${menuOpen ? "is-open" : ""}`}>
        <button className="sidebar-close icon-button" onClick={() => setMenuOpen(false)} aria-label="Закрыть меню">
          <X size={20} />
        </button>
        <div className="brand">
          <div className="brand-mark">
            <ShieldCheck size={21} />
          </div>
          <div>
            <strong>ВОИНСКИЙ УЧЁТ</strong>
            <span>рабочий контур · версия {APP_VERSION}</span>
          </div>
        </div>
        <nav aria-label="Основная навигация">
          {navItems.map((item) => {
            const Icon = item.icon;
            return (
              <button
                key={item.id}
                className={view === item.id ? "active" : ""}
                onClick={() => navigate(item.id)}
              >
                <Icon size={21} />
                <span>{item.label}</span>
                {item.id === "notifications" && pendingNotices.length ? (
                  <b>{pendingNotices.length}</b>
                ) : null}
              </button>
            );
          })}
        </nav>
        <div className="sidebar-foot">
          <Database size={17} />
          <div>
            <strong>Данные на устройстве</strong>
            <span>Сохранение автоматическое</span>
          </div>
        </div>
      </aside>

      {menuOpen ? <button className="sidebar-scrim" aria-label="Закрыть меню" onClick={() => setMenuOpen(false)} /> : null}

      <main className="workspace">
        <header className="topbar">
          <div className="topbar-title">
            <button className="mobile-menu icon-button" onClick={() => setMenuOpen(true)} aria-label="Открыть меню">
              <Menu size={22} />
            </button>
            <div>
              <span>{formatLongDate()}</span>
              <h1>{pageTitle}</h1>
            </div>
          </div>
          <div className="topbar-actions">
            <input
              ref={excelInput}
              hidden
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={importExcel}
            />
            <button className="button secondary" onClick={() => excelInput.current?.click()}>
              <Upload size={18} /> Импорт Excel
            </button>
            <button className="button primary" onClick={openNewEmployee}>
              <Plus size={18} /> Добавить сотрудника
            </button>
          </div>
        </header>

        <div className="content">
          {view === "dashboard" ? (
            <>
              {backupIsDue ? <section className="backup-reminder inline-warning">
                <ArchiveRestore size={20}/>
                <span><strong>{lastBackupAt ? "Пора обновить резервную копию" : "Резервная копия ещё не создана"}</strong>{lastBackupAt ? `Последняя копия создана ${backupAgeDays} дн. назад. Рекомендуемый интервал — не более 7 дней.` : "Создайте первый файл восстановления, чтобы не потерять локальную базу браузера."}</span>
                <button className="button secondary small" onClick={exportBackup}><Download size={16}/> Создать копию</button>
              </section> : null}
              <section className="kpi-grid" aria-label="Основные показатели">
                <button className="kpi-card" onClick={() => navigate("employees")}>
                  <span className="kpi-icon teal"><Users size={25} /></span>
                  <span><small>Всего сотрудников</small><strong>{employees.length}</strong></span>
                  <ChevronRight size={18} />
                </button>
                <button className="kpi-card" onClick={() => navigate("notifications")}>
                  <span className="kpi-icon amber"><CircleAlert size={25} /></span>
                  <span><small>Требует внимания</small><strong>{attentionCount}</strong></span>
                  <ChevronRight size={18} />
                </button>
                <button className="kpi-card" onClick={() => { setNoticeFilter("today"); navigate("notifications"); }}>
                  <span className="kpi-icon blue"><CalendarCheck size={25} /></span>
                  <span><small>Срок сегодня</small><strong>{todayCount}</strong></span>
                  <ChevronRight size={18} />
                </button>
                <button className="kpi-card" onClick={() => { setNoticeFilter("overdue"); navigate("notifications"); }}>
                  <span className="kpi-icon red"><Clock3 size={25} /></span>
                  <span><small>Просрочено</small><strong>{overdueCount}</strong></span>
                  <ChevronRight size={18} />
                </button>
              </section>

              <section className="section-heading">
                <div>
                  <span className="eyebrow">Контроль сроков</span>
                  <h2>Требует внимания</h2>
                </div>
                <div className="heading-actions">
                  <button className="button secondary small" onClick={() => openEventModal()}>
                    <Plus size={16} /> Создать задачу
                  </button>
                  <button className="button ghost small" onClick={() => navigate("notifications")}>
                    Показать все <ChevronRight size={16} />
                  </button>
                </div>
              </section>

              {dashboardItems.length ? (
                <div className="data-panel task-list">
                  <div className="task-head">
                    <span>Сотрудник</span><span>Задача</span><span>Срок</span><span>Статус</span>
                  </div>
                  {dashboardItems.map((item) => (
                    <button
                      className="task-row"
                      key={item.id}
                      onClick={() =>
                        item.type === "missing" && item.employee
                          ? setEmployeeModal(item.employee)
                          : navigate("notifications")
                      }
                    >
                      <span className="employee-cell">
                        <span className="avatar"><UserRound size={18} /></span>
                        <span><strong>{item.employee?.fullName ?? "Общая задача"}</strong><small>{item.employee?.department || "Подразделение не указано"}</small></span>
                      </span>
                      <span>{item.title}</span>
                      <span>{item.dueDate ? formatDate(item.dueDate) : "Без срока"}</span>
                      <span><b className={`status ${item.status}`}>{item.status === "missing" ? "Неполная карточка" : statusLabel(item.status, item.dueDate)}</b></span>
                    </button>
                  ))}
                </div>
              ) : (
                <div className="empty-state">
                  <ShieldCheck size={36} />
                  <h3>{employees.length ? "Срочных задач нет" : "Начните с импорта реестра"}</h3>
                  <p>{employees.length ? "Карточки заполнены, открытые законодательные сроки отсутствуют." : "Загрузите исходный Excel — файл будет обработан только в вашем браузере."}</p>
                  <div>
                    <button className="button primary" onClick={() => excelInput.current?.click()}><FileSpreadsheet size={18} /> Импортировать Excel</button>
                    <button className="button secondary" onClick={openNewEmployee}><Plus size={18} /> Добавить вручную</button>
                  </div>
                </div>
              )}
            </>
          ) : null}

          {view === "employees" ? (
            <>
              <section className="section-heading compact">
                <div><span className="eyebrow">Личный состав</span><h2>Реестр сотрудников</h2><p>Показано {filteredEmployees.length} из {employees.length}</p></div>
                <button className="button secondary small" onClick={resetEmployeeFilters}><RotateCcw size={15} /> Сбросить фильтры</button>
              </section>
              <div className="filter-panel">
                <label className="search-box"><Search size={18} /><input value={employeeSearch} onChange={(e) => setEmployeeSearch(e.target.value)} placeholder="Ф.И.О., СНИЛС, паспорт или ВУС" /></label>
                <label><span>Статус</span><select value={activeFilter} onChange={(e) => setActiveFilter(e.target.value)}><option value="active">Работающие</option><option value="dismissed">Уволенные</option><option value="all">Все</option></select></label>
                <label><span>Подразделение</span><select value={departmentFilter} onChange={(e) => setDepartmentFilter(e.target.value)}><option value="">Все</option>{uniqueValues(employees, "department").map((value) => <option key={value}>{value}</option>)}</select></label>
                <label><span>Карточка</span><select value={attentionFilter} onChange={(e) => setAttentionFilter(e.target.value)}><option value="">Все</option><option value="missing">Есть пропуски</option></select></label>
              </div>
              {filteredEmployees.length ? (
                <div className="data-panel table-scroll">
                  <table className="registry-table">
                    <thead><tr>
                      <th><button onClick={() => toggleSort("fullName")}>Сотрудник <ChevronDown size={14} /></button></th>
                      <th><button onClick={() => toggleSort("department")}>Подразделение <ChevronDown size={14} /></button></th>
                      <th><button onClick={() => toggleSort("position")}>Должность <ChevronDown size={14} /></button></th>
                      <th><button onClick={() => toggleSort("militaryRank")}>Звание <ChevronDown size={14} /></button></th>
                      <th><button onClick={() => toggleSort("vus")}>ВУС <ChevronDown size={14} /></button></th>
                      <th><button onClick={() => toggleSort("fitnessCategory")}>Годность <ChevronDown size={14} /></button></th>
                      <th>Карточка</th><th></th>
                    </tr></thead>
                    <tbody>
                      {filteredEmployees.map((employee) => {
                        const missing = getMissingFields(employee);
                        return <tr key={employee.id} onDoubleClick={() => setEmployeeModal(employee)}>
                          <td><button className="employee-link" onClick={() => setEmployeeModal(employee)}><span className="avatar"><UserRound size={17} /></span><span><strong>{employee.fullName}</strong><small>{employee.snils || "СНИЛС не указан"}</small></span></button></td>
                          <td>{employee.department || "—"}</td><td>{employee.position || "—"}</td><td>{employee.militaryRank || "—"}</td><td className="mono">{employee.vus || "—"}</td><td>{employee.fitnessCategory || "—"}</td>
                          <td><span className={`status ${missing.length ? "missing" : "completed"}`}>{missing.length ? `Заполнено ${cardCompleteness(employee)}% · пропусков: ${missing.length}` : "Заполнено 100%"}</span></td>
                          <td><button className="icon-button" title="Редактировать" onClick={() => setEmployeeModal(employee)}><Pencil size={17} /></button></td>
                        </tr>;
                      })}
                    </tbody>
                  </table>
                </div>
              ) : <div className="empty-state compact-empty"><Users size={32} /><h3>Сотрудники не найдены</h3><p>Измените фильтры или импортируйте Excel.</p></div>}
            </>
          ) : null}

          {view === "notifications" ? (
            <>
              <section className="section-heading compact">
                <div><span className="eyebrow">Законодательные сроки</span><h2>Уведомления и задачи</h2><p>Срок рассчитывается от даты фактического события.</p></div>
                <button className="button primary small" onClick={() => openEventModal()}><Plus size={16} /> Создать задачу</button>
              </section>
              <div className="filter-panel notice-filters">
                <label className="search-box"><Search size={18} /><input value={noticeSearch} onChange={(e) => setNoticeSearch(e.target.value)} placeholder="Сотрудник или задача" /></label>
                <label><span>Статус</span><select value={noticeFilter} onChange={(e) => setNoticeFilter(e.target.value)}><option value="open">Все открытые</option><option value="today">Срок сегодня</option><option value="overdue">Просроченные</option><option value="upcoming">Предстоящие</option><option value="completed">Исполненные</option><option value="all">Все</option></select></label>
              </div>
              {filteredNotices.length ? <div className="notice-grid">
                {filteredNotices.map((notice) => {
                  const employee = employees.find((item) => item.id === notice.employeeId);
                  const rule = RULES.find((item) => item.id === notice.ruleId);
                  const status = noticeStatus(notice);
                  return <article className={`notice-card ${status}`} key={notice.id}>
                    <div className="notice-top"><span className={`status ${status}`}>{statusLabel(status, notice.dueDate)}</span><span>{formatDate(notice.dueDate)}</span></div>
                    <h3>{rule?.title}</h3><p className="notice-person">{employee?.fullName ?? "Сотрудник удалён"}</p>
                    {notice.note ? <p>{notice.note}</p> : null}
                    <div className="notice-meta"><span>Событие: {formatDate(notice.eventDate)}</span><a href={rule?.sourceUrl} target="_blank" rel="noreferrer"><BookOpen size={14} /> {rule?.source}</a></div>
                    <div className="notice-help"><Info size={16} /><span>{rule?.help}<small>{rule?.documentHint}</small></span></div>
                    <div className="card-actions">
                      {status === "completed" ? <button className="button ghost small" onClick={() => reopenNotice(notice)}><RotateCcw size={15} /> Вернуть в работу</button> : <button className="button primary small" onClick={() => completeNotice(notice)}><Check size={15} /> Исполнено</button>}
                      {employee ? <button className="button secondary small" onClick={() => setEmployeeModal(employee)}><UserRound size={15} /> Карточка</button> : null}
                    </div>
                  </article>;
                })}
              </div> : <div className="empty-state compact-empty"><Bell size={32} /><h3>Задач по выбранному фильтру нет</h3><p>Создайте событие — система рассчитает срок и покажет правовое основание.</p></div>}
            </>
          ) : null}

          {view === "workPlan" ? (
            <>
              <section className="section-heading compact">
                <div><span className="eyebrow">Организация работы</span><h2>Рабочий план</h2><p>Задачи автоматически собраны по установленным срокам исполнения.</p></div>
                <button className="button primary small" onClick={exportWorkPlan}><FileSpreadsheet size={17}/> Выгрузить Excel</button>
              </section>
              <div className="filter-panel work-plan-filters">
                <label><span>Месяц</span><input type="month" value={workPlanMonth} onChange={(event)=>setWorkPlanMonth(event.target.value)}/></label>
                <label><span>Исполнение</span><select value={workPlanStatus} onChange={(event)=>setWorkPlanStatus(event.target.value)}><option value="all">Все задачи</option><option value="open">Открытые</option><option value="completed">Исполненные</option></select></label>
              </div>
              <section className="work-plan-summary" aria-label="Итоги рабочего плана">
                <div><span>Всего</span><strong>{workPlanSummary.total}</strong></div>
                <div><span>Открыто</span><strong>{workPlanSummary.open}</strong></div>
                <div className={workPlanSummary.overdue ? "danger" : ""}><span>Просрочено</span><strong>{workPlanSummary.overdue}</strong></div>
                <div><span>Исполнено</span><strong>{workPlanSummary.completed}</strong></div>
              </section>
              {workPlanRows.length ? <div className="data-panel table-scroll">
                <table className="registry-table work-plan-table">
                  <thead><tr><th>Срок</th><th>Сотрудник</th><th>Задача</th><th>Статус</th><th>Документ</th><th>Исходящий №</th><th></th></tr></thead>
                  <tbody>{workPlanRows.map(({notice,employee,rule,document,status})=><tr key={notice.id}>
                    <td className="mono">{formatDate(notice.dueDate)}</td>
                    <td><button className="employee-link" disabled={!employee} onClick={()=>employee&&setEmployeeModal(employee)}><span><strong>{employee?.fullName??"Карточка удалена"}</strong><small>{employee?.department||"Подразделение не указано"}</small></span></button></td>
                    <td>{rule?.title??"Задача"}</td>
                    <td><span className={`status ${status}`}>{statusLabel(status,notice.dueDate)}</span></td>
                    <td><span className="work-plan-document"><strong>{document?.title??"Не сформирован"}</strong>{document?<small>{document.status==="sent"?"Отправлен":document.status==="signed"?"Подписан":"Сформирован"}</small>:["hire","dismissal","change"].includes(notice.ruleId)?<button className="inline-link" onClick={()=>prepareDocumentForNotice(notice)}>Подготовить документ</button>:null}</span></td>
                    <td>{document?.outgoingNumber||notice.outgoingNumber||"—"}</td>
                    <td>{status==="completed"?<button className="button ghost small" onClick={()=>reopenNotice(notice)}><RotateCcw size={14}/> Вернуть</button>:<button className="button secondary small" onClick={()=>completeNotice(notice)}><Check size={14}/> Исполнено</button>}</td>
                  </tr>)}</tbody>
                </table>
              </div>:<div className="empty-state compact-empty"><CalendarCheck size={32}/><h3>В выбранном месяце задач нет</h3><p>Задачи появятся здесь автоматически после создания события.</p></div>}
            </>
          ) : null}

          {view === "documents" ? (
            <>
              <section className="section-heading compact"><div><span className="eyebrow">Редактируемый Word</span><h2>Формирование документов</h2><p>Форма № 10, справка Ф-2 и «Листок сообщений» выгружаются заполненными в формате DOCX.</p></div></section>
              <div className="document-layout">
                <section className="data-panel document-builder">
                  <h3>Параметры документа</h3>
                  {!(["changes", "officerList", "enlistedList"] as DocumentType[]).includes(documentType) ? <label className="field"><span>Сотрудник</span><select value={documentEmployeeId} onChange={(e) => setDocumentEmployeeId(e.target.value)}><option value="">Выберите сотрудника</option>{employees.slice().sort((a,b)=>a.fullName.localeCompare(b.fullName,"ru")).map((employee)=><option value={employee.id} key={employee.id}>{employee.fullName}</option>)}</select></label> : null}
                  <div className="segmented document-types"><button className={documentType === "form10" ? "active" : ""} onClick={() => setDocumentType("form10")}>Форма № 10</button><button className={documentType === "f2" ? "active" : ""} onClick={() => setDocumentType("f2")}>Справка Ф-2</button><button className={documentType === "messageSheet" ? "active" : ""} onClick={() => setDocumentType("messageSheet")}>Листок сообщений</button><button className={documentType === "changes" ? "active" : ""} onClick={() => setDocumentType("changes")}>Изменение данных</button><button className={documentType === "employmentNotice" ? "active" : ""} onClick={() => setDocumentType("employmentNotice")}>Принятие / увольнение</button><button className={documentType === "officerList" ? "active" : ""} onClick={() => setDocumentType("officerList")}>Список: офицеры</button><button className={documentType === "enlistedList" ? "active" : ""} onClick={() => setDocumentType("enlistedList")}>Список: рядовые</button></div>
                  {documentType === "changes" ? <div className="change-document-options">
                    <label className="field"><span>Шапка документа</span><select value={documentHeaderLocation} onChange={(event) => setDocumentHeaderLocation(event.target.value as DocumentHeaderLocation)}><option value="moscow">Москва</option><option value="sochi">Сочи</option></select></label>
                    <div className="change-employee-add"><label className="field"><span>Добавить сотрудника</span><select value={changeEmployeeToAdd} onChange={(event)=>setChangeEmployeeToAdd(event.target.value)}><option value="">Выберите сотрудника</option>{employees.filter((employee)=>!changeDocumentEntries.some((entry)=>entry.employeeId===employee.id)).slice().sort((a,b)=>a.fullName.localeCompare(b.fullName,"ru")).map((employee)=><option value={employee.id} key={employee.id}>{employee.fullName}</option>)}</select></label><button type="button" className="button secondary" disabled={!changeEmployeeToAdd} onClick={()=>{setChangeDocumentEntries([...changeDocumentEntries,{employeeId:changeEmployeeToAdd,content:""}]);setChangeEmployeeToAdd("");}}><Plus size={16}/> Добавить</button></div>
                    <div className="change-entry-list">{changeDocumentEntries.map((entry)=>{const employee=employees.find((item)=>item.id===entry.employeeId);return <div className="change-entry" key={entry.employeeId}><div><strong>{employee?.fullName}</strong><button type="button" className="icon-button" aria-label="Удалить сотрудника из документа" onClick={()=>setChangeDocumentEntries(changeDocumentEntries.filter((item)=>item.employeeId!==entry.employeeId))}><Trash2 size={16}/></button></div>{documentType === "changes" ? <label className="field"><span>Содержание изменений</span><textarea value={entry.content} placeholder="Например: изменение места пребывания…" onChange={(event)=>setChangeDocumentEntries(changeDocumentEntries.map((item)=>item.employeeId===entry.employeeId?{...item,content:event.target.value}:item))}/></label> : null}</div>})}</div>
                  </div> : null}
                  {documentType === "officerList" || documentType === "enlistedList" ? <div className="change-document-options">
                    <label className="field"><span>Шапка документа</span><select value={documentHeaderLocation} onChange={(event) => setDocumentHeaderLocation(event.target.value as DocumentHeaderLocation)}><option value="moscow">Москва</option><option value="sochi">Сочи</option></select></label>
                    <div className="inline-success"><Check size={18}/><span><strong>Список сформирован автоматически</strong>Включены только действующие сотрудники с соответствующими воинскими званиями: {automaticListEntries.length}.</span></div>
                    <div className="document-check-summary"><strong>По званиям</strong><div>{Object.entries(listRankDistribution).map(([rank, count]) => <span className="status completed" key={rank}>{rank}: {count}</span>)}</div></div>
                    {listEmployeesWithGaps.length ? <div className="inline-warning"><AlertCircle size={18}/><span><strong>Незаполненные графы: {listEmployeesWithGaps.length} сотрудник(а)</strong>{listEmployeesWithGaps.slice(0, 4).map(({employee, check}) => `${employee.fullName}: ${check.critical.join(", ")}`).join("; ")}{listEmployeesWithGaps.length > 4 ? ` и ещё ${listEmployeesWithGaps.length - 4}` : ""}.</span></div> : null}
                    {employeesWithUnknownRank.length ? <div className="inline-warning"><AlertCircle size={18}/><span><strong>Не включены: {employeesWithUnknownRank.length}</strong>Не распознано звание: {employeesWithUnknownRank.slice(0, 4).map((employee) => `${employee.fullName} — ${employee.militaryRank || "не указано"}`).join("; ")}{employeesWithUnknownRank.length > 4 ? ` и ещё ${employeesWithUnknownRank.length - 4}` : ""}.</span></div> : null}
                    <div className="change-entry-list">{automaticListEntries.map((entry)=>{const employee=employees.find((item)=>item.id===entry.employeeId);return <div className="change-entry" key={entry.employeeId}><div><strong>{employee?.fullName}</strong><span>{employee?.militaryRank}</span></div></div>})}</div>
                  </div> : null}
                  {documentType === "f2" || documentType === "employmentNotice" ? <div className="f2-options">
                    <label className="field"><span>Шапка документа</span><select value={documentHeaderLocation} onChange={(event) => setDocumentHeaderLocation(event.target.value as DocumentHeaderLocation)}><option value="moscow">Москва</option><option value="sochi">Сочи</option></select></label>
                    <fieldset className="event-checklist"><legend>Отметка в справке</legend>
                      <label><input type="radio" name="f2-event" checked={f2Event==="hire"} onChange={()=>setF2Event("hire")}/><span>Принят (поступил) на работу</span></label>
                      <label><input type="radio" name="f2-event" checked={f2Event==="dismissal"} onChange={()=>setF2Event("dismissal")}/><span>Уволен с работы (отчислен из образовательной организации)</span></label>
                      <small>Выбранный вариант останется обычным, второй будет зачёркнут.</small>
                    </fieldset>
                    <div className="f2-order-fields">
                      <Field required label="Номер приказа (трудового договора)" value={f2OrderNumber} onChange={setF2OrderNumber}/>
                      <Field required type="date" label="Дата приказа (трудового договора)" value={f2OrderDate} onChange={setF2OrderDate}/>
                    </div>
                  </div> : null}
                  {!(["changes", "officerList", "enlistedList"] as DocumentType[]).includes(documentType) && documentEmployeeId ? selectedDocumentCheck.critical.length || documentParameterMissing.length ? <div className="inline-warning"><AlertCircle size={18}/><span><strong>Критические пропуски для этого документа</strong>{[...selectedDocumentCheck.critical, ...documentParameterMissing].join(", ")}.</span></div> : <div className="inline-success"><Check size={18}/> Обязательные поля выбранного документа заполнены</div> : null}
                  {selectedDocumentCheck.warnings.length && !(["changes", "officerList", "enlistedList"] as DocumentType[]).includes(documentType) ? <div className="inline-warning"><Info size={18}/><span><strong>Рекомендуется дополнить</strong>{selectedDocumentCheck.warnings.join(", ")}.</span></div> : null}
                  {documentSettingsMissing.length ? <div className="inline-warning"><Settings size={18}/><span><strong>Проверьте настройки организации</strong>Не заполнено: {documentSettingsMissing.join(", ")}.</span></div> : null}
                  <div className="builder-actions">
                    <button className="button ghost" onClick={() => { const employee = employees.find((item) => item.id === documentEmployeeId); if (employee) setEmployeeModal(employee); }} disabled={(["changes", "officerList", "enlistedList"] as DocumentType[]).includes(documentType) || !documentEmployeeId}><Pencil size={17} /> Редактировать данные</button>
                    <button className="button primary" onClick={downloadWord} disabled={documentType === "changes" ? !changeDocumentEntries.length : documentType === "officerList" || documentType === "enlistedList" ? !automaticListEntries.length : !documentEmployeeId}><FileDown size={18} /> {hasCriticalDocumentGaps ? "Скачать с пропусками" : "Скачать Word"}</button>
                  </div>
                </section>
                <section className="document-preview-card">
                  <div className="paper-preview">
                    <span>{documentType === "form10" ? "Форма № 10" : documentType === "f2" || documentType === "employmentNotice" ? "СВЕДЕНИЯ" : documentType === "changes" ? "СВЕДЕНИЯ ОБ ИЗМЕНЕНИЯХ" : documentType === "officerList" ? "СПИСОК: ОФИЦЕРЫ" : documentType === "enlistedList" ? "СПИСОК: РЯДОВЫЕ" : "ЛИСТОК СООБЩЕНИЯ"}</span>
                    <strong>{documentType === "changes" ? `${changeDocumentEntries.length} сотрудник(а)` : documentType === "officerList" || documentType === "enlistedList" ? `${automaticListEntries.length} сотрудник(а)` : documentEmployeeId ? employees.find((item) => item.id === documentEmployeeId)?.fullName : "Выберите сотрудника"}</strong>
                    <div>{Array.from({ length: 10 }).map((_, index) => <i key={index} />)}</div>
                  </div>
                  <div className="preview-data">
                    <h4>Проверка перед выгрузкой</h4>
                    {(documentType === "f2" || documentType === "employmentNotice" || documentType === "changes" || documentType === "officerList" || documentType === "enlistedList") ? <div><span>Шапка</span><strong>{documentHeaderLocation === "moscow" ? "Москва" : "Сочи"}</strong></div> : null}
                    {selectedDocumentEmployee && !(["changes", "officerList", "enlistedList"] as DocumentType[]).includes(documentType) ? <>
                      <div><span>Сотрудник</span><strong>{selectedDocumentEmployee.fullName}</strong></div>
                      <div><span>Звание / ВУС</span><strong>{selectedDocumentEmployee.militaryRank || "—"} / {selectedDocumentEmployee.vus || "—"}</strong></div>
                      <div><span>Документ воинского учёта</span><strong>{selectedDocumentEmployee.militaryDocNumber || "—"}</strong></div>
                      <div><span>Военный комиссариат</span><strong>{selectedDocumentEmployee.militaryCommissariat || "—"}</strong></div>
                      {documentType === "f2" || documentType === "employmentNotice" ? <>
                        <div><span>Событие</span><strong>{f2Event === "hire" ? "Принятие" : "Увольнение"}</strong></div>
                        <div><span>Приказ</span><strong>№ {f2OrderNumber || "—"} от {formatDate(f2OrderDate, "—")}</strong></div>
                      </> : null}
                    </> : null}
                    {documentType === "officerList" || documentType === "enlistedList" ? <>
                      <div><span>Включено</span><strong>{automaticListEntries.length}</strong></div>
                      <div><span>С пропусками</span><strong>{listEmployeesWithGaps.length}</strong></div>
                      <div><span>Нераспознанные звания</span><strong>{employeesWithUnknownRank.length}</strong></div>
                    </> : null}
                    {documentType === "changes" ? <div><span>Включено сотрудников</span><strong>{changeDocumentEntries.length}</strong></div> : null}
                  </div>
                  <p>{documentType === "form10" ? "Две страницы A4: лицевая и оборотная стороны." : documentType === "f2" || documentType === "employmentNotice" ? "Одна страница A4: сведения о принятии или увольнении." : documentType === "officerList" || documentType === "enlistedList" ? "Список автоматически сформирован по воинским званиям действующих сотрудников." : documentType === "changes" ? "Групповой документ: одна строка на каждого выбранного сотрудника." : "Одна страница A4 с корешком по представленному образцу."}</p>
                </section>
              </div>
              {documents.length ? <section className="history-section">
                <div className="section-heading compact"><div><h3>История документов</h3><p>Контроль формирования, подписания и отправки.</p></div></div>
                <div className="filter-panel document-history-filters">
                  <label><span>Документ</span><select value={documentHistoryType} onChange={(event)=>setDocumentHistoryType(event.target.value as DocumentType | "")}><option value="">Все документы</option><option value="form10">Форма № 10</option><option value="f2">Справка Ф-2</option><option value="messageSheet">Листок сообщений</option><option value="changes">Изменение данных</option><option value="employmentNotice">Принятие / увольнение</option><option value="officerList">Список: офицеры</option><option value="enlistedList">Список: рядовые</option></select></label>
                  <label><span>Сотрудник</span><select value={documentHistoryEmployee} onChange={(event)=>setDocumentHistoryEmployee(event.target.value)}><option value="">Все сотрудники</option>{employees.slice().sort((a,b)=>a.fullName.localeCompare(b.fullName,"ru")).map((employee)=><option value={employee.id} key={employee.id}>{employee.fullName}</option>)}</select></label>
                  <label><span>Статус</span><select value={documentHistoryStatus} onChange={(event)=>setDocumentHistoryStatus(event.target.value)}><option value="">Все статусы</option><option value="formed">Сформирован</option><option value="signed">Подписан</option><option value="sent">Отправлен</option></select></label>
                </div>
                {filteredDocumentHistory.length ? <div className="document-history-list">{filteredDocumentHistory.map((record) => {
                  const status = record.status ?? "formed";
                  const person = employees.find((employee)=>employee.id===record.employeeId);
                  return <article className="data-panel document-history-card" key={record.id}>
                    <div className="document-history-main"><FileText size={20}/><span><strong>{record.title}</strong><small>{record.employeeCount && record.employeeCount > 1 ? `${record.employeeCount} сотрудник(а)` : person?.fullName ?? "Карточка удалена"}</small></span></div>
                    <div className="document-history-meta">
                      <span><b>Создан:</b> {new Intl.DateTimeFormat("ru-RU",{dateStyle:"short",timeStyle:"short"}).format(new Date(record.createdAt))}</span>
                      {record.headerLocation ? <span><b>Шапка:</b> {record.headerLocation === "moscow" ? "Москва" : "Сочи"}</span> : null}
                      {record.eventType ? <span><b>Событие:</b> {record.eventType === "hire" ? "Приём" : "Увольнение"}</span> : null}
                      {record.orderNumber || record.orderDate ? <span><b>Приказ:</b> № {record.orderNumber || "—"} от {formatDate(record.orderDate ?? "", "—")}</span> : null}
                      {record.noticeIds?.length ? <span><b>Связанные задачи:</b> {record.noticeIds.length}{status === "sent" ? " · закрыты автоматически" : " · закроются после отправки"}</span> : null}
                      {status === "sent" ? <span><b>Отправлен:</b> {formatDate(record.sentAt ?? "", "—")} · исх. № {record.outgoingNumber || "—"}</span> : null}
                    </div>
                    <div className="document-history-actions">
                      <span className={`document-status ${status}`}>{status === "formed" ? "Сформирован" : status === "signed" ? "Подписан" : "Отправлен"}</span>
                      <button className="button ghost small" onClick={()=>repeatDocument(record)}><RotateCcw size={15}/> Повторить</button>
                      {status === "formed" ? <button className="button secondary small" onClick={()=>setDocumentRecordStatus(record,"signed")}><Check size={15}/> Подписан</button> : status === "signed" ? <button className="button ghost small" onClick={()=>setDocumentRecordStatus(record,"formed")}><RotateCcw size={15}/> Вернуть к сформирован</button> : <button className="button ghost small" onClick={()=>setDocumentRecordStatus(record,"signed")}><RotateCcw size={15}/> Вернуть к подписан</button>}
                      {status !== "sent" ? <button className="button primary small" onClick={()=>setDocumentRecordStatus(record,"sent")}><FileDown size={15}/> Отправлен</button> : null}
                    </div>
                  </article>;
                })}</div> : <div className="empty-state compact-empty"><FileText size={30}/><h3>Документы не найдены</h3><p>Измените фильтры истории.</p></div>}
              </section> : null}
            </>
          ) : null}

          {view === "reconciliations" ? (
            <>
              <section className="section-heading compact"><div><span className="eyebrow">Ежегодный контроль</span><h2>Сверки</h2><p>Отдельно с документами гражданина и с военным комиссариатом.</p></div></section>
              <div className="filter-panel notice-filters"><label className="search-box"><Search size={18}/><input value={reconciliationSearch} onChange={(e)=>setReconciliationSearch(e.target.value)} placeholder="Поиск сотрудника"/></label><label><span>Состояние</span><select value={reconciliationFilter} onChange={(e)=>setReconciliationFilter(e.target.value)}><option value="attention">Требует внимания</option><option value="all">Все</option></select></label></div>
              <div className="reconciliation-grid">
                {employees.filter((employee)=>employee.active && employee.fullName.toLowerCase().includes(reconciliationSearch.toLowerCase())).filter((employee)=>{
                  if(reconciliationFilter==="all") return true;
                  const employeeDue = employee.lastEmployeeVerification ? daysFromToday(addCalendarDays(employee.lastEmployeeVerification,365)) : -1;
                  const vkDue = employee.lastCommissariatVerification ? daysFromToday(addCalendarDays(employee.lastCommissariatVerification,365)) : -1;
                  return employeeDue === null || employeeDue <= 30 || vkDue === null || vkDue <= 30;
                }).map((employee)=> {
                  const employeeNext = employee.lastEmployeeVerification ? addCalendarDays(employee.lastEmployeeVerification,365) : "";
                  const vkNext = employee.lastCommissariatVerification ? addCalendarDays(employee.lastCommissariatVerification,365) : "";
                  return <article className="reconciliation-card" key={employee.id}><div className="reconciliation-person"><span className="avatar"><UserRound size={18}/></span><span><strong>{employee.fullName}</strong><small>{employee.militaryCommissariat || "Военкомат не указан"}</small></span></div>
                    <div className="verification-block"><span>С документами сотрудника</span><strong>{employee.lastEmployeeVerification ? formatDate(employee.lastEmployeeVerification) : "Не проводилась"}</strong><small>Следующая: {employeeNext ? formatDate(employeeNext) : "требуется"}</small><button className="button secondary small" onClick={()=>updateVerification(employee,"employee")}><Check size={14}/> Отметить сверку</button></div>
                    <div className="verification-block"><span>С военным комиссариатом</span><strong>{employee.lastCommissariatVerification ? formatDate(employee.lastCommissariatVerification) : "Не проводилась"}</strong><small>Следующая: {vkNext ? formatDate(vkNext) : "требуется"}</small><button className="button secondary small" onClick={()=>updateVerification(employee,"commissariat")}><Check size={14}/> Отметить сверку</button></div>
                  </article>;
                })}
              </div>
              {!employees.length ? <div className="empty-state compact-empty"><ListChecks size={32}/><h3>Нет сотрудников для сверки</h3><p>Импортируйте реестр или добавьте карточку.</p></div> : null}
            </>
          ) : null}

          {view === "settings" ? (
            <>
              <section className="section-heading compact"><div><span className="eyebrow">Параметры рабочего контура</span><h2>Настройки и помощь</h2><p>Изменения сохраняются автоматически на этом устройстве.</p></div></section>
              <div className="settings-grid">
                <section className="data-panel settings-card">
                  <div className="card-title"><span className="kpi-icon teal"><Settings size={20}/></span><div><h3>Организация</h3><p>Используется при заполнении Word.</p></div></div>
                  <div className="form-grid two-col">
                    <Field label="Полное наименование" value={organization.organizationName} onChange={(value)=>setOrganization({...organization,organizationName:value})}/>
                    <Field label="Краткое наименование" value={organization.shortName} onChange={(value)=>setOrganization({...organization,shortName:value})}/>
                    <Field label="Адрес организации" value={organization.organizationAddress} onChange={(value)=>setOrganization({...organization,organizationAddress:value})}/>
                    <Field label="Должность руководителя" value={organization.directorPosition} onChange={(value)=>setOrganization({...organization,directorPosition:value})}/>
                    <Field label="Ф.И.О. руководителя" value={organization.directorName} onChange={(value)=>setOrganization({...organization,directorName:value})}/>
                    <Field label="Должность специалиста по воинскому учёту" value={organization.responsiblePosition} onChange={(value)=>setOrganization({...organization,responsiblePosition:value})}/>
                    <Field label="Ф.И.О. специалиста по воинскому учёту" value={organization.responsibleName} onChange={(value)=>setOrganization({...organization,responsibleName:value})}/>
                    <Field label="Телефон специалиста по воинскому учёту" value={organization.responsiblePhone} onChange={(value)=>setOrganization({...organization,responsiblePhone:value})}/>
                    <Field label="Военкомат по умолчанию" value={organization.defaultCommissariat} onChange={(value)=>setOrganization({...organization,defaultCommissariat:value})}/>
                    <Field label="Адрес военкомата" value={organization.defaultCommissariatAddress} onChange={(value)=>setOrganization({...organization,defaultCommissariatAddress:value})}/>
                    <Field label="Дополнительные нерабочие дни" value={organization.extraHolidays} placeholder="2026-04-20, 2026-09-01" help="Укажите даты через запятую. Они будут исключены при расчёте сроков в рабочих днях." onChange={(value)=>setOrganization({...organization,extraHolidays:value})}/>
                  </div>
                </section>
                <section className="data-panel settings-card">
                  <div className="card-title"><span className="kpi-icon blue"><Database size={20}/></span><div><h3>Локальные данные</h3><p>В сайт и на сервер персональные данные не отправляются.</p></div></div>
                  <div className="storage-stat"><strong>{employees.length}</strong><span>карточек сотрудников</span></div>
                  <div className="settings-actions"><button className="button primary" onClick={exportBackup}><Download size={17}/> Создать резервную копию</button><input ref={backupInput} hidden type="file" accept=".json" onChange={restoreBackup}/><button className="button secondary" onClick={()=>backupInput.current?.click()}><ArchiveRestore size={17}/> Восстановить</button><button className="button danger-outline" onClick={clearAllData}><Trash2 size={17}/> Очистить данные</button></div>
                </section>
              </div>
              <section className="data-panel diagnostics-card">
                <div className="diagnostics-heading">
                  <div className="card-title"><span className={`kpi-icon ${diagnosticErrorCount ? "red" : diagnosticIssues.length ? "amber" : "teal"}`}><ShieldCheck size={20}/></span><div><h3>Диагностика базы</h3><p>Проверка связей, дублей, полноты карточек, сроков и резервного копирования.</p></div></div>
                  <button className="button secondary small" onClick={exportControlReport}><FileSpreadsheet size={16}/> Контрольный отчёт</button>
                </div>
                <div className="diagnostics-summary"><span><strong>{diagnosticErrorCount}</strong> ошибок целостности</span><span><strong>{diagnosticIssues.length-diagnosticErrorCount}</strong> предупреждений</span><span><strong>{employees.length+notices.length+documents.length}</strong> записей проверено</span></div>
                {diagnosticIssues.length ? <div className="diagnostics-list">{diagnosticIssues.map((issue,index)=><div className={issue.level} key={`${issue.title}-${index}`}><span>{issue.level==="error"?<AlertCircle size={17}/>:<Info size={17}/>}<strong>{issue.title}</strong><small>{issue.detail}</small></span>{issue.view&&issue.view!=="settings"?<button className="button ghost small" onClick={()=>navigate(issue.view!)}>Открыть <ChevronRight size={14}/></button>:issue.title==="Требуется резервная копия"?<button className="button ghost small" onClick={exportBackup}>Создать копию <ChevronRight size={14}/></button>:null}</div>)}</div>:<div className="diagnostics-ok"><Check size={19}/><span><strong>Проверка пройдена</strong>Ошибок и предупреждений не найдено.</span></div>}
              </section>
              <section className="legal-section">
                <div className="section-heading compact"><div><span className="eyebrow">Актуально на 27.07.2026</span><h2>Правила уведомлений</h2><p>Нажмите на основание, чтобы открыть действующую редакцию.</p></div></div>
                <div className="legal-grid">{RULES.map((rule)=><a href={rule.sourceUrl} target="_blank" rel="noreferrer" key={rule.id}><span className="rule-days">{rule.days === null ? "!" : rule.days}<small>{rule.days === null ? "срочно" : rule.workingDays ? "раб. дн." : "дней"}</small></span><span><strong>{rule.shortTitle}</strong><small>{rule.source}</small><p>{rule.help}</p></span><ChevronRight size={17}/></a>)}</div>
              </section>
              <section className="quick-help data-panel"><div className="card-title"><span className="kpi-icon amber"><Info size={20}/></span><div><h3>Как начать работу</h3><p>Первая настройка занимает несколько минут.</p></div></div><ol><li><span>1</span><div><strong>Заполните данные организации</strong><p>Они автоматически подставятся в Форму № 10 и Ф-2.</p></div></li><li><span>2</span><div><strong>Импортируйте Excel</strong><p>Сначала программа покажет количество карточек и предупреждения.</p></div></li><li><span>3</span><div><strong>Исправьте неполные карточки</strong><p>Фильтр «Есть пропуски» найдёт сотрудников, которых нужно проверить.</p></div></li><li><span>4</span><div><strong>Создавайте события</strong><p>Приём, увольнение и изменения создают задачи с правовым основанием и сроком.</p></div></li></ol></section>
            </>
          ) : null}
        </div>
      </main>

      {employeeModal ? (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setEmployeeModal(null)}>
          <form className="modal employee-modal" onSubmit={saveEmployee}>
            <div className="modal-header"><div><span className="eyebrow">{employees.some((employee)=>employee.id===employeeModal.id) ? "Редактирование" : "Новая карточка"}</span><h2>{employeeModal.fullName || "Сотрудник"}</h2></div><button type="button" className="icon-button" onClick={()=>setEmployeeModal(null)} aria-label="Закрыть"><X size={20}/></button></div>
            <div className="modal-body">
              {getMissingFields(employeeModal).length ? <div className="inline-warning"><AlertCircle size={18}/><span><strong>Карточка заполнена на {cardCompleteness(employeeModal)}%</strong>Не заполнено: {getMissingFields(employeeModal).slice(0,6).join(", ")}{getMissingFields(employeeModal).length>6 ? ` и ещё ${getMissingFields(employeeModal).length-6}` : ""}.</span></div> : <div className="inline-success"><Check size={18}/> Карточка заполнена на 100%</div>}
              <fieldset><legend>Основные сведения</legend><div className="form-grid three-col">
                <Field required label="Ф.И.О." value={employeeModal.fullName} onChange={(value)=>setEmployeeModal({...employeeModal,fullName:value})}/>
                <label className="field"><span>Пол</span><select value={employeeModal.sex} onChange={(e)=>setEmployeeModal({...employeeModal,sex:e.target.value as "male"|"female"})}><option value="male">Мужской</option><option value="female">Женский</option></select></label>
                <Field required type="date" label="Дата рождения" value={employeeModal.birthDate} onChange={(value)=>setEmployeeModal({...employeeModal,birthDate:value})}/>
                <Field required label="Место рождения" value={employeeModal.birthPlace} onChange={(value)=>setEmployeeModal({...employeeModal,birthPlace:value})}/>
                <Field label="СНИЛС" value={employeeModal.snils} onChange={(value)=>setEmployeeModal({...employeeModal,snils:value})}/>
                <Field label="ИНН" value={employeeModal.inn} onChange={(value)=>setEmployeeModal({...employeeModal,inn:value})}/>
              </div></fieldset>
              <fieldset><legend>Работа</legend><div className="form-grid three-col">
                <Field label="Подразделение" value={employeeModal.department} onChange={(value)=>setEmployeeModal({...employeeModal,department:value})}/>
                <Field label="Должность" value={employeeModal.position} onChange={(value)=>setEmployeeModal({...employeeModal,position:value})}/>
                {employees.some((employee)=>employee.id===employeeModal.id) ? <label className="field"><span>Статус</span><select value={employeeModal.active ? "active":"dismissed"} onChange={(e)=>setEmployeeModal({...employeeModal,active:e.target.value==="active"})}><option value="active">Работает</option><option value="dismissed">Уволен</option></select></label> : null}
                <Field type="date" label="Дата приёма" value={employeeModal.hireDate} onChange={(value)=>setEmployeeModal({...employeeModal,hireDate:value})}/>
                <Field label="Номер приказа" value={employeeModal.orderNumber} onChange={(value)=>setEmployeeModal({...employeeModal,orderNumber:value})}/>
                <Field type="date" label="Дата приказа" value={employeeModal.orderDate} onChange={(value)=>setEmployeeModal({...employeeModal,orderDate:value})}/>
                {!employeeModal.active ? <Field type="date" label="Дата увольнения" value={employeeModal.dismissalDate} onChange={(value)=>setEmployeeModal({...employeeModal,dismissalDate:value})}/> : null}
              </div></fieldset>
              <fieldset><legend>Паспорт, адреса и контакты</legend><div className="form-grid three-col">
                <Field required label="Серия паспорта" value={employeeModal.passportSeries} onChange={(value)=>setEmployeeModal({...employeeModal,passportSeries:value})}/>
                <Field required label="Номер паспорта" value={employeeModal.passportNumber} onChange={(value)=>setEmployeeModal({...employeeModal,passportNumber:value})}/>
                <Field type="date" label="Дата выдачи" value={employeeModal.passportIssueDate} onChange={(value)=>setEmployeeModal({...employeeModal,passportIssueDate:value})}/>
                <Field label="Кем выдан" value={employeeModal.passportIssuedBy} onChange={(value)=>setEmployeeModal({...employeeModal,passportIssuedBy:value})}/>
                <Field required label="Адрес регистрации" value={employeeModal.registrationAddress} onChange={(value)=>setEmployeeModal({...employeeModal,registrationAddress:value})}/>
                <Field type="date" label="Дата регистрации" value={employeeModal.registrationDate} onChange={(value)=>setEmployeeModal({...employeeModal,registrationDate:value})}/>
                <Field label="Фактический адрес" value={employeeModal.actualAddress} onChange={(value)=>setEmployeeModal({...employeeModal,actualAddress:value})}/>
                <Field label="Сотовый телефон" value={employeeModal.phone} onChange={(value)=>setEmployeeModal({...employeeModal,phone:value})}/>
              </div></fieldset>
              <fieldset><legend>Образование и семья</legend><div className="form-grid three-col">
                <Field label="Образование" value={employeeModal.education} onChange={(value)=>setEmployeeModal({...employeeModal,education:value})}/>
                <Field label="Профессия" value={employeeModal.profession} onChange={(value)=>setEmployeeModal({...employeeModal,profession:value})}/>
                <Field label="Иностранные языки" value={employeeModal.languages} onChange={(value)=>setEmployeeModal({...employeeModal,languages:value})}/>
                <Field label="Водительское удостоверение" value={employeeModal.driverLicense} onChange={(value)=>setEmployeeModal({...employeeModal,driverLicense:value})}/>
                <label className="field"><span>Семейное положение</span><select value={employeeModal.maritalStatus} onChange={(event)=>setEmployeeModal({...employeeModal,maritalStatus:event.target.value})}><option value="">Не указано</option><option>Не состоит в браке (не замужем / не женат)</option><option>Состоит в зарегистрированном браке (женат / замужем)</option><option>Разведён (разведена)</option><option>Вдовец (вдова)</option></select></label>
              </div><FamilyMembersEditor value={employeeModal.familyMembers} onChange={(value)=>setEmployeeModal({...employeeModal,familyMembers:value})}/></fieldset>
              <fieldset><legend>Воинский учёт</legend><div className="form-grid three-col">
                <Field label="Вид документа" value={employeeModal.militaryDocType} onChange={(value)=>setEmployeeModal({...employeeModal,militaryDocType:value})}/>
                <Field required label="Серия и номер документа" value={employeeModal.militaryDocNumber} onChange={(value)=>setEmployeeModal({...employeeModal,militaryDocNumber:value})}/>
                <Field type="date" label="Дата выдачи документа" value={employeeModal.militaryDocIssueDate} onChange={(value)=>setEmployeeModal({...employeeModal,militaryDocIssueDate:value})}/>
                <Field label="Кем выдан документ" value={employeeModal.militaryDocIssuedBy} onChange={(value)=>setEmployeeModal({...employeeModal,militaryDocIssuedBy:value})}/>
                <Field required label="Воинское звание" value={employeeModal.militaryRank} onChange={(value)=>setEmployeeModal({...employeeModal,militaryRank:value})}/>
                <Field required label="Состав" value={employeeModal.composition} onChange={(value)=>setEmployeeModal({...employeeModal,composition:value})}/>
                <Field label="Профиль" value={employeeModal.profile} onChange={(value)=>setEmployeeModal({...employeeModal,profile:value})}/>
                <Field required label="ВУС" value={employeeModal.vus} help="Хранится как текст, чтобы не потерять начальные нули." onChange={(value)=>setEmployeeModal({...employeeModal,vus:value})}/>
                <label className="field"><span>Категория запаса *</span><select required value={employeeModal.reserveCategory} onChange={(event)=>setEmployeeModal({...employeeModal,reserveCategory:event.target.value})}><option value="">Не указано</option><option value="1">1</option><option value="2">2</option><option value="3">3</option></select></label>
                <Field required label="Категория годности" value={employeeModal.fitnessCategory} onChange={(value)=>setEmployeeModal({...employeeModal,fitnessCategory:value})}/>
                <Field label="Состояние здоровья" value={employeeModal.healthStatus ?? ""} onChange={(value)=>setEmployeeModal({...employeeModal,healthStatus:value})}/>
                <Field required label="Военный комиссариат" value={employeeModal.militaryCommissariat} onChange={(value)=>setEmployeeModal({...employeeModal,militaryCommissariat:value})}/>
                <Field label="Адрес военного комиссариата" value={employeeModal.militaryCommissariatAddress} onChange={(value)=>setEmployeeModal({...employeeModal,militaryCommissariatAddress:value})}/>
                <Field type="date" label="Сверка с документами сотрудника" value={employeeModal.lastEmployeeVerification} onChange={(value)=>setEmployeeModal({...employeeModal,lastEmployeeVerification:value})}/>
                <Field type="date" label="Сверка с военкоматом" value={employeeModal.lastCommissariatVerification} onChange={(value)=>setEmployeeModal({...employeeModal,lastCommissariatVerification:value})}/>
                <Field label="Дополнительные сведения" value={employeeModal.notes} onChange={(value)=>setEmployeeModal({...employeeModal,notes:value})}/>
              </div></fieldset>
              {employees.some((employee)=>employee.id===employeeModal.id) ? <fieldset className="employee-change-history"><legend>История изменений</legend>
                {employeeChanges.filter((record)=>record.employeeId===employeeModal.id).length ? <div className="employee-change-list">{employeeChanges.filter((record)=>record.employeeId===employeeModal.id).slice(0,10).map((record) => <article className={record.undoneAt ? "undone" : ""} key={record.id}>
                  <div className="employee-change-heading"><span><strong>{new Intl.DateTimeFormat("ru-RU",{dateStyle:"short",timeStyle:"short"}).format(new Date(record.createdAt))}</strong>{record.undoneAt ? <small>Изменение отменено</small> : <small>Изменено полей: {record.changes.length}</small>}</span><div><button type="button" className="button ghost small" disabled={Boolean(record.undoneAt)} onClick={()=>prepareChangeDocument(record)}><FileText size={14}/> В документ</button>{employeeChanges.find((item)=>item.employeeId===employeeModal.id&&!item.undoneAt)?.id === record.id ? <button type="button" className="button danger-outline small" onClick={()=>undoEmployeeChange(record)}><RotateCcw size={14}/> Отменить</button> : null}</div></div>
                  <div className="employee-change-fields">{record.changes.map((change)=><div key={String(change.key)}><strong>{change.label}</strong><span className="old-value">{displayChangeValue(change,change.oldValue)}</span><ChevronRight size={14}/><span className="new-value">{displayChangeValue(change,change.newValue)}</span></div>)}</div>
                </article>)}</div> : <p className="empty-history">Изменения этой карточки ещё не зафиксированы.</p>}
              </fieldset> : null}
            </div>
            <div className="modal-footer">{employees.some((employee)=>employee.id===employeeModal.id) ? <button type="button" className="button danger-outline" onClick={()=>deleteEmployee(employeeModal)}><Trash2 size={16}/> Удалить</button> : <span/>}<div><button type="button" className="button secondary" onClick={()=>setEmployeeModal(null)}>Отмена</button><button className="button primary" type="submit"><Check size={16}/> Сохранить</button></div></div>
          </form>
        </div>
      ) : null}

      {eventModal ? (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event)=>event.target===event.currentTarget && setEventModal(false)}>
          <form className="modal event-modal" onSubmit={saveEvent}>
            <div className="modal-header"><div><span className="eyebrow">Новое событие</span><h2>Законодательная задача</h2></div><button type="button" className="icon-button" onClick={()=>setEventModal(false)}><X size={20}/></button></div>
            <div className="modal-body">
              <label className="field"><span>Сотрудник</span><select required value={eventDraft.employeeId} onChange={(e)=>setEventDraft({...eventDraft,employeeId:e.target.value})}><option value="">Выберите сотрудника</option>{employees.filter((employee)=>employee.active || eventDraft.ruleId==="dismissal").map((employee)=><option value={employee.id} key={employee.id}>{employee.fullName}</option>)}</select></label>
              <label className="field"><span>Событие</span><select value={eventDraft.ruleId} onChange={(e)=>changeEventRule(e.target.value)}>{RULES.map((rule)=><option value={rule.id} key={rule.id}>{rule.shortTitle}</option>)}</select></label>
              <div className="form-grid two-col"><Field type="date" label="Дата события" value={eventDraft.eventDate} onChange={changeEventDate}/><Field type="date" label={eventDraft.ruleId==="summons" ? "Дата явки":"Крайний срок"} value={eventDraft.dueDate} onChange={(value)=>setEventDraft({...eventDraft,dueDate:value})}/></div>
              <Field label="Примечание" value={eventDraft.note} placeholder="Что изменилось, номер запроса, особенности вручения…" onChange={(value)=>setEventDraft({...eventDraft,note:value})}/>
              {(() => {const rule=RULES.find((item)=>item.id===eventDraft.ruleId); return <div className="rule-callout"><BookOpen size={18}/><span><strong>{rule?.source}</strong>{rule?.help}<small>Рекомендуемый документ: {rule?.documentHint}</small><a href={rule?.sourceUrl} target="_blank" rel="noreferrer">Открыть источник <ChevronRight size={14}/></a></span></div>})()}
            </div>
            <div className="modal-footer"><span/><div><button type="button" className="button secondary" onClick={()=>setEventModal(false)}>Отмена</button><button className="button primary" type="submit"><Plus size={16}/> Создать задачу</button></div></div>
          </form>
        </div>
      ) : null}

      {importPreview ? (
        <div className="modal-backdrop">
          <section className="modal import-modal">
            <div className="modal-header"><div><span className="eyebrow">Предварительная проверка</span><h2>Импорт Excel</h2></div><button className="icon-button" onClick={()=>setImportPreview(null)}><X size={20}/></button></div>
            <div className="modal-body">
              <div className="import-file"><FileSpreadsheet size={28}/><span><strong>{importPreview.filename}</strong><small>Найдено карточек: {importPreview.employees.length}</small></span></div>
              {importPreview.warnings.length ? <div className="inline-warning"><AlertCircle size={18}/><span><strong>После импорта потребуется проверка</strong>{importPreview.warnings.join(". ")}. Пропуски будут видны в фильтре «Есть пропуски».</span></div> : <div className="inline-success"><Check size={18}/> Явных проблем не обнаружено</div>}
              <p className="import-note">Импорт не создаёт просроченные задачи по старым приказам. Новые сроки появляются только после создания события в программе.</p>
            </div>
            <div className="modal-footer split-footer"><button className="button secondary" onClick={()=>setImportPreview(null)}>Отмена</button><div><button className="button secondary" onClick={()=>commitImport("merge")}>Добавить и обновить</button><button className="button primary" onClick={()=>commitImport("replace")}>Заменить реестр</button></div></div>
          </section>
        </div>
      ) : null}

      {toast ? <div className="toast"><Check size={17}/>{toast}</div> : null}
    </div>
  );
}
