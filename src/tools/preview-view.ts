import type { Connection } from "../context.js";
import { CATALOGS } from "../config/mapping.js";
import { documentTitle, printNumber } from "../odata/names.js";
import { listOrganizations } from "../odata/orgs.js";
import { buildQuery, cmp, odataGuid, or } from "../odata/query.js";
import type { EntityMeta, ODataEntity } from "../types/odata.js";

/**
 * Читаемый предпросмотр записи: то, что человек подтверждает, — подписи вместо технических
 * имён и названия вместо GUID («Контрагент: ООО Ромашка», «Счет учета: 41.01 Товары на складах»),
 * строки табличных частей и для изменения — «было → станет». Машинная часть ответа (payload,
 * fields) не меняется: подтверждение отправляет именно её. Идея — из форка 1c-odata-mcp-kz (#34).
 *
 * Цель ссылки берётся из связей $metadata (Association), у составных полей — из соседнего *_Type.
 * Названия читаются одним запросом на справочник (не больше MAX_SETS справочников и MAX_REFS
 * ссылок); сбой чтения вид не роняет — остаётся GUID.
 */

export interface PreviewView {
  title: string;
  fields: Record<string, unknown>;
  tables?: Record<string, Array<Record<string, unknown>>>;
  changes?: Array<{ field: string; from: unknown; to: unknown }>;
}

const MAX_REFS = 50;
/**
 * Названия читаются одним запросом на справочник, а на облачной 1С каждый запрос — около секунды.
 * Поэтому — не больше MAX_SETS справочников, в порядке важности для того, кто подтверждает:
 * с кем, по какому договору, что и на какие счета. Ссылки прочих справочников остаются GUID.
 */
const MAX_SETS = 4;
const NAME_PRIORITY = [
  "Catalog_Контрагенты",
  "Catalog_ДоговорыКонтрагентов",
  "Catalog_Номенклатура",
  "ChartOfAccounts_Хозрасчетный",
  "Catalog_БанковскиеСчета",
  "Catalog_Склады",
  "Catalog_СтатьиДвиженияДенежныхСредств",
];
const EMPTY_REF = "00000000-0000-0000-0000-000000000000";
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Служебные поля: человеку при подтверждении ничего не говорят. */
const HIDDEN = new Set(["Ref_Key", "DataVersion", "LineNumber", "Predefined", "PredefinedDataName"]);

/** Стандартные реквизиты 1С публикуются по-английски. */
const STANDARD_LABELS: Record<string, string> = {
  Date: "Дата",
  Number: "Номер",
  Posted: "Проведён",
  DeletionMark: "Пометка удаления",
  Description: "Наименование",
  Code: "Код",
  IsFolder: "Это группа",
  Parent_Key: "Группа",
  Owner_Key: "Владелец",
};
/** Ложь, которую человеку важно видеть и при создании. */
const MEANINGFUL_FALSE = new Set(["СуммаВключаетНДС"]);

const isRef = (v: unknown): v is string => typeof v === "string" && GUID.test(v) && v !== EMPTY_REF;
const isEmptyRef = (v: unknown) => v === EMPTY_REF || v === "";

/**
 * Значение по умолчанию, которое при создании ничего не говорит: пустая строка и ссылка, пустая дата
 * 1С, ноль, «нет». При изменении не скрывается — там это и есть новое значение.
 */
function isDefault(name: string, v: unknown): boolean {
  if (v === 0) return name !== "СуммаДокумента";
  if (v === null || v === "" || v === EMPTY_REF) return true;
  if (v === false) return !MEANINGFUL_FALSE.has(name);
  // Служебное значение в JSON (адрес, контактная информация) — его читаемый вид лежит в соседнем поле.
  if (typeof v === "string" && /^\s*\{[\s\S]*\}\s*$/.test(v)) return true;
  return typeof v === "string" && v.startsWith("0001-01-01");
}

/** «ДоговорКонтрагента_Key» → «Договор контрагента». */
export function fieldLabel(name: string): string {
  return STANDARD_LABELS[name] ?? documentTitle(name.replace(/_Key$/, ""));
}

/** Имя связи поля и его цель: по Association, у составного поля — по значению *_Type. */
function targetOf(
  meta: EntityMeta | undefined,
  row: Record<string, unknown>,
  field: string,
): string | undefined {
  const base = field.replace(/_Key$/, "");
  const typed = row[`${base}_Type`];
  if (typeof typed === "string" && typed) return typed.replace(/^StandardODATA\./, "");
  if (!field.endsWith("_Key")) return undefined;
  return meta?.navigations.find((n) => n.name === base)?.toType || undefined;
}

function display(value: unknown): unknown {
  if (typeof value === "boolean") return value ? "да" : "нет";
  // Дата 1С: «2026-10-24T00:00:00» → «2026-10-24», со временем — «2026-10-24 05:55».
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(value))
    return value.endsWith("T00:00:00") ? value.slice(0, 10) : value.slice(0, 16).replace("T", " ");
  return value;
}

/** Название объекта: «41.01 Товары на складах», «ООО Ромашка», «Счет на оплату покупателю № 33 от 2026-09-01». */
function titleOf(set: string, r: ODataEntity): string | undefined {
  const s = (k: string) => (typeof r[k] === "string" ? (r[k] as string).trim() : "");
  if (set.startsWith("ChartOfAccounts_"))
    return [s("Code"), s("Description")].filter(Boolean).join(" ") || undefined;
  if (set.startsWith("Document_") && s("Number"))
    return `${documentTitle(set)} № ${printNumber(s("Number"))}${s("Date") ? ` от ${s("Date").slice(0, 10)}` : ""}`;
  return s("Description") || s("Code") || undefined;
}

async function readNames(conn: Connection, wanted: Map<string, Set<string>>): Promise<Map<string, string>> {
  const names = new Map<string, string>();
  if (wanted.size === 0) return names;
  // Организации — из кеша подключения (их уже читал выбор организации), без отдельного запроса.
  if (wanted.has(CATALOGS.organizations[0]))
    for (const o of await listOrganizations(conn).catch(() => [])) names.set(o.ref.toLowerCase(), o.name);
  const [meta, available] = await Promise.all([conn.getMetadata(), conn.available()]);
  const sets = NAME_PRIORITY.filter((set) => wanted.has(set) && available.has(set)).slice(0, MAX_SETS);
  for (const set of sets) {
    const refs = wanted.get(set)!;
    const props = new Set((meta.entities.get(set)?.properties ?? []).map((p) => p.name));
    const select = ["Ref_Key", "Description", "Code", "Number", "Date"].filter((f) => props.has(f));
    try {
      const page = await conn.client.getCollection(
        `${set}${buildQuery({ select, filter: or(...[...refs].map((g) => cmp("Ref_Key", "eq", odataGuid(g)))) })}`,
      );
      for (const r of page.value ?? []) {
        const text = titleOf(set, r);
        if (text) names.set(String(r["Ref_Key"]).toLowerCase(), text);
      }
    } catch {
      // Название не прочиталось — в виде останется GUID.
    }
  }
  return names;
}

/** Пустое значение в «было → станет» — «—», чтобы очистка поля читалась как изменение. */
const shown = (v: unknown) => (v === undefined || v === null || v === "" ? "—" : v);

/** Строит вид записи. before — текущие значения полей объекта (для изменения). */
export async function buildPreviewView(
  conn: Connection,
  entitySet: string,
  data: Record<string, unknown>,
  before?: Record<string, unknown>,
): Promise<PreviewView> {
  const meta = await conn.getMetadata();
  const head = meta.entities.get(entitySet);

  // Сначала собираем ссылки всех уровней, чтобы прочитать названия пачкой.
  const wanted = new Map<string, Set<string>>();
  let count = 0;
  const want = (target: string | undefined, value: unknown) => {
    if (!target || !isRef(value) || count >= MAX_REFS) return;
    const refs = wanted.get(target) ?? new Set<string>();
    if (!refs.has(value.toLowerCase())) count++;
    refs.add(value.toLowerCase());
    wanted.set(target, refs);
  };
  const visit = (em: EntityMeta | undefined, row: Record<string, unknown>) => {
    for (const [k, v] of Object.entries(row)) want(targetOf(em, row, k), v);
  };
  visit(head, data);
  if (before) visit(head, before);
  for (const [k, v] of Object.entries(data))
    if (Array.isArray(v))
      for (const row of v as Array<Record<string, unknown>>)
        visit(meta.entities.get(`${entitySet}_${k}`), row);
  const names = await readNames(conn, wanted);

  // При изменении очистка поля — тоже изменение: пустую ссылку показываем «—», а не скрываем.
  const render = (em: EntityMeta | undefined, row: Record<string, unknown>): Record<string, unknown> => {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(row)) {
      if (HIDDEN.has(k) || k.endsWith("_Type") || Array.isArray(v) || v === undefined) continue;
      // «Удалить…» — устаревшие реквизиты конфигурации, в форме документа их нет.
      if (!before && (k.startsWith("Удалить") || isDefault(k, v))) continue;
      if (isEmptyRef(v) && targetOf(em, row, k)) {
        if (before) out[fieldLabel(k)] = "—";
        continue; // пустая ссылка — поле не заполнено
      }
      out[fieldLabel(k)] = isRef(v) ? (names.get(v.toLowerCase()) ?? v) : display(v);
    }
    return out;
  };

  const view: PreviewView = { title: documentTitle(entitySet), fields: render(head, data) };
  const tables = Object.entries(data).filter(([, v]) => Array.isArray(v) && v.length > 0);
  if (tables.length)
    view.tables = Object.fromEntries(
      tables.map(([k, v]) => [
        fieldLabel(k),
        (v as Array<Record<string, unknown>>).map((row) =>
          render(meta.entities.get(`${entitySet}_${k}`), row),
        ),
      ]),
    );
  if (before) {
    const was = render(head, before);
    const changes = Object.entries(view.fields)
      .filter(([label, to]) => String(was[label] ?? "") !== String(to ?? ""))
      .map(([field, to]) => ({ field, from: shown(was[field]), to: shown(to) }));
    if (changes.length) view.changes = changes;
  }
  return view;
}

/** Вид или undefined: предпросмотр без читаемого вида лучше, чем отказ предпросмотра. */
export async function tryPreviewView(
  conn: Connection,
  entitySet: string,
  data: Record<string, unknown>,
  before?: Record<string, unknown>,
): Promise<{ view?: PreviewView }> {
  try {
    return { view: await buildPreviewView(conn, entitySet, data, before) };
  } catch {
    return {};
  }
}
