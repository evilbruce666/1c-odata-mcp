/**
 * Названия вместо ссылок в документах: контрагент и организация в строках поиска, «…_Название» в компактном
 * документе. Читаются пачками по 20 (resolveNames), организации — из кэша подключения. Сбой справочника не роняет
 * ответ: названия просто не появляются.
 */
import type { Connection } from "../context.js";
import { CATALOGS } from "../config/mapping.js";
import { GUID_RE, resolveNames } from "../odata/accounting.js";
import { listOrganizations } from "../odata/orgs.js";
import { readCatalogBrief, wholeCatalogCheaper } from "../odata/catalog-whole.js";
import { fitToChars, renderJson } from "./_shared.js";

const EMPTY_REF = "00000000-0000-0000-0000-000000000000";
const COUNTERPARTIES = CATALOGS.counterparties[0];
const ORGANIZATIONS = CATALOGS.organizations[0];
const NOMENCLATURE = CATALOGS.nomenclature[0];
/** Больше стольких позиций номенклатуры — справочником целиком, если так дешевле (как в предпросмотре). */
const WHOLE_FROM = 200;

const asRef = (v: unknown): string | undefined =>
  typeof v === "string" && GUID_RE.test(v) && v !== EMPTY_REF ? v : undefined;

/** Названия по справочникам: ключ — набор, значения — ссылки. Ответ — «набор|ссылка в нижнем регистре» → название. */
export async function namesBySet(
  conn: Connection,
  wanted: ReadonlyMap<string, ReadonlySet<string>>,
): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const available = await conn.available();
  for (const [set, refs] of wanted) {
    if (!refs.size || !available.has(set)) continue;
    const put = (ref: string, name: unknown) => {
      if (typeof name === "string" && name) out.set(`${set}|${ref.toLowerCase()}`, name);
    };
    try {
      if (set === ORGANIZATIONS) {
        for (const o of await listOrganizations(conn)) put(o.ref, o.name);
      } else if (
        set === NOMENCLATURE &&
        refs.size > WHOLE_FROM &&
        (await wholeCatalogCheaper(conn, set, refs.size, 20))
      ) {
        const brief = await readCatalogBrief(conn, set);
        for (const ref of refs) put(ref, brief.get(ref.toLowerCase())?.["Description"]);
      } else {
        for (const [ref, name] of await resolveNames(conn, set, refs)) put(ref, name);
      }
    } catch {
      // Справочник не прочитан — названий этого справочника в ответе нет.
    }
  }
  return out;
}

const nameOf = (names: Map<string, string>, set: string, ref: string | undefined) =>
  ref ? names.get(`${set}|${ref.toLowerCase()}`) : undefined;

/** Ссылки строки поиска: до подстановки названий хранятся отдельно и в ответ не попадают. */
export interface DocumentRefs {
  counterparty?: string | undefined;
  organization?: string | undefined;
}

export function documentRefs(row: Record<string, unknown>): DocumentRefs {
  const polyType = String(row["Контрагент_Type"] ?? "").replace(/^StandardODATA\./, "");
  return {
    counterparty:
      asRef(row["Контрагент_Key"]) ?? (polyType === COUNTERPARTIES ? asRef(row["Контрагент"]) : undefined),
    organization: asRef(row["Организация_Key"]),
  };
}

/** Дописывает counterparty и organization строкам (на месте). */
async function addDocumentNames<T extends object>(
  conn: Connection,
  rows: readonly T[],
  refsOf: WeakMap<object, DocumentRefs>,
): Promise<void> {
  const wanted = new Map<string, Set<string>>([
    [COUNTERPARTIES, new Set()],
    [ORGANIZATIONS, new Set()],
  ]);
  for (const r of rows) {
    const refs = refsOf.get(r);
    if (refs?.counterparty) wanted.get(COUNTERPARTIES)!.add(refs.counterparty);
    if (refs?.organization) wanted.get(ORGANIZATIONS)!.add(refs.organization);
  }
  const names = await namesBySet(conn, wanted);
  for (const r of rows) {
    const refs = refsOf.get(r);
    const counterparty = nameOf(names, COUNTERPARTIES, refs?.counterparty);
    const organization = nameOf(names, ORGANIZATIONS, refs?.organization);
    Object.assign(r, counterparty ? { counterparty } : {}, organization ? { organization } : {});
  }
}

/**
 * Названия строкам будущей страницы — одним чтением. Названия только удлиняют строку, а pageOf берёт начало списка до
 * первой непомещающейся (fitToChars), поэтому дальше строк, помещающихся без названий, страница не уйдёт. Их ссылки —
 * начало списка ссылок всех limit строк: запросов не больше, чем при чтении всем, повторы и сбой справочника — общие.
 */
export async function addPageNames<T extends object>(
  conn: Connection,
  rows: readonly T[],
  page: { offset: number; limit: number; head: unknown },
  refsOf: WeakMap<object, DocumentRefs>,
): Promise<void> {
  const candidates = rows.slice(page.offset, page.offset + page.limit);
  await addDocumentNames(conn, fitToChars(candidates, renderJson(page.head).length), refsOf);
}

/** Поле-ссылка компактного документа → справочник названия. */
const REF_FIELDS: Record<string, string> = {
  Контрагент_Key: COUNTERPARTIES,
  Организация_Key: ORGANIZATIONS,
  ДоговорКонтрагента_Key: "Catalog_ДоговорыКонтрагентов",
  Договор_Key: "Catalog_ДоговорыКонтрагентов",
  Партнер_Key: "Catalog_Партнеры",
  Склад_Key: "Catalog_Склады",
  СкладОтправитель_Key: "Catalog_Склады",
  СкладПолучатель_Key: "Catalog_Склады",
  Номенклатура_Key: NOMENCLATURE,
};

const isEmpty = (v: unknown) =>
  v === null ||
  v === "" ||
  v === EMPTY_REF ||
  (Array.isArray(v) && v.length === 0) ||
  (typeof v === "string" && v.startsWith("0001-01-01T"));
const isServiceKey = (k: string) => k.startsWith("odata.") || k.includes("@");

/** Ключ названия: «Контрагент_Key» и составной «Контрагент» → «Контрагент_Название». */
const nameKey = (field: string) => `${field.replace(/_Key$/, "")}_Название`;

function target(row: Record<string, unknown>, field: string): { set: string; ref: string } | undefined {
  const ref = asRef(row[field]);
  if (!ref) return undefined;
  const typed = String(row[`${field}_Type`] ?? "").replace(/^StandardODATA\./, "");
  if (typed.startsWith("Catalog_")) return { set: typed, ref };
  const set = REF_FIELDS[field];
  return set ? { set, ref } : undefined;
}

function walk(value: unknown, visit: (row: Record<string, unknown>) => Record<string, unknown>): unknown {
  if (Array.isArray(value)) return value.map((v) => walk(v, visit));
  if (!value || typeof value !== "object") return value;
  const row = visit(value as Record<string, unknown>);
  return Object.fromEntries(Object.entries(row).map(([k, v]) => [k, walk(v, visit)]));
}

/**
 * Документ без пустых значений и служебных ключей OData, с названиями рядом со ссылками. Ноль и «нет» остаются: они
 * несут смысл (не проведён, сумма 0). Строки табличных частей — так же.
 */
export async function compactDocument(
  conn: Connection,
  entity: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const wanted = new Map<string, Set<string>>();
  walk(entity, (row) => {
    for (const k of Object.keys(row)) {
      const t = target(row, k);
      if (!t) continue;
      const refs = wanted.get(t.set) ?? new Set<string>();
      refs.add(t.ref);
      wanted.set(t.set, refs);
    }
    return row;
  });
  const names = await namesBySet(conn, wanted);
  return walk(entity, (row) => {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(row)) {
      if (isServiceKey(k) || isEmpty(v)) continue;
      out[k] = v;
      const t = target(row, k);
      const name = t && nameOf(names, t.set, t.ref);
      const key = nameKey(k);
      if (name && !(key in row)) out[key] = name;
    }
    return out;
  }) as Record<string, unknown>;
}
