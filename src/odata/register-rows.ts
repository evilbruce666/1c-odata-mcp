import type { Connection } from "../context.js";
import { InputError } from "../errors.js";
import type { EntityMeta, MetaProperty } from "../types/odata.js";
import { classify } from "./metadata.js";
import { fetchAll } from "./pagination.js";
import { requirePublished } from "./publication.js";
import { and, cmp, contains, odataGuid, odataString } from "./query.js";

/**
 * Строки любого опубликованного регистра 1С — для вопросов, которых не покрывают профильные инструменты. Только
 * чтение и только то, что названо: поля обязательны, а отбор, который к регистру неприменим, — отказ, а не пропуск
 * (молча пропущенный отбор по организации вернул бы строки всех организаций).
 */

export const FILTER_OPERATORS = ["eq", "ne", "gt", "ge", "lt", "le", "contains"] as const;
export type FilterOperator = (typeof FILTER_OPERATORS)[number];

export interface RegisterFilter {
  field: string;
  op: FilterOperator;
  value: string | number | boolean;
  /** Тип ссылки для составного поля: «Catalog_Контрагенты». */
  type?: string | undefined;
}

export interface RegisterRequest {
  register: string;
  select: string[];
  filters?: RegisterFilter[] | undefined;
  orgRef?: string | undefined;
  from?: string | undefined;
  to?: string | undefined;
  orderBy?: string | undefined;
  descending?: boolean | undefined;
}

export interface RegisterRows {
  /** Набор, который прочитан на самом деле: у регистра с регистратором это таблица записей «…_RecordType». */
  register: string;
  rows: Array<Record<string, unknown>>;
  /** Строк больше предела: прочитана часть. */
  capped: boolean;
  limit: number;
}

const REGISTER_CLASSES = new Set([
  "accumulationRegister",
  "accountingRegister",
  "informationRegister",
  "calculationRegister",
]);
const RECORDS = "_RecordType";
/** Поля с паролями и токенами не читаются: модели они не нужны, а ответ уходит в её контекст. */
/** «Секретарь» — должность, а не секрет. */
const SECRET_FIELD = /парол|password|секрет(?!ар)|secret(?!ar)|токен|token/i;
const DAY = /^\d{4}-\d{2}-\d{2}$/;
const DATE_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/;
const GUID = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;

function notARegister(name: string): InputError {
  const instead = name.startsWith("Catalog_")
    ? "справочники читают read.counterparty.find_counterparty, read.catalog.list_items и другие инструменты справочников"
    : name.startsWith("Document_")
      ? "документы читают read.document.search_documents и read.document.get_document"
      : "список регистров — read.schema.list_entities с class=informationRegister, accumulationRegister или accountingRegister";
  return new InputError(`«${name}» — не регистр: этот инструмент читает только регистры. Здесь: ${instead}.`);
}

/** Плоская таблица регистра по схеме: «…_RecordType», если она есть, иначе сам независимый регистр. */
async function flatRegister(conn: Connection, requested: string): Promise<EntityMeta> {
  const base = requested.endsWith(RECORDS) ? requested.slice(0, -RECORDS.length) : requested;
  if (!REGISTER_CLASSES.has(classify(base).class) || /[/(?]/.test(requested)) throw notARegister(requested);
  const records = async () => {
    const found = (await conn.getMetadata()).entities.get(base + RECORDS);
    return found?.properties.length ? found : undefined;
  };
  const flat = await records();
  if (flat) return flat;
  // Имя задал пользователь: перед отказом схема перечитывается — регистр могли только что опубликовать.
  await requirePublished(conn, base, "Регистр");
  const added = await records();
  if (added) return added;
  const meta = await conn.getMetadata();
  const own = meta.entities.get(base)!;
  if (own.properties.some((p) => p.name === "RecordSet"))
    throw new InputError(
      `У регистра «${base}» нет таблицы записей «${base}${RECORDS}» в OData: строки вложены в набор записей, ` +
        "и отобрать по ним нельзя.",
    );
  return own;
}

const readable = (p: MetaProperty): boolean =>
  !p.type.startsWith("Collection(") && p.type !== "Edm.Binary" && p.type !== "Edm.Stream";

function field(entity: EntityMeta, name: string, where: string): MetaProperty {
  const found = entity.properties.find((p) => p.name === name);
  if (!found)
    throw new InputError(
      `${where}: у «${entity.entitySet}» нет поля «${name}». Поля — read.schema.describe_entity.`,
    );
  if (!readable(found) || name.endsWith("_Base64Data"))
    throw new InputError(`${where}: поле «${name}» — вложенные или двоичные данные, их здесь не читают.`);
  if (SECRET_FIELD.test(name))
    throw new InputError(`${where}: поле «${name}» похоже на пароль или токен, такие поля не читаются.`);
  return found;
}

function condition(entity: EntityMeta, f: RegisterFilter, published: ReadonlySet<string>): string {
  const meta = field(entity, f.field, "filters");
  const composite = entity.properties.some((p) => p.name === `${f.field}_Type`);
  const bad = (expected: string) =>
    new InputError(`filters: поле «${f.field}» (${meta.type}) — ${expected}.`);
  const onlyEquality = () => {
    if (f.op !== "eq" && f.op !== "ne") throw bad("для ссылки только eq и ne");
  };
  if (f.op === "contains") {
    if (meta.type !== "Edm.String" || composite || typeof f.value !== "string")
      throw bad("contains — только для обычной строки");
    return contains(f.field, f.value);
  }
  if (meta.type === "Edm.Guid") {
    onlyEquality();
    if (typeof f.value !== "string" || !GUID.test(f.value)) throw bad("значение — Ref_Key (GUID)");
    return cmp(f.field, f.op, odataGuid(f.value));
  }
  if (meta.type === "Edm.DateTime") {
    if (typeof f.value !== "string" || !(DAY.test(f.value) || DATE_TIME.test(f.value)))
      throw bad("значение — дата «2026-03-31» или «2026-03-31T00:00:00»");
    // Дата без времени — весь этот день: у записей регистратора есть время, и равенство полуночи их теряет.
    if (DAY.test(f.value) && (f.op === "eq" || f.op === "ne")) {
      const from = cmp(f.field, f.op === "eq" ? "ge" : "lt", `datetime'${f.value}T00:00:00'`);
      const to = cmp(f.field, f.op === "eq" ? "le" : "gt", `datetime'${f.value}T23:59:59'`);
      return `(${from} ${f.op === "eq" ? "and" : "or"} ${to})`;
    }
    const time = f.op === "le" || f.op === "gt" ? "T23:59:59" : "T00:00:00";
    return cmp(f.field, f.op, `datetime'${DAY.test(f.value) ? f.value + time : f.value}'`);
  }
  if (meta.type === "Edm.Boolean") {
    if (typeof f.value !== "boolean" || (f.op !== "eq" && f.op !== "ne"))
      throw bad("true или false, eq и ne");
    return cmp(f.field, f.op, String(f.value));
  }
  if (meta.type === "Edm.String") {
    if (typeof f.value !== "string") throw bad("значение — строка");
    if (composite && GUID.test(f.value)) {
      // Составное поле 1С сравнивает со ссылкой только через cast: с голым guid'…' отвечает ошибкой 500.
      onlyEquality();
      if (!f.type || !published.has(f.type) || !/^[A-Za-z]+_[А-Яа-яЁёA-Za-z0-9_]+$/.test(f.type))
        throw bad("составная ссылка: укажите в type опубликованный тип, напр. «Catalog_Контрагенты»");
      return cmp(f.field, f.op, `cast(${odataGuid(f.value)}, '${f.type}')`);
    }
    return cmp(f.field, f.op, odataString(f.value));
  }
  if (typeof f.value !== "number" || !Number.isFinite(f.value)) throw bad("значение — число");
  return cmp(f.field, f.op, String(f.value));
}

function compare(a: unknown, b: unknown): number {
  if (typeof a === "number" && typeof b === "number") return a - b;
  return String(a ?? "").localeCompare(String(b ?? ""), "ru");
}

export async function readRegisterRows(conn: Connection, request: RegisterRequest): Promise<RegisterRows> {
  const entity = await flatRegister(conn, request.register);
  const published = await conn.available();
  const has = (name: string) => entity.properties.some((p) => p.name === name);
  const select = [...new Set(request.select)];
  for (const name of select) field(entity, name, "select");
  if (request.orderBy) {
    field(entity, request.orderBy, "orderBy");
    if (!select.includes(request.orderBy)) select.push(request.orderBy);
  }

  const conditions = (request.filters ?? []).map((f) => condition(entity, f, published));
  if (request.orgRef !== undefined) {
    if (!has("Организация_Key"))
      throw new InputError(
        `organization: в «${entity.entitySet}» нет поля «Организация_Key». Уберите organization ` +
          "или отберите по другому полю в filters.",
      );
    conditions.push(cmp("Организация_Key", "eq", odataGuid(request.orgRef)));
  }
  if (request.from !== undefined || request.to !== undefined) {
    if (!has("Period"))
      throw new InputError(
        `from и to: в «${entity.entitySet}» нет поля «Period». Отберите по своему полю даты в filters.`,
      );
    if (request.from) conditions.push(cmp("Period", "ge", `datetime'${request.from}T00:00:00'`));
    if (request.to) conditions.push(cmp("Period", "le", `datetime'${request.to}T23:59:59'`));
  }

  // Один запрос: $skip на регистрах теряет строки (pagination.ts), а предел — размер одного ответа 1С.
  const limit = conn.behavior.maxRows;
  const direction = request.descending ? "desc" : "asc";
  const { rows, truncated } = await fetchAll(
    conn.client,
    entity.entitySet,
    {
      select,
      filter: and(...conditions) || undefined,
      orderby: request.orderBy ? `${request.orderBy} ${direction}` : undefined,
    },
    limit,
    limit,
  );
  const own = rows.map((r) => Object.fromEntries(select.map((name) => [name, r[name]])));
  // Порядок, в котором 1С отдаёт равные строки, не обещан: задаём свой, чтобы страницы не перемешивались.
  const key = (r: Record<string, unknown>) => JSON.stringify(select.map((name) => r[name]));
  const sign = request.descending ? -1 : 1;
  own.sort(
    (a, b) =>
      (request.orderBy ? sign * compare(a[request.orderBy], b[request.orderBy]) : 0) ||
      key(a).localeCompare(key(b), "ru"),
  );
  return { register: entity.entitySet, rows: own, capped: truncated, limit };
}
