/**
 * Типобезопасный билдер OData-запросов (v3).
 *
 * Пользовательский ввод НИКОГДА не склеивается в $filter напрямую —
 * только через эти хелперы с корректным экранированием. Это защита
 * от инъекций в запрос и от сломанного синтаксиса.
 */

/** Экранирует строковый литерал OData: одинарная кавычка удваивается. */
export function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** GUID-литерал OData v3: guid'xxxxxxxx-...'. Валидирует формат. */
export function odataGuid(value: string): string {
  const v = value.trim().replace(/[{}]/g, "");
  if (!/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(v)) {
    throw new Error(`Некорректный GUID: ${value}`);
  }
  return `guid'${v}'`;
}

/** datetime-литерал OData v3 из Date или ISO-строки. */
export function odataDateTime(value: string | Date): string {
  const d = typeof value === "string" ? new Date(value) : value;
  if (Number.isNaN(d.getTime())) throw new Error(`Некорректная дата: ${String(value)}`);
  // 1С ожидает datetime'YYYY-MM-DDTHH:mm:ss' (без зоны/мс).
  const iso = d.toISOString().replace(/\.\d{3}Z$/, "");
  return `datetime'${iso}'`;
}

/** Условие «поле содержит подстроку» (OData v3: substringof). */
export function contains(field: string, substr: string): string {
  return `substringof(${odataString(substr)}, ${field})`;
}

export type Comparison = "eq" | "ne" | "gt" | "ge" | "lt" | "le";

/** Бинарное условие: <field> <op> <literal>. literal уже должен быть экранирован. */
export function cmp(field: string, op: Comparison, literal: string): string {
  return `${field} ${op} ${literal}`;
}

export function and(...parts: Array<string | undefined>): string {
  return parts.filter(Boolean).join(" and ");
}

export function or(...parts: Array<string | undefined>): string {
  const items = parts.filter(Boolean);
  return items.length > 1 ? `(${items.join(" or ")})` : items.join(" or ");
}

/** Опции OData-запроса. */
export interface QueryOptions {
  select?: string[];
  filter?: string;
  orderby?: string;
  top?: number;
  skip?: number;
  expand?: string[];
  /** Запросить общее число записей (odata.count). */
  count?: boolean;
}

/**
 * Собирает query-string из опций. $format=json добавляется всегда.
 * Возвращает строку вида "?$top=100&$filter=...".
 */
export function buildQuery(opts: QueryOptions): string {
  const params = new URLSearchParams();
  params.set("$format", "json");

  if (opts.select?.length) params.set("$select", opts.select.join(","));
  if (opts.filter) params.set("$filter", opts.filter);
  if (opts.orderby) params.set("$orderby", opts.orderby);
  if (typeof opts.top === "number") params.set("$top", String(opts.top));
  if (typeof opts.skip === "number") params.set("$skip", String(opts.skip));
  if (opts.expand?.length) params.set("$expand", opts.expand.join(","));
  if (opts.count) params.set("$inlinecount", "allpages");

  return `?${params.toString()}`;
}
