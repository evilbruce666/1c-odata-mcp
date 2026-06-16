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
 * Кодирует значение параметра для 1С OData.
 *
 * ВАЖНО: используем encodeURIComponent, а НЕ URLSearchParams. Последний
 * кодирует пробел как '+', а 1С внутри $filter НЕ декодирует '+' в пробел
 * (отвечает 400). encodeURIComponent кодирует пробел как %20 — это работает.
 * Ключи параметров ($filter и т.п.) оставляем без кодирования.
 */
function encodeValue(v: string): string {
  return encodeURIComponent(v);
}

/**
 * Собирает query-string из опций. $format=json добавляется всегда.
 * Возвращает строку вида "?$top=100&$filter=...".
 */
export function buildQuery(opts: QueryOptions): string {
  const parts: string[] = ["$format=json"];

  if (opts.select?.length) parts.push(`$select=${encodeValue(opts.select.join(","))}`);
  if (opts.filter) parts.push(`$filter=${encodeValue(opts.filter)}`);
  if (opts.orderby) parts.push(`$orderby=${encodeValue(opts.orderby)}`);
  if (typeof opts.top === "number") parts.push(`$top=${opts.top}`);
  if (typeof opts.skip === "number") parts.push(`$skip=${opts.skip}`);
  if (opts.expand?.length) parts.push(`$expand=${encodeValue(opts.expand.join(","))}`);
  if (opts.count) parts.push("$inlinecount=allpages");

  return `?${parts.join("&")}`;
}
