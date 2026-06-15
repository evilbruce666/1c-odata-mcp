import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid, odataString, or } from "./query.js";
import { REGISTERS } from "../config/mapping.js";
import { requireEntity } from "./publication.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Аналитика для 1С:Бухгалтерия 3.0 строится на регистре бухгалтерии
 * «Хозрасчетный» и его виртуальной таблице Balance. И дебиторка (сч. 62),
 * и остатки товаров (сч. 41) — это сальдо по соответствующему счёту,
 * с контрагентом/номенклатурой в ExtDimension1.
 */

const CHART_CANDIDATES = ["ChartOfAccounts_Хозрасчетный"] as const;

export interface Account {
  key: string;
  code: string;
  description: string;
}

/** Возвращает счета, код которых начинается с одного из префиксов (напр. "62", "41"). */
export async function resolveAccounts(
  conn: Connection,
  prefixes: readonly string[],
): Promise<Account[]> {
  const chart = await requireEntity(conn, CHART_CANDIDATES, "План счетов «Хозрасчётный»");
  const filter = or(...prefixes.map((p) => `startswith(Code, ${odataString(p)})`));
  const { rows } = await fetchAll(
    conn.client,
    chart,
    { filter, select: ["Ref_Key", "Code", "Description"], orderby: "Code" },
    conn.behavior.pageSize,
    conn.behavior.maxRows,
  );
  return rows.map((r) => ({
    key: String(r["Ref_Key"] ?? ""),
    code: String(r["Code"] ?? ""),
    description: String(r["Description"] ?? ""),
  }));
}

/**
 * Возвращает карту «код счёта → Ref_Key» для заданных точных кодов
 * (напр. ["41.01","60.01","90.01.1"]). Нужно для заполнения счетов учёта
 * в документах — через OData автозаполнение 1С не срабатывает.
 */
export async function accountsByCode(
  conn: Connection,
  codes: readonly string[],
): Promise<Map<string, string>> {
  const result = new Map<string, string>();
  if (codes.length === 0) return result;
  const chart = await requireEntity(conn, CHART_CANDIDATES, "План счетов «Хозрасчётный»");
  const filter = or(...codes.map((c) => cmp("Code", "eq", odataString(c))));
  const { rows } = await fetchAll(
    conn.client,
    chart,
    { filter, select: ["Ref_Key", "Code"] },
    codes.length,
    codes.length,
  );
  for (const r of rows) result.set(String(r["Code"]), String(r["Ref_Key"]));
  return result;
}

/**
 * Тянет строки сальдо регистра Хозрасчетный, отфильтрованные по набору счетов
 * и (необязательно) по организации.
 */
export async function balanceByAccounts(
  conn: Connection,
  accountKeys: string[],
  maxRows: number,
  orgKey?: string,
): Promise<ODataEntity[]> {
  if (accountKeys.length === 0) return [];
  const reg = await requireEntity(conn, REGISTERS.accounting, "Регистр бухгалтерии «Хозрасчётный»");
  const filter = and(
    or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))),
    orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
  );
  const { rows } = await fetchAll(conn.client, `${reg}/Balance`, { filter }, conn.behavior.pageSize, maxRows);
  return rows;
}

/**
 * Резолвит GUID-ы в наименования (Description) из справочника батчами.
 * Используется, чтобы показать имена контрагентов/номенклатуры вместо GUID.
 */
export async function resolveNames(
  conn: Connection,
  entitySet: string,
  keys: Iterable<string>,
): Promise<Map<string, string>> {
  const unique = [...new Set([...keys].filter(Boolean))];
  const result = new Map<string, string>();
  const available = await conn.available();
  if (!available.has(entitySet)) return result;

  const BATCH = 20;
  for (let i = 0; i < unique.length; i += BATCH) {
    const batch = unique.slice(i, i + BATCH);
    const filter = or(...batch.map((k) => cmp("Ref_Key", "eq", odataGuid(k))));
    const { rows } = await fetchAll(
      conn.client,
      entitySet,
      { filter, select: ["Ref_Key", "Description"] },
      BATCH,
      BATCH,
    );
    for (const r of rows) result.set(String(r["Ref_Key"]), String(r["Description"] ?? ""));
  }
  return result;
}

/** Число из поля сальдо (OData может отдавать строкой). */
export function num(v: unknown): number {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}

export { and, cmp };
