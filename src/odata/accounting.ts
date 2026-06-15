import type { ServerContext } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid, odataString, or } from "./query.js";
import { REGISTERS, resolveEntity } from "../config/mapping.js";
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
  ctx: ServerContext,
  prefixes: readonly string[],
): Promise<Account[]> {
  const available = await ctx.available();
  const chart = resolveEntity(CHART_CANDIDATES, available);
  if (!chart) {
    throw new Error(
      "План счетов Хозрасчетный не опубликован в OData. Добавьте ChartOfAccounts_Хозрасчетный в «Состав OData».",
    );
  }
  const filter = or(...prefixes.map((p) => `startswith(Code, ${odataString(p)})`));
  const { rows } = await fetchAll(
    ctx.client,
    chart,
    { filter, select: ["Ref_Key", "Code", "Description"], orderby: "Code" },
    ctx.cfg.ODATA_PAGE_SIZE,
    ctx.cfg.ODATA_MAX_ROWS,
  );
  return rows.map((r) => ({
    key: String(r["Ref_Key"] ?? ""),
    code: String(r["Code"] ?? ""),
    description: String(r["Description"] ?? ""),
  }));
}

/** Тянет строки сальдо регистра Хозрасчетный, отфильтрованные по набору счетов. */
export async function balanceByAccounts(
  ctx: ServerContext,
  accountKeys: string[],
  maxRows: number,
): Promise<ODataEntity[]> {
  const available = await ctx.available();
  const reg = resolveEntity(REGISTERS.accounting, available);
  if (!reg) {
    throw new Error(
      "Регистр бухгалтерии не опубликован в OData. Добавьте AccountingRegister_Хозрасчетный в «Состав OData».",
    );
  }
  if (accountKeys.length === 0) return [];
  const filter = or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k))));
  const { rows } = await fetchAll(ctx.client, `${reg}/Balance`, { filter }, ctx.cfg.ODATA_PAGE_SIZE, maxRows);
  return rows;
}

/**
 * Резолвит GUID-ы в наименования (Description) из справочника батчами.
 * Используется, чтобы показать имена контрагентов/номенклатуры вместо GUID.
 */
export async function resolveNames(
  ctx: ServerContext,
  entitySet: string,
  keys: Iterable<string>,
): Promise<Map<string, string>> {
  const unique = [...new Set([...keys].filter(Boolean))];
  const result = new Map<string, string>();
  const available = await ctx.available();
  if (!available.has(entitySet)) return result;

  const BATCH = 20;
  for (let i = 0; i < unique.length; i += BATCH) {
    const batch = unique.slice(i, i + BATCH);
    const filter = or(...batch.map((k) => cmp("Ref_Key", "eq", odataGuid(k))));
    const { rows } = await fetchAll(
      ctx.client,
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
