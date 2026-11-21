import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { pageCached } from "./cache.js";
import { callPageStart } from "./call-context.js";
import { num } from "./accounting.js";
import { tradeWithoutAccounting } from "./configuration.js";
import type { StockLine } from "./inventory.js";
import { balancePath } from "./stock.js";

/**
 * Остатки с суммами в Управлении торговлей (1.1): регистр накопления «Себестоимость товаров». В УТ нет Хозрасчётного,
 * а «Товары на складах» хранит только количество; стоимость запаса — здесь, по ключу аналитики учёта номенклатуры
 * (товар, место хранения). Сумма — в валюте регламентированного учёта (СтоимостьРегл). Проверено на копии УТ 11.5
 * (21.11.2026): $expand ключа аналитики отдаёт товар и склад за один проход; количество сходится с «Товарами на
 * складах» с точностью до товаров вне складов (в пути, у комиссионера).
 */
export const COST_REGISTER = "AccumulationRegister_СебестоимостьТоваров";
const KEY = "АналитикаУчетаНоменклатуры";
const KEYS_CATALOG = "Catalog_КлючиАналитикиУчетаНоменклатуры";
const COST_CALC = "Document_РасчетСебестоимостиТоваров";
const WAREHOUSE_TYPE = "StandardODATA.Catalog_Склады";
const EMPTY = "00000000-0000-0000-0000-000000000000";

/** УТ без Хозрасчётного, где опубликованы регистр себестоимости и ключи аналитики. */
export async function tradeCostAvailable(conn: Connection): Promise<boolean> {
  if (!(await tradeWithoutAccounting(conn))) return false;
  const available = await conn.available();
  return available.has(COST_REGISTER) && available.has(KEYS_CATALOG);
}

const refOrUndefined = (v: unknown): string | undefined =>
  typeof v === "string" && v && v !== EMPTY ? v : undefined;

/**
 * Остатки себестоимости, свёрнутые по товару (и складу при byWarehouse). Место хранения не склад (партнёр, подразделение)
 * при разбивке по складам — строка без склада.
 */
export async function readTradeCost(
  conn: Connection,
  args: { orgKey?: string | undefined; asOf?: string | undefined; byWarehouse: boolean },
): Promise<{ rows: number; lines: StockLine[]; lastCalc?: string }> {
  const rows = await pageCached(
    conn,
    `cost-balance:${args.asOf ?? ""}:${args.orgKey ?? ""}`,
    callPageStart() === 0,
    async () =>
      (
        await fetchAllForAggregation(
          conn,
          balancePath(
            COST_REGISTER,
            args.asOf,
            args.orgKey ? `Организация_Key eq guid'${args.orgKey}'` : undefined,
          ),
          {
            select: [
              "КоличествоBalance",
              "СтоимостьРеглBalance",
              `${KEY}/Номенклатура_Key`,
              `${KEY}/МестоХранения`,
              `${KEY}/МестоХранения_Type`,
            ],
            expand: [KEY],
          },
          `себестоимость на ${args.asOf ?? "сейчас"}`,
        )
      ).rows,
  );
  const lines = new Map<string, StockLine>();
  for (const r of rows) {
    const key = (r[KEY] ?? {}) as ODataEntity;
    const itemRef = refOrUndefined(key["Номенклатура_Key"]);
    const warehouseRef =
      args.byWarehouse && key["МестоХранения_Type"] === WAREHOUSE_TYPE
        ? refOrUndefined(key["МестоХранения"])
        : undefined;
    const k = `${itemRef ?? ""}|${warehouseRef ?? ""}`;
    const line = lines.get(k) ?? {
      ...(itemRef ? { itemRef } : {}),
      ...(warehouseRef ? { warehouseRef, warehouseSet: "Catalog_Склады" } : {}),
      quantity: 0,
      cents: 0,
    };
    line.quantity += num(r["КоличествоBalance"]);
    line.cents += Math.round(num(r["СтоимостьРеглBalance"]) * 100);
    lines.set(k, line);
  }
  const lastCalc = await lastCostCalculation(conn, args.asOf);
  return {
    rows: rows.length,
    lines: [...lines.values()].filter((l) => Math.abs(l.quantity) > 0.0001 || l.cents !== 0),
    ...(lastCalc ? { lastCalc } : {}),
  };
}

/** Дата последнего проведённого расчёта себестоимости не позже даты остатков — для заметки о точности сумм. */
async function lastCostCalculation(conn: Connection, asOf: string | undefined): Promise<string | undefined> {
  if (!(await conn.available()).has(COST_CALC)) return undefined;
  // Без asOf — не позже сегодняшнего дня: документ расчёта датируется концом месяца и может стоять впереди.
  const day = asOf ?? new Date().toISOString().slice(0, 10);
  const until = ` and Date le datetime'${day}T23:59:59'`;
  const r = await conn.client
    .getCollection<ODataEntity>(
      `${COST_CALC}?$format=json&$top=1&$orderby=Date desc&$select=Date&$filter=${encodeURIComponent(`Posted eq true${until}`)}`,
    )
    .catch(() => undefined);
  const d = r?.value[0]?.["Date"];
  return typeof d === "string" ? d.slice(0, 10) : undefined;
}
