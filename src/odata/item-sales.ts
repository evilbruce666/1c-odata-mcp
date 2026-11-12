import type { Connection } from "../context.js";
import { resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { nextDay, num } from "./accounting.js";
import { requireEntity } from "./publication.js";
import { and, cmp, odataGuid, or } from "./query.js";
import { entitySetOfType, refOf } from "./settlements-aging.js";
import { WAREHOUSE_BATCH, qty } from "./stock.js";

/**
 * Продажи по номенклатуре в УТ 11, КА 2 и ERP 2: регистр «Выручка и себестоимость продаж». Номенклатура в нём — через
 * ключ аналитики, и $expand даёт её тем же запросом: выгружать справочник ключей (сотни тысяч строк) не нужно.
 * Возвраты и сторно приходят отрицательным количеством, поэтому нетто — простая сумма.
 */

export const SALES_REGISTER = "AccumulationRegister_ВыручкаИСебестоимостьПродаж";
const SALES_LABEL = "Регистр накопления «Выручка и себестоимость продаж»";
const KEY_NAV = "АналитикаУчетаНоменклатуры";
export const RETAIL_REPORTS = ["Document_ОтчетОРозничныхПродажах"] as const;

/**
 * Обороты за [from, to] включительно. EndPeriod в 1С включает границу (проверено: документ ровно в 00:00:00 попадает в
 * период, который этим моментом кончается), поэтому конец — to 23:59:59, а не начало следующего дня: иначе документ
 * в полночь попал бы в оба соседних периода. Секунда — шаг времени 1С, конец дня не теряется.
 */
export function turnoversPath(register: string, from: string, to: string, condition?: string): string {
  const params = [
    `StartPeriod=datetime'${from}T00:00:00'`,
    `EndPeriod=datetime'${to}T23:59:59'`,
    ...(condition ? [`Condition='${condition.replace(/'/g, "''")}'`] : []),
  ];
  return `${register}/Turnovers(${params.join(",")})`;
}

/** Проверка периода: даты существуют, начало не позже конца. */
export function checkPeriod(from: string, to: string): void {
  nextDay(from);
  nextDay(to);
  if (from > to) throw new InputError(`Начало периода ${from} позже конца ${to}.`);
}

const SELECT = [
  `${KEY_NAV}_Key`,
  `${KEY_NAV}/Номенклатура_Key`,
  `${KEY_NAV}/Характеристика_Key`,
  "Склад",
  "Склад_Type",
  "ХозяйственнаяОперация",
  "КоличествоTurnover",
];

/** Строки оборотов; при складах — пачками, отбор на стороне 1С. Пачки не пересекаются, строки не задваиваются. */
export async function readItemSales(
  conn: Connection,
  {
    from,
    to,
    warehouses,
    warehouseSet,
  }: {
    from: string;
    to: string;
    warehouses?: readonly string[] | undefined;
    warehouseSet?: string | undefined;
  },
): Promise<ODataEntity[]> {
  const register = await requireEntity(conn, [SALES_REGISTER], SALES_LABEL);
  const read = async (condition?: string) =>
    (
      await fetchAllForAggregation(
        conn,
        turnoversPath(register, from, to, condition),
        { select: SELECT, expand: [KEY_NAV] },
        `${from}…${to}`,
      )
    ).rows;
  if (!warehouses?.length) return read();
  // Склад в регистре составной: сравнение только через cast, иначе 1С отвечает «несовместимые типы».
  const type = (warehouseSet ?? "Catalog_Склады").replace(/'/g, "");
  const unique = [...new Set(warehouses.map((w) => w.toLowerCase()))];
  const rows: ODataEntity[] = [];
  for (let i = 0; i < unique.length; i += WAREHOUSE_BATCH) {
    const batch = unique.slice(i, i + WAREHOUSE_BATCH);
    rows.push(
      ...(await read(or(...batch.map((w) => cmp("Склад", "eq", `cast(${odataGuid(w)}, '${type}')`))))),
    );
  }
  return rows;
}

/** Строка продаж до свёртки: одна операция, один склад. */
export interface SalesRow {
  itemRef?: string;
  characteristicRef?: string;
  warehouseRef?: string;
  operation: string;
  quantity: number;
}

/** Склад строки — только если это склад из справочника складов; «Неопределено» и прочие типы — без склада. */
export function toSalesRows(rows: readonly ODataEntity[], warehouseSets: readonly string[]): SalesRow[] {
  return rows.map((r) => {
    const key = (r[KEY_NAV] ?? {}) as ODataEntity;
    const itemRef = refOf(key["Номенклатура_Key"]);
    const characteristicRef = refOf(key["Характеристика_Key"]);
    const set = entitySetOfType(r["Склад_Type"]);
    const warehouseRef = set && warehouseSets.includes(set) ? refOf(r["Склад"]) : undefined;
    return {
      ...(itemRef ? { itemRef } : {}),
      ...(characteristicRef ? { characteristicRef } : {}),
      ...(warehouseRef ? { warehouseRef } : {}),
      operation: String(r["ХозяйственнаяОперация"] ?? ""),
      quantity: num(r["КоличествоTurnover"]),
    };
  });
}

export interface SalesLine {
  itemRef?: string;
  characteristicRef?: string;
  warehouseRef?: string;
  /** Нетто: продано минус возвращено. */
  quantity: number;
  sold: number;
  returns: number;
}

/** Свёртка по номенклатуре, характеристике и, при byWarehouse, складу. Нулевое нетто без движений не показывается. */
export function aggregateSales(rows: readonly SalesRow[], byWarehouse: boolean): SalesLine[] {
  const lines = new Map<string, SalesLine>();
  for (const r of rows) {
    const warehouseRef = byWarehouse ? r.warehouseRef : undefined;
    const key = `${r.itemRef ?? ""}|${r.characteristicRef ?? ""}|${warehouseRef ?? ""}`;
    const line = lines.get(key) ?? {
      ...(r.itemRef ? { itemRef: r.itemRef } : {}),
      ...(r.characteristicRef ? { characteristicRef: r.characteristicRef } : {}),
      ...(warehouseRef ? { warehouseRef } : {}),
      quantity: 0,
      sold: 0,
      returns: 0,
    };
    if (r.quantity >= 0) line.sold += r.quantity;
    else line.returns -= r.quantity;
    lines.set(key, line);
  }
  return [...lines.values()]
    .map((l) => ({ ...l, sold: qty(l.sold), returns: qty(l.returns), quantity: qty(l.sold - l.returns) }))
    .filter((l) => l.sold !== 0 || l.returns !== 0);
}

/** Страница чтения отчётов о розничных продажах: две колонки, за год их десятки тысяч. */
const REPORTS_PAGE = 5000;

/**
 * Дата последнего проведённого отчёта о розничных продажах по каждому складу в периоде. Розница попадает в регистр
 * продаж только отчётом о продажах после закрытия смены: продажи позже этой даты по складу ещё не видны. undefined —
 * отчёты не опубликованы или в них нет склада и даты.
 */
export async function lastRetailReports(
  conn: Connection,
  from: string,
  to: string,
): Promise<Map<string, string> | undefined> {
  const set = resolveEntity(RETAIL_REPORTS, await conn.available());
  if (!set) return undefined;
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));
  if (!props.has("Склад_Key") || !props.has("Date") || !props.has("Posted")) return undefined;
  const { rows } = await fetchAllForAggregation(
    conn,
    set,
    {
      filter: and(
        cmp("Posted", "eq", "true"),
        cmp("Date", "ge", `datetime'${from}T00:00:00'`),
        cmp("Date", "le", `datetime'${to}T23:59:59'`),
      ),
      select: ["Ref_Key", "Склад_Key", "Date"],
    },
    `${from}…${to}`,
    REPORTS_PAGE,
  );
  const last = new Map<string, string>();
  for (const r of rows) {
    const wh = refOf(r["Склад_Key"]);
    const date = typeof r["Date"] === "string" ? r["Date"].slice(0, 10) : undefined;
    if (wh && date && (last.get(wh) ?? "") < date) last.set(wh, date);
  }
  return last;
}
