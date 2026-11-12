import type { Connection } from "../context.js";
import { accountingRegister } from "../config/accounting.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { nextDay, num, readByRefs } from "./accounting.js";
import { NotPublishedError, requireEntity } from "./publication.js";
import { cmp, odataGuid, or } from "./query.js";
import { refOf } from "./settlements-aging.js";

/**
 * Остатки по складам в конфигурациях с регистром накопления «Товары на складах» (УТ 11, КА 2, ERP 2). Там нет
 * Хозрасчётного, а в регистре только количество: сумм нет, и инструмент их не выводит.
 */

export const STOCK_REGISTER = "AccumulationRegister_ТоварыНаСкладах";
const STOCK_LABEL = "Регистр накопления «Товары на складах»";
export const CHARACTERISTICS = ["Catalog_ХарактеристикиНоменклатуры"] as const;

/** Складов в одном Condition: длинное «или» упирается в лимит вложенности условий 1С. */
export const WAREHOUSE_BATCH = 10;

/**
 * Отказ остатков по счетам, если в базе нет Хозрасчётного, но есть регистр склада: совет опубликовать Хозрасчётный
 * в УТ неверен, его нет в самой конфигурации.
 */
export async function stockInsteadOfAccounts(conn: Connection): Promise<void> {
  const available = await conn.available();
  const registers = accountingRegister(conn.cfg);
  if (resolveEntity(registers, available) || !available.has(STOCK_REGISTER)) return;
  throw new NotPublishedError(
    [{ label: "Регистр бухгалтерии «Хозрасчётный»", candidates: registers }],
    "В этой базе нет регистра бухгалтерии «Хозрасчётный», но есть регистр «Товары на складах»: похоже, это " +
      "УТ, КА или ERP, а не Бухгалтерия. Остатки по складам (количество, без сумм) — read.analytics.get_stock.",
  );
}

/** Значение для параметра виртуальной таблицы в кавычках: кавычка внутри удваивается, как в 1С. */
const quoted = (text: string) => `'${text.replace(/'/g, "''")}'`;

/**
 * Путь к остаткам. Period — начало следующего дня, как у Хозрасчётного: остаток на конец asOf без потери движений
 * последней секунды. Condition — отбор на стороне 1С, без него она отдала бы остатки всех складов.
 */
export function balancePath(register: string, asOf?: string, condition?: string): string {
  const params = [
    ...(asOf ? [`Period=datetime'${nextDay(asOf)}T00:00:00'`] : []),
    ...(condition ? [`Condition=${quoted(condition)}`] : []),
  ];
  return params.length ? `${register}/Balance(${params.join(",")})` : `${register}/Balance`;
}

const SELECT = ["Номенклатура_Key", "Характеристика_Key", "Склад_Key", "ВНаличииBalance", "КОтгрузкеBalance"];

/** Строки остатков; при складах — пачками по WAREHOUSE_BATCH. Пачки не пересекаются, строки не задваиваются. */
export async function readStockBalance(
  conn: Connection,
  { asOf, warehouses }: { asOf?: string | undefined; warehouses?: readonly string[] | undefined },
): Promise<ODataEntity[]> {
  const register = await requireEntity(conn, [STOCK_REGISTER], STOCK_LABEL);
  const read = async (condition?: string) =>
    (
      await fetchAllForAggregation(
        conn,
        balancePath(register, asOf, condition),
        { select: SELECT },
        `остатки на ${asOf ?? "сейчас"}`,
      )
    ).rows;
  if (!warehouses?.length) return read();
  const unique = [...new Set(warehouses.map((w) => w.toLowerCase()))];
  const rows: ODataEntity[] = [];
  for (let i = 0; i < unique.length; i += WAREHOUSE_BATCH) {
    const batch = unique.slice(i, i + WAREHOUSE_BATCH);
    rows.push(...(await read(or(...batch.map((w) => cmp("Склад_Key", "eq", odataGuid(w)))))));
  }
  return rows;
}

export interface WarehouseStockLine {
  itemRef?: string;
  characteristicRef?: string;
  warehouseRef?: string;
  inStock: number;
  toShip: number;
}

/** Количество с точностью до тысячных: float-сумма тысяч строк иначе даёт 2.9999999. */
export const qty = (n: number): number => Math.round(n * 1000) / 1000;

/**
 * Сводит строки по номенклатуре, характеристике и, при byWarehouse, складу. Назначение, помещение, серия и
 * контролируемая партия сворачиваются: это разрез внутри склада, а не другой товар.
 */
export function aggregateWarehouseStock(
  rows: readonly ODataEntity[],
  byWarehouse: boolean,
): WarehouseStockLine[] {
  const lines = new Map<string, WarehouseStockLine>();
  for (const r of rows) {
    const itemRef = refOf(r["Номенклатура_Key"]);
    const characteristicRef = refOf(r["Характеристика_Key"]);
    const warehouseRef = byWarehouse ? refOf(r["Склад_Key"]) : undefined;
    const key = `${itemRef ?? ""}|${characteristicRef ?? ""}|${warehouseRef ?? ""}`;
    const line = lines.get(key) ?? {
      ...(itemRef ? { itemRef } : {}),
      ...(characteristicRef ? { characteristicRef } : {}),
      ...(warehouseRef ? { warehouseRef } : {}),
      inStock: 0,
      toShip: 0,
    };
    line.inStock += num(r["ВНаличииBalance"]);
    line.toShip += num(r["КОтгрузкеBalance"]);
    lines.set(key, line);
  }
  return [...lines.values()]
    .map((l) => ({ ...l, inStock: qty(l.inStock), toShip: qty(l.toShip) }))
    .filter((l) => l.inStock !== 0 || l.toShip !== 0);
}

export interface Warehouse {
  ref: string;
  name: string;
  type?: string;
}

/** Не больше складов за вызов: отбор идёт пачками по WAREHOUSE_BATCH, каждая — запрос к 1С. */
export const MAX_WAREHOUSES = 100;

/**
 * Склады отбора: каждый должен существовать и не быть группой. Иначе отбор дал бы тихий ноль, а ответ не отличался
 * бы от «товара нет».
 */
export async function readWarehouses(conn: Connection, refs: readonly string[]): Promise<Warehouse[]> {
  const set = await requireEntity(conn, CATALOGS.warehouses, "Справочник «Склады»");
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));
  const select = ["Ref_Key", "Description", ...["ТипСклада", "IsFolder"].filter((f) => props.has(f))];
  const found = await readByRefs(conn, set, refs, select);
  const byLower = new Map([...found].map(([ref, row]) => [ref.toLowerCase(), row]));
  const missing = refs.filter((r) => !byLower.has(r.toLowerCase()));
  if (missing.length) {
    throw new InputError(
      `Склады не найдены: ${missing.join(", ")}. Ref_Key складов есть в ответе с groupBy=warehouse (byWarehouse).`,
    );
  }
  const folders = refs.filter((r) => byLower.get(r.toLowerCase())?.["IsFolder"] === true);
  if (folders.length) {
    throw new InputError(
      `Это группы складов, а не склады: ${folders.join(", ")}. Перечислите сами склады: остатки ведутся по ним.`,
    );
  }
  return refs.map((ref) => {
    const row = byLower.get(ref.toLowerCase())!;
    const type = typeof row["ТипСклада"] === "string" && row["ТипСклада"] ? row["ТипСклада"] : undefined;
    return {
      ref: String(row["Ref_Key"]),
      name: String(row["Description"] ?? ref),
      ...(type ? { type } : {}),
    };
  });
}

/** Склады по Ref_Key для подписей строк: название и тип. Неопубликованный справочник — пустой ответ, а не отказ. */
export async function describeWarehouses(
  conn: Connection,
  refs: Iterable<string>,
): Promise<Map<string, Omit<Warehouse, "ref">>> {
  const set = resolveEntity(CATALOGS.warehouses, await conn.available());
  if (!set) return new Map();
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));
  const select = ["Ref_Key", "Description", ...(props.has("ТипСклада") ? ["ТипСклада"] : [])];
  const rows = await readByRefs(conn, set, refs, select);
  return new Map(
    [...rows].map(([ref, r]) => [
      ref,
      {
        name: String(r["Description"] ?? ref),
        ...(typeof r["ТипСклада"] === "string" && r["ТипСклада"] ? { type: r["ТипСклада"] } : {}),
      },
    ]),
  );
}
