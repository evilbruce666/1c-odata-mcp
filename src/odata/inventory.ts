import type { Connection } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { num, readByRefs, resolveNames, GUID_RE } from "./accounting.js";
import { cmp } from "./query.js";
import { entitySetOfType, refOf } from "./settlements-aging.js";

/**
 * Остатки по счетам 41/10/43. Субконто идут по порядку плана счетов, и позиция у счетов разная: на 41.01
 * Номенклатура | Партии, на 41.11 Номенклатура | Склады | Партии, на 41.12 только Склады, на 10.07
 * Контрагенты | Номенклатура | Партии. Поэтому значение ищется по типу (ExtDimensionN_Type), а не по номеру.
 */

const SUBCONTO_POSITIONS = [1, 2, 3] as const;

/** Значение субконто из указанных справочников, на какой бы позиции оно ни стояло. */
export function dimensionOf(
  row: ODataEntity,
  sets: readonly string[],
): { ref: string; set: string } | undefined {
  for (const i of SUBCONTO_POSITIONS) {
    const set = entitySetOfType(row[`ExtDimension${i}_Type`]);
    const ref = refOf(row[`ExtDimension${i}`]);
    if (set && ref && sets.includes(set)) return { ref, set };
  }
  return undefined;
}

export interface StockLine {
  itemRef?: string;
  warehouseRef?: string;
  warehouseSet?: string;
  quantity: number;
  cents: number;
}

/** Сводит строки сальдо по номенклатуре и, при byWarehouse, складу. Копейки — целые, количество — float. */
export function aggregateStock(rows: readonly ODataEntity[], byWarehouse: boolean): StockLine[] {
  const lines = new Map<string, StockLine>();
  for (const r of rows) {
    const item = dimensionOf(r, CATALOGS.nomenclature);
    const warehouse = byWarehouse ? dimensionOf(r, CATALOGS.warehouses) : undefined;
    const key = `${item?.ref ?? ""}|${warehouse?.ref ?? ""}`;
    const line = lines.get(key) ?? {
      ...(item ? { itemRef: item.ref } : {}),
      ...(warehouse ? { warehouseRef: warehouse.ref, warehouseSet: warehouse.set } : {}),
      quantity: 0,
      cents: 0,
    };
    line.quantity += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
    line.cents += Math.round(num(r["СуммаBalanceDr"]) * 100) - Math.round(num(r["СуммаBalanceCr"]) * 100);
    lines.set(key, line);
  }
  return [...lines.values()].filter((l) => Math.abs(l.quantity) > 0.0001 || l.cents !== 0);
}

export interface ItemGroup {
  ref: string;
  name: string;
  parent?: string;
}

/** Все группы (папки) справочника номенклатуры. */
export async function loadItemGroups(conn: Connection, set: string): Promise<ItemGroup[]> {
  const { rows } = await fetchAllForAggregation(
    conn,
    set,
    { filter: cmp("IsFolder", "eq", "true"), select: ["Ref_Key", "Description", "Parent_Key"] },
    "группы номенклатуры",
  );
  return rows.map((r) => {
    const parent = refOf(r["Parent_Key"]);
    return { ref: String(r["Ref_Key"]), name: String(r["Description"] ?? ""), ...(parent ? { parent } : {}) };
  });
}

/** Группа по Ref_Key или названию: точное совпадение важнее части, неоднозначность — отказ, а не выбор наугад. */
export function resolveItemGroup(query: string, groups: readonly ItemGroup[]): ItemGroup {
  const q = query.trim();
  const lower = q.toLowerCase();
  const byRef = GUID_RE.test(q) ? groups.find((g) => g.ref.toLowerCase() === lower) : undefined;
  const exact = groups.filter((g) => g.name.toLowerCase() === lower);
  const found = byRef
    ? [byRef]
    : exact.length
      ? exact
      : groups.filter((g) => g.name.toLowerCase().includes(lower));
  if (found.length === 1) return found[0]!;
  if (!found.length) throw new InputError(`Группа номенклатуры "${query}" не найдена.`);
  const shown = found
    .slice(0, 10)
    .map((g) => `${g.name} (${g.ref})`)
    .join(", ");
  throw new InputError(`Под "${query}" подходит несколько групп: ${shown}. Укажите Ref_Key нужной.`);
}

/** Группа и все её подгруппы на любой глубине. Цикл в данных не зациклит обход. */
export function subtreeOf(root: string, groups: readonly ItemGroup[]): Set<string> {
  const seen = new Set([root]);
  for (const queue = [root]; queue.length;) {
    const parent = queue.pop()!;
    for (const g of groups) {
      if (g.parent === parent && !seen.has(g.ref)) {
        seen.add(g.ref);
        queue.push(g.ref);
      }
    }
  }
  return seen;
}

export interface ItemAttributes {
  name?: string;
  article?: string;
  groupRef?: string;
}

/** Название, артикул и родитель. Поля, которых нет в схеме базы, не запрашиваются: иначе 1С ответит 400. */
export async function readItemAttributes(
  conn: Connection,
  set: string,
  refs: Iterable<string>,
): Promise<Map<string, ItemAttributes>> {
  const meta = await conn.getMetadata();
  const props = new Set(meta.entities.get(set)?.properties.map((p) => p.name));
  const select = ["Ref_Key", "Description", ...["Артикул", "Parent_Key"].filter((f) => props.has(f))];
  const rows = await readByRefs(conn, set, refs, select);
  const out = new Map<string, ItemAttributes>();
  for (const [ref, r] of rows) {
    const article = typeof r["Артикул"] === "string" ? r["Артикул"].trim() : "";
    const groupRef = refOf(r["Parent_Key"]);
    out.set(ref, {
      ...(r["Description"] ? { name: String(r["Description"]) } : {}),
      ...(article ? { article } : {}),
      ...(groupRef ? { groupRef } : {}),
    });
  }
  return out;
}

/** Названия складов: справочник берётся из типа значения, в разных конфигурациях он называется по-разному. */
export async function warehouseNames(
  conn: Connection,
  lines: readonly StockLine[],
): Promise<Map<string, string>> {
  const bySet = new Map<string, Set<string>>();
  for (const l of lines) {
    if (l.warehouseRef && l.warehouseSet) {
      bySet.set(l.warehouseSet, (bySet.get(l.warehouseSet) ?? new Set()).add(l.warehouseRef));
    }
  }
  const out = new Map<string, string>();
  for (const [set, refs] of bySet) {
    for (const [ref, name] of await resolveNames(conn, set, refs)) out.set(ref, name);
  }
  return out;
}

export const nomenclatureSet = async (conn: Connection): Promise<string | undefined> =>
  resolveEntity(CATALOGS.nomenclature, await conn.available());
