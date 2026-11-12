import type { Connection } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { AggregateOverflowError, fetchAllForAggregation } from "./aggregate.js";
import { callPageStart } from "./call-context.js";
import { cached } from "./cache.js";
import { num, readByRefs, resolveNames, GUID_RE } from "./accounting.js";
import { and, cmp, odataGuid, or } from "./query.js";
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

/** Отдаёт ли 1С типы субконто. Старые конфигурации их не отдают: тогда склад определить нельзя. */
export const hasSubcontoTypes = (row: ODataEntity): boolean =>
  SUBCONTO_POSITIONS.some((i) => `ExtDimension${i}_Type` in row);

/** Сводит строки сальдо по номенклатуре и, при byWarehouse, складу. Копейки — целые, количество — float. */
export function aggregateStock(rows: readonly ODataEntity[], byWarehouse: boolean): StockLine[] {
  const lines = new Map<string, StockLine>();
  for (const r of rows) {
    // Старые конфигурации отдают субконто без _Type: номенклатура там определялась первой позицией, и других данных нет.
    const typed = hasSubcontoTypes(r);
    const legacyRef = typed ? undefined : refOf(r["ExtDimension1"]);
    const item = typed
      ? dimensionOf(r, CATALOGS.nomenclature)
      : legacyRef
        ? { ref: legacyRef, set: "" }
        : undefined;
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
export function loadItemGroups(conn: Connection, set: string): Promise<ItemGroup[]> {
  return readCatalog(() =>
    cached(
      conn,
      `inventory-groups:${set}`,
      async () => {
        const { rows } = await fetchAllForAggregation(
          conn,
          set,
          { filter: cmp("IsFolder", "eq", "true"), select: ["Ref_Key", "Description", "Parent_Key"] },
          "группы номенклатуры",
          CATALOG_PAGE,
        );
        return rows.map((r) => {
          const parent = refOf(r["Parent_Key"]);
          return {
            ref: String(r["Ref_Key"]),
            name: String(r["Description"] ?? ""),
            ...(parent ? { parent } : {}),
          };
        });
      },
      itemsTtl(),
    ),
  );
}

/** Разделитель пути группы: «Телефоны/Смартфоны/Apple». */
const PATH_SEPARATOR = "/";

/** Цепочка групп от корня до указанной. Цикл в данных обрывает цепочку, а не зацикливает её. */
function ancestry(group: ItemGroup, byRef: ReadonlyMap<string, ItemGroup>): ItemGroup[] {
  const chain = [group];
  const seen = new Set([group.ref]);
  for (let p = group.parent; p && !seen.has(p);) {
    const parent = byRef.get(p);
    if (!parent) break;
    chain.unshift(parent);
    seen.add(p);
    p = parent.parent;
  }
  return chain;
}

/** Путь группы от корня: по нему различаются одноимённые группы в разных ветках. */
export function groupPath(group: ItemGroup, groups: readonly ItemGroup[]): string {
  return pathIn(group, new Map(groups.map((g) => [g.ref, g])));
}

const pathIn = (group: ItemGroup, byRef: ReadonlyMap<string, ItemGroup>): string =>
  ancestry(group, byRef)
    .map((g) => g.name.trim())
    .join(PATH_SEPARATOR);

/**
 * Группы, путь которых заканчивается указанными частями: «Смартфоны/Apple» находит Apple внутри любых «Смартфонов».
 * Части сравниваются целиком, сначала буквально, похожие буквы — только если буквально ничего нет.
 */
function byPath(parts: readonly string[], groups: readonly ItemGroup[]): ItemGroup[] {
  const byRef = new Map(groups.map((g) => [g.ref, g]));
  const matching = (fold: (text: string) => string) =>
    groups.filter((g) => {
      const chain = ancestry(g, byRef);
      if (chain.length < parts.length) return false;
      const tail = chain.slice(chain.length - parts.length);
      return tail.every((link, i) => fold(link.name.trim()) === fold(parts[i]!));
    });
  const literal = matching(strictText);
  return literal.length ? literal : matching(foldText);
}

/**
 * Группа по Ref_Key, названию или пути: точное совпадение важнее части, неоднозначность — отказ со списком путей, а не
 * выбор наугад. Название целиком проверяется раньше пути: в названии бывает «/» («Кабели/переходники»).
 */
export function resolveItemGroup(query: string, groups: readonly ItemGroup[]): ItemGroup {
  const q = query.trim();
  const strict = strictText(q);
  const named = (g: ItemGroup) => strictText(g.name.trim());
  const byRef = GUID_RE.test(q) ? groups.find((g) => g.ref.toLowerCase() === q.toLowerCase()) : undefined;
  const exact = groups.filter((g) => named(g) === strict);
  // Похожие буквы только для названия целиком: по части «TP» выбрал бы «Стройматериалы», итог был бы чужой группы.
  const lookalike = groups.filter((g) => foldText(g.name.trim()) === foldText(q));
  const parts = q
    .split(PATH_SEPARATOR)
    .map((p) => p.trim())
    .filter(Boolean);
  const path = parts.length > 1 ? byPath(parts, groups) : [];
  const part = groups.filter((g) => named(g).includes(strict));
  const found = byRef
    ? [byRef]
    : exact.length
      ? exact
      : lookalike.length
        ? lookalike
        : path.length
          ? path
          : part;
  if (found.length === 1) return found[0]!;
  if (!found.length) throw new InputError(`Группа номенклатуры "${query}" не найдена.`);
  const index = new Map(groups.map((g) => [g.ref, g]));
  const shown = found
    .map((g) => `${pathIn(g, index)} (${g.ref})`)
    .sort()
    .slice(0, 10)
    .join("; ");
  const more = found.length > 10 ? ` и ещё ${found.length - 10}` : "";
  throw new InputError(
    `Под "${query}" подходит несколько групп (${found.length}): ${shown}${more}. ` +
      `Укажите путь, например «Родитель${PATH_SEPARATOR}Группа», или Ref_Key нужной.`,
  );
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

/** Поля, которых нет в схеме базы, не запрашиваются: иначе 1С ответит 400. */
async function itemFields(conn: Connection, set: string): Promise<string[]> {
  const meta = await conn.getMetadata();
  const props = new Set(meta.entities.get(set)?.properties.map((p) => p.name));
  return ["Ref_Key", "Description", ...["Артикул", "Parent_Key"].filter((f) => props.has(f))];
}

const toAttributes = (r: ODataEntity): ItemAttributes => {
  const article = typeof r["Артикул"] === "string" ? r["Артикул"].trim() : "";
  const groupRef = refOf(r["Parent_Key"]);
  return {
    ...(r["Description"] ? { name: String(r["Description"]) } : {}),
    ...(article ? { article } : {}),
    ...(groupRef ? { groupRef } : {}),
  };
};

/** Название, артикул и родитель указанных позиций. */
export async function readItemAttributes(
  conn: Connection,
  set: string,
  refs: Iterable<string>,
): Promise<Map<string, ItemAttributes>> {
  const rows = await readByRefs(conn, set, refs, await itemFields(conn, set));
  return new Map([...rows].map(([ref, r]) => [ref, toAttributes(r)]));
}

/**
 * Страница чтения справочника номенклатуры целиком: строки короткие (4 поля), 5000 — около 1 МБ. По 1000 справочник
 * на 15 тыс. позиций читался 15 запросами, около 15 с.
 */
const CATALOG_PAGE = 5000;

/** Сколько держать справочник для соседних страниц одного чтения: на каждую страницу он иначе читался бы заново. */
const ITEMS_TTL_MS = 60_000;

/**
 * Кэш только для продолжения чтения: первая страница читает свежий справочник, иначе только что заведённый товар или
 * группа дали бы ложный ноль, а остатки при этом свежие.
 */
const itemsTtl = (): number => (callPageStart() === 0 ? 0 : ITEMS_TTL_MS);

/** Отказ вместо технического «документов за номенклатура»: периода у справочника нет, сузить его нечем. */
async function readCatalog<T>(load: () => Promise<T>): Promise<T> {
  try {
    return await load();
  } catch (e) {
    if (!(e instanceof AggregateOverflowError)) throw e;
    throw new InputError(
      `В справочнике номенклатуры больше ${e.cap} строк: поиск и отбор по группе требуют прочитать его целиком. ` +
        "Увеличьте ODATA_ANALYTICS_MAX_ROWS или запросите остатки без отбора.",
    );
  }
}

/**
 * Вся номенклатура без групп. Поиск идёт здесь, а не через substringof: ему нужны е/ё и похожие буквы разных раскладок,
 * а 1С сравнивает строки буквально.
 */
export function readAllItems(conn: Connection, set: string): Promise<Map<string, ItemAttributes>> {
  return readCatalog(() =>
    cached(
      conn,
      `inventory-items:${set}`,
      async () => {
        const { rows } = await fetchAllForAggregation(
          conn,
          set,
          { filter: cmp("IsFolder", "eq", "false"), select: await itemFields(conn, set) },
          "номенклатура",
          CATALOG_PAGE,
        );
        return new Map(rows.map((r) => [String(r["Ref_Key"]), toAttributes(r)]));
      },
      itemsTtl(),
    ),
  );
}

/** Опубликован ли артикул: иначе поиск идёт только по названию, и ответ должен это сказать. */
export async function articlePublished(conn: Connection, set: string): Promise<boolean> {
  return (await itemFields(conn, set)).includes("Артикул");
}

const LOOKALIKE: Record<string, string> = {
  а: "a",
  в: "b",
  е: "e",
  ё: "e",
  к: "k",
  м: "m",
  н: "h",
  о: "o",
  р: "p",
  с: "c",
  т: "t",
  у: "y",
  х: "x",
};

/** Текст для сравнения: без регистра, ё как е. */
export const strictText = (text: string): string => text.toLowerCase().replace(/ё/g, "е");

/** То же, и похожие буквы кириллицы и латиницы одинаковы («А4» и «A4»). Даёт ложные совпадения, поэтому только вторым шагом. */
export const foldText = (text: string): string =>
  strictText(text).replace(/[авекмнорстух]/g, (c) => LOOKALIKE[c] ?? c);

/** Все части запроса есть в названии или артикуле: «бумаг а4» находит «Бумага офисная (A4)». */
export function matchesItemQuery(item: ItemAttributes, query: string, fold = foldText): boolean {
  const text = fold(`${item.name ?? ""} ${item.article ?? ""}`);
  return fold(query)
    .split(/\s+/)
    .filter(Boolean)
    .every((word) => text.includes(word));
}

/** Групп в одном $filter: длинный or упирается в лимит вложенности условий 1С. */
const GROUP_FILTER_BATCH = 10;

/**
 * Больше стольких групп — справочник читается целиком и отбирается в памяти: на УТ с 1236 папками в ветке пачки по
 * 10 давали 124 запроса и 53 с, а весь справочник на 15 тыс. позиций — 3 страницы.
 */
const WHOLE_CATALOG_FROM = GROUP_FILTER_BATCH * 3;

/**
 * Позиции указанных групп. Запрос идёт по группам, а не по остаткам: иначе отбор на 10 000 позициях стоил бы
 * 500 запросов на каждую страницу.
 */
export async function readItemsInGroups(
  conn: Connection,
  set: string,
  groupRefs: Iterable<string>,
): Promise<Map<string, ItemAttributes>> {
  const refs = [...groupRefs].sort();
  if (refs.length > WHOLE_CATALOG_FROM) {
    const wanted = new Set(refs);
    const all = await readAllItems(conn, set);
    return new Map([...all].filter(([, a]) => !!a.groupRef && wanted.has(a.groupRef)));
  }
  return readCatalog(() =>
    cached(
      conn,
      `inventory-items-in:${set}:${refs.join(",")}`,
      async () => {
        const select = await itemFields(conn, set);
        const out = new Map<string, ItemAttributes>();
        for (let i = 0; i < refs.length; i += GROUP_FILTER_BATCH) {
          const { rows } = await fetchAllForAggregation(
            conn,
            set,
            {
              filter: and(
                or(
                  ...refs
                    .slice(i, i + GROUP_FILTER_BATCH)
                    .map((ref) => cmp("Parent_Key", "eq", odataGuid(ref))),
                ),
                cmp("IsFolder", "eq", "false"),
              ),
              select,
            },
            "номенклатура группы",
          );
          for (const r of rows) out.set(String(r["Ref_Key"]), toAttributes(r));
        }
        return out;
      },
      itemsTtl(),
    ),
  );
}

/** Результат отбора строк остатков по товару: общий для остатков по счетам (БП) и по складам (УТ). */
export interface ItemSelection<L> {
  lines: L[];
  /** Название, артикул и родитель прочитанных позиций; без отбора пусто, его заполняют для страницы. */
  attributes: Map<string, ItemAttributes>;
  groups?: ItemGroup[];
  selectedGroup?: ItemGroup;
  /** Найдено только с учётом похожих букв кириллицы и латиницы. */
  lookalikeMatch: boolean;
  /** Позиции, похожие на запрос другой раскладкой букв, при буквальных совпадениях: в итог не входят. */
  otherSimilar: L[];
}

/** Отбор строк по группе номенклатуры (с подгруппами) и поиску по названию или артикулу. */
export async function selectByItem<L extends { itemRef?: string | undefined }>(
  conn: Connection,
  nomSet: string | undefined,
  rows: L[],
  { group, query }: { group?: string | undefined; query?: string | undefined },
): Promise<ItemSelection<L>> {
  const attributes = new Map<string, ItemAttributes>();
  let lines = rows;
  let groups: ItemGroup[] | undefined;
  let selectedGroup: ItemGroup | undefined;
  let lookalikeMatch = false;
  let otherSimilar: L[] = [];
  if (group) {
    if (!nomSet) throw new InputError("Справочник номенклатуры не опубликован: отбор по группе невозможен.");
    groups = await loadItemGroups(conn, nomSet);
    selectedGroup = resolveItemGroup(group, groups);
    const allowed = subtreeOf(selectedGroup.ref, groups);
    for (const [ref, a] of await readItemsInGroups(conn, nomSet, allowed)) attributes.set(ref, a);
    lines = lines.filter((l) => !!l.itemRef && attributes.has(l.itemRef));
  }
  if (query) {
    if (!nomSet) throw new InputError("Справочник номенклатуры не опубликован: поиск по товару невозможен.");
    if (!group) for (const [ref, a] of await readAllItems(conn, nomSet)) attributes.set(ref, a);
    const stock = lines;
    const matching = (fold: (text: string) => string) =>
      stock.filter((l) => {
        const item = l.itemRef ? attributes.get(l.itemRef) : undefined;
        return !!item && matchesItemQuery(item, query, fold);
      });
    // Сначала буквально; похожие буквы только если ничего нет: «3M» иначе находил бы и «Кабель 3м».
    const literal = matching(strictText);
    const similar = matching(foldText);
    lines = literal.length ? literal : similar;
    lookalikeMatch = !literal.length && lines.length > 0;
    // «А4» и «A4» в одной базе: итог только по буквальным занизил бы остаток, поэтому остальные названы.
    otherSimilar = literal.length ? similar.filter((l) => !literal.includes(l)) : [];
  }
  return {
    lines,
    attributes,
    ...(groups ? { groups } : {}),
    ...(selectedGroup ? { selectedGroup } : {}),
    lookalikeMatch,
    otherSimilar,
  };
}

/** Примечания к отбору по товару. absent — чего нет у найденного товара: «остатка на складах», «продаж за период». */
export async function selectionNotes(
  conn: Connection,
  nomSet: string | undefined,
  sel: ItemSelection<unknown>,
  { group, query }: { group?: string | undefined; query?: string | undefined },
  absent: string,
): Promise<string[]> {
  const notes: string[] = [];
  if (query && nomSet && !(await articlePublished(conn, nomSet))) {
    notes.push("Артикул не опубликован в OData: поиск шёл только по названию.");
  }
  if (sel.otherSimilar.length) {
    notes.push(
      "Ещё есть позиции, похожие на запрос другой раскладкой букв (А и A): в итог они не вошли, см. otherSimilar. Чтобы показать их, повторите поиск теми же буквами, что в их названии.",
    );
  }
  if (sel.lookalikeMatch) {
    notes.push(
      "Точных совпадений нет: найдено с учётом похожих букв кириллицы и латиницы, проверьте позиции.",
    );
  }
  if (query && !sel.lines.length) {
    const scope = group ? "в выбранной группе" : "в справочнике";
    const items = [...sel.attributes.values()];
    const known = items.some((a) => matchesItemQuery(a, query, strictText));
    const similar = !known && items.some((a) => matchesItemQuery(a, query));
    notes.push(
      known
        ? `Товар есть ${scope}, но ${absent} нет (с учётом других отборов).`
        : similar
          ? `Товар с похожими буквами другой раскладки (А и A) есть ${scope}, но ${absent} нет (с учётом других отборов).`
          : group
            ? "В выбранной группе такого товара нет; вне группы поиск не проверял."
            : "В справочнике такого товара не найдено. Попробуйте основу слова (бумаг, а не бумаги) или другую часть названия.",
    );
  }
  return notes;
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
