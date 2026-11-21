import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { EMPTY_REF } from "../odata/pricing.js";
import { buildQuery, cmp, odataGuid, or } from "../odata/query.js";
import { InputError } from "../errors.js";
import { withoutAccounting } from "../odata/configuration.js";
import { createCore, docDate, resolveCatalogItem, resolveOrgRef } from "./write.js";

/**
 * Складские документы Управления торговлей: перемещение, оприходование излишков, списание недостач, пересчёт.
 * Проверено на копии рабочей УТ 11.5 (21.11.2026): документ с минимальным набором полей 1С принимает, а вид запасов,
 * ключ аналитики номенклатуры и идентификаторы строк заполняет сама при проведении. Проверка проведения — по
 * регистру «Товары на складах»: движения самого документа совпадают с его строками.
 */

export const STOCK_DOCS = {
  transfer: "Document_ПеремещениеТоваров",
  surplus: "Document_ОприходованиеИзлишковТоваров",
  writeoff: "Document_СписаниеНедостачТоваров",
  inventory: "Document_ПересчетТоваров",
} as const;
export type StockKind = keyof typeof STOCK_DOCS;
const STOCK_REGISTER = "AccumulationRegister_ТоварыНаСкладах";
const ITEMS = "Catalog_Номенклатура";
const KINDS = "Catalog_ВидыНоменклатуры";
const EXPENSE_ITEMS = "ChartOfCharacteristicTypes_СтатьиРасходов";
const INCOME_ITEMS = "ChartOfCharacteristicTypes_СтатьиДоходов";

/**
 * Складской документ УТ для записи: база без регистра бухгалтерии (КА и ERP с «Хозрасчётным» — как раньше, они не
 * проверены) и документ устроен как в УТ. undefined — сборщик Бухгалтерии.
 */
export async function tradeStockSet(conn: Connection, kind: StockKind): Promise<string | undefined> {
  const set = STOCK_DOCS[kind];
  if (!(await withoutAccounting(conn))) return undefined;
  return (await tradeStockSchema(conn, set)) === kind ? set : undefined;
}

/** Поля, которые пишет сборщик УТ: документ подходит, если в схеме есть все. */
const HEADER: Record<StockKind, readonly string[]> = {
  transfer: [
    "Организация_Key",
    "СкладОтправитель_Key",
    "СкладПолучатель_Key",
    "ХозяйственнаяОперация",
    "Статус",
  ],
  surplus: [
    "Организация_Key",
    "Склад_Key",
    "ВидЦены_Key",
    "СтатьяДоходов_Key",
    "Валюта_Key",
    "ПересчетТоваров_Key",
  ],
  writeoff: ["Организация_Key", "Склад_Key", "СтатьяРасходов", "ПересчетТоваров_Key"],
  inventory: ["Склад_Key", "Статус", "УчетныеДанныеЗаполнены"],
};
const ROWS: Record<StockKind, readonly string[]> = {
  transfer: ["Номенклатура_Key", "Количество"],
  surplus: ["Номенклатура_Key", "Количество", "Цена", "Сумма", "Серия_Key"],
  writeoff: ["Номенклатура_Key", "Количество", "Серия_Key"],
  inventory: ["Номенклатура_Key", "Количество", "КоличествоФакт"],
};

/**
 * Документ устроен как в УТ: есть все поля сборщика УТ и нет счетов учёта Бухгалтерии в строках. Перемещение
 * называется одинаково в обеих конфигурациях — различает его только состав полей.
 */
export async function tradeStockSchema(conn: Connection, set: string): Promise<StockKind | undefined> {
  const kind = (Object.keys(STOCK_DOCS) as StockKind[]).find((k) => STOCK_DOCS[k] === set);
  if (!kind || !(await conn.available()).has(set)) return undefined;
  const meta = await conn.getMetadata();
  const head = new Set(meta.entities.get(set)?.properties.map((p) => p.name) ?? []);
  const rows = new Set(meta.entities.get(`${set}_Товары`)?.properties.map((p) => p.name) ?? []);
  if (rows.has("СчетУчета_Key")) return undefined;
  return HEADER[kind].every((f) => head.has(f)) && ROWS[kind].every((f) => rows.has(f)) ? kind : undefined;
}

export interface StockLine {
  nomenclatureRef: string;
  quantity: number;
  price?: number | undefined;
  seriesRef?: string | undefined;
}

const guidOf = (s: string) => s.replace(/[{}']/g, "").trim();
const round2 = (n: number) => Math.round(n * 100) / 100;

/**
 * Позиции, вид которых ведёт серии, а серия не указана. Не отказ: нужна ли серия, решает политика учёта серий на
 * складе (на копии рабочей базы у одного вида — разные политики по складам); откажет, если нужно, сама 1С при
 * проведении. Схема без видов номенклатуры — проверки нет.
 */
async function seriesNotes(conn: Connection, lines: readonly StockLine[]): Promise<string[]> {
  const without = [...new Set(lines.filter((l) => !l.seriesRef).map((l) => guidOf(l.nomenclatureRef)))];
  const available = await conn.available();
  if (!without.length || !available.has(KINDS)) return [];
  const meta = await conn.getMetadata();
  const has = (set: string, f: string) => !!meta.entities.get(set)?.properties.some((p) => p.name === f);
  if (!has(ITEMS, "ВидНоменклатуры_Key") || !has(KINDS, "ИспользоватьСерии")) return [];
  const items = await conn.client.getCollection<ODataEntity>(
    `${ITEMS}${buildQuery({
      filter: or(...without.map((r) => cmp("Ref_Key", "eq", odataGuid(r)))),
      select: ["Ref_Key", "Description", "ВидНоменклатуры_Key"],
    })}`,
  );
  const kinds = [...new Set(items.value.map((i) => String(i["ВидНоменклатуры_Key"] ?? "")))].filter(
    (k) => k && k !== EMPTY_REF,
  );
  if (!kinds.length) return [];
  const withSeries = await conn.client.getCollection<ODataEntity>(
    `${KINDS}${buildQuery({
      filter: or(...kinds.map((k) => cmp("Ref_Key", "eq", odataGuid(k)))),
      select: ["Ref_Key", "ИспользоватьСерии"],
    })}`,
  );
  const series = new Set(
    withSeries.value.filter((k) => k["ИспользоватьСерии"] === true).map((k) => String(k["Ref_Key"])),
  );
  const names = items.value
    .filter((i) => series.has(String(i["ВидНоменклатуры_Key"])))
    .map((i) => `«${i["Description"]}»`);
  return names.length
    ? [
        `Вид номенклатуры ведёт серии, а серия не указана: ${names.join(", ")}. Если склад учитывает серии, 1С ` +
          "не проведёт документ — укажите seriesRef в строке.",
      ]
    : [];
}

export interface TransferInput {
  organization?: string | undefined;
  from: string;
  to: string;
  date?: string | undefined;
  status?: "Принято" | "Отгружено" | undefined;
  comment?: string | undefined;
  lines: StockLine[];
}

/** Перемещение: по умолчанию «Принято» — один документ, расход с отправителя и приход на получателя. */
export async function createTradeTransfer(
  conn: Connection,
  input: TransferInput,
  confirm: boolean,
): Promise<CallToolResult> {
  if (input.from === input.to)
    throw new InputError("Склад-отправитель и склад-получатель совпадают. Ничего не записано.");
  const series = input.lines.filter((l) => l.seriesRef);
  const notes = await seriesNotes(conn, input.lines);
  if (input.lines.some((l) => l.price))
    notes.push("Цена в перемещении Управления торговлей не указывается: 1С берёт себестоимость сама.");
  const payload = {
    Date: await docDate(conn, input.date),
    Posted: false,
    Организация_Key: await resolveOrgRef(conn, input.organization),
    СкладОтправитель_Key: input.from,
    СкладПолучатель_Key: input.to,
    ХозяйственнаяОперация: "ПеремещениеТоваров",
    Статус: input.status ?? "Принято",
    ...(input.comment ? { Комментарий: input.comment } : {}),
    Товары: input.lines.map((l, i) => ({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.quantity,
    })),
    // Серии перемещения — в своей табличной части, как в документах, созданных в 1С.
    ...(series.length
      ? {
          Серии: series.map((l, i) => ({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Серия_Key: l.seriesRef,
            Количество: l.quantity,
          })),
        }
      : {}),
  };
  return createCore(conn, STOCK_DOCS.transfer, payload, confirm, notes, () => undefined);
}

export interface SurplusInput {
  organization?: string | undefined;
  warehouse: string;
  priceKind: string;
  incomeItem: string;
  inventoryRef?: string | undefined;
  date?: string | undefined;
  comment?: string | undefined;
  lines: StockLine[];
}

/** Оприходование излишков: вид цены и статья доходов — обязательные, коннектор их не угадывает (решение владельца). */
export async function createTradeSurplus(
  conn: Connection,
  input: SurplusInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const [kind, item, currency] = [
    await resolveCatalogItem(conn, ["Catalog_ВидыЦен"], "Вид цены", input.priceKind),
    await resolveCatalogItem(conn, [INCOME_ITEMS], "Статья доходов", input.incomeItem),
    await resolveCatalogItem(conn, ["Catalog_Валюты"], "Справочник «Валюты»", "643"),
  ];
  const payload = {
    Date: await docDate(conn, input.date),
    Posted: false,
    Организация_Key: await resolveOrgRef(conn, input.organization),
    Склад_Key: input.warehouse,
    ВидЦены_Key: kind.ref,
    СтатьяДоходов_Key: item.ref,
    Валюта_Key: currency.ref,
    ...(input.inventoryRef ? { ПересчетТоваров_Key: guidOf(input.inventoryRef) } : {}),
    ...(input.comment ? { Комментарий: input.comment } : {}),
    Товары: input.lines.map((l, i) => ({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.quantity,
      Цена: l.price ?? 0,
      Сумма: round2(l.quantity * (l.price ?? 0)),
      ...(l.seriesRef ? { Серия_Key: l.seriesRef } : {}),
    })),
  };
  return createCore(
    conn,
    STOCK_DOCS.surplus,
    payload,
    confirm,
    await seriesNotes(conn, input.lines),
    () => undefined,
  );
}

export interface WriteoffInput {
  organization?: string | undefined;
  warehouse: string;
  expenseItem: string;
  inventoryRef?: string | undefined;
  date?: string | undefined;
  comment?: string | undefined;
  lines: StockLine[];
}

/** Списание недостач: статья расходов обязательна; налогообложение НДС 1С заполняет сама по организации. */
export async function createTradeWriteoff(
  conn: Connection,
  input: WriteoffInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const item = await resolveCatalogItem(conn, [EXPENSE_ITEMS], "Статья расходов", input.expenseItem);
  const notes = await seriesNotes(conn, input.lines);
  if (input.lines.some((l) => l.price))
    notes.push("Цена в списании Управления торговлей не указывается: 1С списывает по себестоимости.");
  const payload = {
    Date: await docDate(conn, input.date),
    Posted: false,
    Организация_Key: await resolveOrgRef(conn, input.organization),
    Склад_Key: input.warehouse,
    // Составное поле: тип рядом со значением, как у документов, созданных в 1С.
    СтатьяРасходов: item.ref,
    СтатьяРасходов_Type: `StandardODATA.${EXPENSE_ITEMS}`,
    ...(input.inventoryRef ? { ПересчетТоваров_Key: guidOf(input.inventoryRef) } : {}),
    ...(input.comment ? { Комментарий: input.comment } : {}),
    Товары: input.lines.map((l, i) => ({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.quantity,
      ...(l.seriesRef ? { Серия_Key: l.seriesRef } : {}),
    })),
  };
  return createCore(conn, STOCK_DOCS.writeoff, payload, confirm, notes, () => undefined);
}

export interface InventoryInput {
  warehouse: string;
  date?: string | undefined;
  comment?: string | undefined;
  lines: Array<{ nomenclatureRef: string; factQuantity: number; accountingQuantity: number; price: number }>;
}

/**
 * Пересчёт товаров: «Количество» — по учёту, «КоличествоФакт» — по факту (так в пересчётах, созданных в 1С).
 * Учёт и факт переданы — пересчёт выполнен: статус «Выполнено», учётные данные заполнены. Движений по складу нет.
 */
export async function createTradeInventory(
  conn: Connection,
  input: InventoryInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const payload = {
    Date: await docDate(conn, input.date),
    Posted: false,
    Склад_Key: input.warehouse,
    Статус: "Выполнено",
    УчетныеДанныеЗаполнены: true,
    ...(input.comment ? { Комментарий: input.comment } : {}),
    Товары: input.lines.map((l, i) => ({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.accountingQuantity,
      КоличествоФакт: l.factQuantity,
      ...(l.price
        ? {
            Цена: l.price,
            Сумма: round2(l.accountingQuantity * l.price),
            СуммаФакт: round2(l.factQuantity * l.price),
          }
        : {}),
    })),
  };
  return createCore(conn, STOCK_DOCS.inventory, payload, confirm, [], () => undefined);
}

type Direction = "Receipt" | "Expense";
interface StockMove {
  warehouse: string;
  item: string;
  characteristic: string;
  direction: Direction;
  quantity: number;
}
const ref = (v: unknown) => {
  const s = String(v ?? "");
  return s && s !== EMPTY_REF ? s.toLowerCase() : "";
};
// Серия не в ключе: на копии рабочей базы регистр не хранит серию даже у документов с сериями.
const moveKey = (m: Omit<StockMove, "quantity">) =>
  `${m.direction} ${m.warehouse} ${m.item} ${m.characteristic}`;
const qty = (n: number) => Math.round(n * 1000);

/** Ожидаемые движения по строкам документа; undefined — статус, для которого ожидание неизвестно. */
export function expectedStockMoves(kind: StockKind, doc: ODataEntity): StockMove[] | undefined {
  const rows = (doc["Товары"] as ODataEntity[] | undefined) ?? [];
  const sides: Array<[Direction, string]> =
    kind === "surplus"
      ? [["Receipt", ref(doc["Склад_Key"])]]
      : kind === "writeoff"
        ? [["Expense", ref(doc["Склад_Key"])]]
        : kind === "transfer"
          ? doc["Статус"] === "Принято"
            ? [
                ["Expense", ref(doc["СкладОтправитель_Key"])],
                ["Receipt", ref(doc["СкладПолучатель_Key"])],
              ]
            : doc["Статус"] === "Отгружено"
              ? [["Expense", ref(doc["СкладОтправитель_Key"])]]
              : []
          : [];
  if (kind === "transfer" && !sides.length) return undefined;
  return sides.flatMap(([direction, warehouse]) =>
    rows.map((r) => ({
      direction,
      warehouse,
      item: ref(r["Номенклатура_Key"]),
      characteristic: ref(r["Характеристика_Key"]),
      quantity: Number(r["Количество"] ?? 0),
    })),
  );
}

/** Сверка: по каждому складу, товару и направлению количество равно, лишних движений нет. */
export function compareStockMoves(
  expected: readonly StockMove[],
  records: readonly ODataEntity[],
): { matched: boolean; expected: number; found: number; mismatches: string[] } {
  const sum = (items: Array<[string, number]>) => {
    const m = new Map<string, number>();
    for (const [k, n] of items) m.set(k, (m.get(k) ?? 0) + qty(n));
    return m;
  };
  const want = sum(expected.map((e) => [moveKey(e), e.quantity]));
  const got = sum(
    records.map((r) => [
      moveKey({
        direction: r["RecordType"] === "Expense" ? "Expense" : "Receipt",
        warehouse: ref(r["Склад_Key"]),
        item: ref(r["Номенклатура_Key"]),
        characteristic: ref(r["Характеристика_Key"]),
      }),
      Number(r["ВНаличии"] ?? 0),
    ]),
  );
  const mismatches: string[] = [];
  for (const [k, n] of want) {
    const g = got.get(k);
    if (g === undefined) mismatches.push(`нет движения: ${k}`);
    else if (g !== n) mismatches.push(`${k}: в регистре ${g / 1000}, в документе ${n / 1000}`);
    got.delete(k);
  }
  for (const [k, n] of got) mismatches.push(`лишнее движение: ${k} — ${n / 1000}`);
  return { matched: !mismatches.length, expected: want.size, found: records.length, mismatches };
}

async function stockRecords(conn: Connection, set: string, guid: string): Promise<ODataEntity[]> {
  try {
    const s = await conn.client.getEntity<ODataEntity>(
      `${STOCK_REGISTER}(Recorder='${guid}', Recorder_Type='StandardODATA.${set}')?$format=json`,
    );
    return ((s["RecordSet"] as ODataEntity[] | undefined) ?? []).filter((r) => r["Active"] !== false);
  } catch (e) {
    // Живая УТ 11.5: у документа без движений набора записей нет — 404, а не пустой набор.
    if (e instanceof ODataError && e.kind === "not_found") return [];
    throw e;
  }
}

/**
 * Проверка проведения складского документа УТ по регистру «Товары на складах». undefined — не складской документ УТ
 * (проверяют другие). Пересчёт движений не делает — только пояснение, проверено состояние документа.
 */
export async function verifyStockMovements(
  conn: Connection,
  entitySet: string,
  docRef: string,
  post: boolean,
): Promise<Record<string, unknown> | undefined> {
  const kind = await tradeStockSchema(conn, entitySet);
  if (!kind) return undefined;
  if (kind === "inventory")
    return { movementsNote: "Пересчёт товаров не делает движений по складу: проверено состояние документа." };
  const guid = guidOf(docRef);
  try {
    const records = await stockRecords(conn, entitySet, guid);
    let expected: StockMove[] | undefined = [];
    if (post) {
      const doc = await conn.client.getEntity<ODataEntity>(`${entitySet}(${odataGuid(guid)})?$format=json`);
      expected = expectedStockMoves(kind, doc);
      if (!expected)
        return {
          movementsNote: `Статус перемещения «${String(doc["Статус"])}»: какие движения ждать, коннектор не знает — сверка не выполнялась.`,
          movements: { register: STOCK_REGISTER, found: records.length },
        };
    }
    const check = compareStockMoves(expected, records);
    return {
      movementsVerified: check.matched,
      movements: {
        register: STOCK_REGISTER,
        expected: check.expected,
        found: check.found,
        ...(check.mismatches.length ? { mismatches: check.mismatches.slice(0, 20) } : {}),
      },
      ...(check.matched
        ? {}
        : {
            verified: false,
            note: post
              ? "Документ проведён, но движения по складу не совпали со строками. Проверьте документ в 1С; отменить " +
                "проведение — write.document.post_document с post=false."
              : "Проведение отменено, но в регистре «Товары на складах» остались движения документа. Проверьте его в 1С.",
          }),
    };
  } catch (e) {
    return {
      movementsVerified: false,
      verified: false,
      note:
        `Действие выполнено, но движения по складу прочитать не удалось (${e instanceof Error ? e.message : String(e)}). ` +
        "Проверьте документ в 1С, прежде чем повторять.",
    };
  }
}
