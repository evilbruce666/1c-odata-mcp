import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { AggregateOverflowError, fetchAllForAggregation } from "../odata/aggregate.js";
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
const WAREHOUSES = "Catalog_Склады";
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

/**
 * Складской документ устроен как в УТ, а база — с регистром бухгалтерии (КА 2, ERP 2). Сборщик Бухгалтерии записал
 * бы в него счета учёта, которых у документа нет, а сборщик УТ на такой базе не проверен: отказ с причиной (решение
 * владельца, 21.11.2026). undefined — отказывать не нужно.
 */
export async function stockInAccountingBaseRefusal(
  conn: Connection,
  set: string,
): Promise<string | undefined> {
  const kind = (Object.keys(STOCK_DOCS) as StockKind[]).find((k) => STOCK_DOCS[k] === set);
  if (!kind || (await withoutAccounting(conn)) || (await tradeStockSchema(conn, set)) !== kind)
    return undefined;
  return (
    `Документ ${set} устроен как в Управлении торговлей, а в базе есть регистр бухгалтерии — похоже, это КА 2 или ` +
    "ERP 2. Запись складских документов в таких базах не проверена: ничего не записано. Создайте документ в 1С."
  );
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
  transfer: ["Номенклатура_Key", "Количество", "КоличествоУпаковок"],
  surplus: ["Номенклатура_Key", "Количество", "Цена", "Сумма", "Серия_Key"],
  writeoff: ["Номенклатура_Key", "Количество", "Серия_Key"],
  inventory: [
    "Номенклатура_Key",
    "Количество",
    "КоличествоФакт",
    "КоличествоУпаковок",
    "КоличествоУпаковокФакт",
  ],
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

const REF_CHUNK = 20;
/** С этого числа ссылок справочник читается целиком, страницами по WHOLE_SET_PAGE (строки короткие: 3 поля). */
const WHOLE_SET_FROM = 200;
const WHOLE_SET_PAGE = 5000;
/**
 * Справочник, который не поместился в предел выборки (ODATA_ANALYTICS_MAX_ROWS): 10 минут читаем сразу по ссылкам,
 * а не пробуем целиком на каждом предпросмотре (каждая неудачная попытка — страницы до предела, как у readPageAttributes).
 */
const wholeSetTooBig = new WeakMap<object, Map<string, number>>();
const TOO_BIG_TTL_MS = 10 * 60_000;
/** Сколько позиций назвать в заметке о сериях. */
const NOTE_NAMES = 10;
const guidOf = (s: string) => s.replace(/[{}']/g, "").trim();
const round2 = (n: number) => Math.round(n * 100) / 100;

/**
 * Позиции, вид которых ведёт серии, а серия не указана. Не отказ: нужна ли серия, решает политика учёта серий на
 * складе (на копии рабочей базы у одного вида — разные политики по складам). Проверено на копии: без серии 1С
 * проводит документ без отказа — учёт серий остаётся неполным, поэтому предупреждение. Схема без видов — проверки нет.
 */
async function seriesNotes(
  conn: Connection,
  lines: readonly StockLine[],
  /** Что посоветовать: в перемещении — указать серию, в оприходовании и списании — оформить в 1С. */
  advice = "укажите seriesRef в строке, если склад их учитывает.",
): Promise<string[]> {
  const without = [...new Set(lines.filter((l) => !l.seriesRef).map((l) => guidOf(l.nomenclatureRef)))];
  const available = await conn.available();
  if (!without.length || !available.has(KINDS)) return [];
  const meta = await conn.getMetadata();
  const has = (set: string, f: string) => !!meta.entities.get(set)?.properties.some((p) => p.name === f);
  if (!has(ITEMS, "ВидНоменклатуры_Key") || !has(KINDS, "ИспользоватьСерии")) return [];
  // По ссылкам — частями: длинный отбор «or» 1С не примет (как у цен). Много ссылок — справочник целиком: документ на
  // 1000 позиций читал номенклатуру 50 запросами (около минуты в облачной 1С), справочник на 15 тыс. — 3 запроса.
  const byRefs = async (set: string, refs: readonly string[], select: string[]) => {
    const tooBig = wholeSetTooBig.get(conn)?.get(set);
    if (refs.length >= WHOLE_SET_FROM && !(tooBig !== undefined && Date.now() - tooBig < TOO_BIG_TTL_MS))
      try {
        const wanted = new Set(refs.map((r) => r.toLowerCase()));
        const { rows } = await fetchAllForAggregation(conn, set, { select }, undefined, WHOLE_SET_PAGE);
        return rows.filter((r) => wanted.has(String(r["Ref_Key"]).toLowerCase()));
      } catch (e) {
        // Справочник больше предела выборки — по ссылкам, как раньше; следующие предпросмотры сразу по ссылкам.
        if (!(e instanceof AggregateOverflowError)) throw e;
        const marks = wholeSetTooBig.get(conn) ?? new Map<string, number>();
        marks.set(set, Date.now());
        wholeSetTooBig.set(conn, marks);
      }
    const out: ODataEntity[] = [];
    for (let i = 0; i < refs.length; i += REF_CHUNK)
      out.push(
        ...(
          await conn.client.getCollection<ODataEntity>(
            `${set}${buildQuery({
              filter: or(...refs.slice(i, i + REF_CHUNK).map((r) => cmp("Ref_Key", "eq", odataGuid(r)))),
              select,
            })}`,
          )
        ).value,
      );
    return out;
  };
  const items = { value: await byRefs(ITEMS, without, ["Ref_Key", "Description", "ВидНоменклатуры_Key"]) };
  const kinds = [...new Set(items.value.map((i) => String(i["ВидНоменклатуры_Key"] ?? "")))].filter(
    (k) => k && k !== EMPTY_REF,
  );
  if (!kinds.length) return [];
  const withSeries = { value: await byRefs(KINDS, kinds, ["Ref_Key", "ИспользоватьСерии"]) };
  const series = new Set(
    withSeries.value.filter((k) => k["ИспользоватьСерии"] === true).map((k) => String(k["Ref_Key"])),
  );
  const names = items.value
    .filter((i) => series.has(String(i["ВидНоменклатуры_Key"])))
    .map((i) => `«${i["Description"]}»`);
  // Первые NOTE_NAMES позиций, остальные — числом: на 1000 строк перечень занимал 58 тыс. знаков (копия УТ, 23.11.2026).
  const listed =
    names.length > NOTE_NAMES
      ? `${names.slice(0, NOTE_NAMES).join(", ")} и ещё ${names.length - NOTE_NAMES}`
      : names.join(", ");
  return names.length
    ? [
        `Вид номенклатуры ведёт серии, а серия не указана: ${listed}. 1С проведёт документ и без ` +
          `серии, но учёт серий на складе останется неполным — ${advice}`,
      ]
    : [];
}

const VAT_ACTIVITY = "ПеремещениеПодДеятельность";
const SERIES_IN_1C =
  "серии в этом документе коннектор пока не пишет; если склад их учитывает, оформите документ в 1С.";
const SERIES_IN_TABLE = {
  СтатусУказанияСерий: 2,
  СтатусУказанияСерийОтправитель: 2,
  СтатусУказанияСерийПолучатель: 2,
} as const;

/**
 * Серии в оприходовании и списании УТ не поддержаны: на копии рабочей базы нет ни одной такой строки с серией, а без
 * верного статуса указания серий 1С очищает серию при записи — она потерялась бы молча.
 */
function refuseSeries(lines: readonly StockLine[], what: string): void {
  if (lines.some((l) => l.seriesRef))
    throw new InputError(
      `Серии в ${what} Управления торговлей пока не поддержаны: 1С сбросила бы серию без ошибки. Серию можно ` +
        "указать в перемещении; здесь — уберите seriesRef или оформите документ в 1С. Ничего не записано.",
    );
}

/**
 * Налогообложение НДС перемещения. При записи через OData 1С его не заполняет (у списания — заполняет сама), а у
 * организации оно одно во всех перемещениях копии рабочей базы. Берётся из последнего проведённого перемещения той
 * же организации, с заметкой; нет такого — поле пустое, тоже с заметкой (решение владельца, 21.11.2026).
 */
async function transferVatActivity(conn: Connection, org: string): Promise<{ value?: string; note: string }> {
  const meta = await conn.getMetadata();
  if (!meta.entities.get(STOCK_DOCS.transfer)?.properties.some((p) => p.name === VAT_ACTIVITY))
    return { note: "Налогообложение НДС перемещения в этой схеме не опубликовано — не заполняется." };
  const last = (
    await conn.client.getCollection<ODataEntity>(
      `${STOCK_DOCS.transfer}${buildQuery({
        filter: `${cmp("Организация_Key", "eq", odataGuid(org))} and Posted eq true and DeletionMark eq false`,
        orderby: "Date desc",
        top: 1,
        select: ["Number", "Date", VAT_ACTIVITY],
      })}`,
    )
  ).value[0];
  const value = String(last?.[VAT_ACTIVITY] ?? "");
  if (!last || !value)
    return {
      note:
        "Налогообложение НДС перемещения («ПеремещениеПодДеятельность») не заполнено: у организации нет проведённых " +
        "перемещений, откуда его взять. Проверьте документ в 1С.",
    };
  return {
    value,
    note: `Налогообложение НДС «${value}» — как в последнем проведённом перемещении организации № ${String(last["Number"])} от ${String(last["Date"]).slice(0, 10)}.`,
  };
}

/**
 * Строки складского документа УТ из позиций: одна сборка для создания и правки строк. Перемещение — с количеством
 * упаковок и серией в таблице «Серии» (статус 2, без него 1С очищает серии при записи); оприходование — с ценой и
 * суммой; списание — только количество. Пересчёт строится отдельно: у него учёт и факт.
 */
export function stockRows(
  kind: Exclude<StockKind, "inventory">,
  lines: readonly StockLine[],
): { Товары: Array<Record<string, unknown>>; Серии?: Array<Record<string, unknown>> } {
  const Товары = lines.map((l, i) => ({
    LineNumber: i + 1,
    Номенклатура_Key: l.nomenclatureRef,
    Количество: l.quantity,
    ...(kind === "transfer"
      ? {
          // Форма 1С показывает и правит количество упаковок; без упаковки оно равно количеству (так у всех
          // перемещений копии рабочей базы).
          КоличествоУпаковок: l.quantity,
          ...(l.seriesRef ? SERIES_IN_TABLE : {}),
        }
      : {}),
    ...(kind === "surplus" ? { Цена: l.price!, Сумма: round2(l.quantity * l.price!) } : {}),
  }));
  const series = kind === "transfer" ? lines.filter((l) => l.seriesRef) : [];
  return series.length
    ? {
        Товары,
        Серии: series.map((l, i) => ({
          LineNumber: i + 1,
          Номенклатура_Key: l.nomenclatureRef,
          Серия_Key: l.seriesRef,
          Количество: l.quantity,
        })),
      }
    : { Товары };
}

/** Молчаливую потерю серий видно сразу: 1С возвращает записанный документ. */
export function seriesSavedNote(sent: number, saved: ODataEntity): string | undefined {
  const n = (saved["Серии"] as unknown[] | undefined)?.length ?? 0;
  return sent && n !== sent
    ? `1С сохранила серий ${n} из ${sent}. Проверьте документ в 1С, прежде чем проводить.`
    : undefined;
}

const DEPARTMENTS = "Catalog_СтруктураПредприятия";

/**
 * Подразделение оприходования и списания — аналитика доходов и расходов, заполнена в 95–98% документов копии рабочей
 * базы, но с подразделением склада совпадает у списаний лишь в трети случаев: не угадываем. Не указано — пусто, с
 * заметкой (решение владельца, 21.11.2026).
 */
async function departmentOf(
  conn: Connection,
  department: string | undefined,
): Promise<{ field: Record<string, string>; note?: string }> {
  if (!department)
    return {
      field: {},
      note:
        "Подразделение не указано — поле пустое; для аналитики доходов и расходов укажите department " +
        "(название, код или Ref_Key).",
    };
  const d = await resolveCatalogItem(conn, [DEPARTMENTS], "Подразделение", department);
  return { field: { Подразделение_Key: d.ref } };
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
  const org = await resolveOrgRef(conn, input.organization);
  const activity = await transferVatActivity(conn, org);
  notes.push(activity.note);
  const payload = {
    Date: await docDate(conn, input.date),
    Posted: false,
    Организация_Key: org,
    ...(activity.value ? { ПеремещениеПодДеятельность: activity.value } : {}),
    СкладОтправитель_Key: input.from,
    СкладПолучатель_Key: input.to,
    ХозяйственнаяОперация: "ПеремещениеТоваров",
    Статус: input.status ?? "Принято",
    ...(input.comment ? { Комментарий: input.comment } : {}),
    ...stockRows("transfer", input.lines),
  };
  return createCore(conn, STOCK_DOCS.transfer, payload, confirm, notes, (created) =>
    seriesSavedNote(series.length, created),
  );
}

export interface SurplusInput {
  organization?: string | undefined;
  department?: string | undefined;
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
  refuseSeries(input.lines, "оприходовании");
  const [kind, item, currency] = [
    await resolveCatalogItem(conn, ["Catalog_ВидыЦен"], "Вид цены", input.priceKind),
    await resolveCatalogItem(conn, [INCOME_ITEMS], "Статья доходов", input.incomeItem),
    await resolveCatalogItem(conn, ["Catalog_Валюты"], "Справочник «Валюты»", "643"),
  ];
  requirePrices(input.lines, kind.name ?? input.priceKind);
  const date = await docDate(conn, input.date);
  const department = await departmentOf(conn, input.department);
  const payload = {
    Date: date,
    Posted: false,
    Организация_Key: await resolveOrgRef(conn, input.organization),
    ...department.field,
    Склад_Key: input.warehouse,
    ВидЦены_Key: kind.ref,
    СтатьяДоходов_Key: item.ref,
    Валюта_Key: currency.ref,
    ...(input.inventoryRef ? { ПересчетТоваров_Key: guidOf(input.inventoryRef) } : {}),
    ...(input.comment ? { Комментарий: input.comment } : {}),
    ...stockRows("surplus", input.lines),
  };
  const notes = [
    ...(department.note ? [department.note] : []),
    ...(await seriesNotes(conn, input.lines, SERIES_IN_1C)),
  ];
  return createCore(conn, STOCK_DOCS.surplus, payload, confirm, notes, () => undefined);
}

/**
 * Цена каждой строки оприходования обязательна (решение владельца, 21.11.2026). Подстановка по виду цены дважды дала
 * находки ревью: цена в валюте и за упаковку уходила бы без пересчёта. Узнать цену — read.analytics.get_prices.
 */
function requirePrices(lines: readonly StockLine[], kindName: string): void {
  const missing = lines.map((l, i) => (l.price ? 0 : i + 1)).filter(Boolean);
  if (missing.length)
    throw new InputError(
      `Оприходование излишков: укажите price в строках ${missing.join(", ")} — коннектор цену не подставляет. ` +
        `Действующую цену вида «${kindName}» покажет read.analytics.get_prices; уточните её у пользователя. ` +
        "Ничего не записано.",
    );
}

export interface WriteoffInput {
  organization?: string | undefined;
  department?: string | undefined;
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
  refuseSeries(input.lines, "списании");
  const notes = await seriesNotes(conn, input.lines, SERIES_IN_1C);
  if (input.lines.some((l) => l.price))
    notes.push("Цена в списании Управления торговлей не указывается: 1С списывает по себестоимости.");
  const department = await departmentOf(conn, input.department);
  if (department.note) notes.unshift(department.note);
  const payload = {
    Date: await docDate(conn, input.date),
    Posted: false,
    Организация_Key: await resolveOrgRef(conn, input.organization),
    ...department.field,
    Склад_Key: input.warehouse,
    // Составное поле: тип рядом со значением, как у документов, созданных в 1С.
    СтатьяРасходов: item.ref,
    СтатьяРасходов_Type: `StandardODATA.${EXPENSE_ITEMS}`,
    ...(input.inventoryRef ? { ПересчетТоваров_Key: guidOf(input.inventoryRef) } : {}),
    ...(input.comment ? { Комментарий: input.comment } : {}),
    ...stockRows("writeoff", input.lines),
  };
  return createCore(conn, STOCK_DOCS.writeoff, payload, confirm, notes, () => undefined);
}

export interface InventoryInput {
  organization?: string | undefined;
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
      // Без упаковки — те же числа, как у всех пересчётов копии рабочей базы.
      КоличествоУпаковок: l.accountingQuantity,
      КоличествоУпаковокФакт: l.factQuantity,
      ...(l.price
        ? {
            Цена: l.price,
            Сумма: round2(l.accountingQuantity * l.price),
            СуммаФакт: round2(l.factQuantity * l.price),
          }
        : {}),
    })),
  };
  const notes = input.organization
    ? ["У пересчёта товаров в Управлении торговлей нет организации: organization не используется."]
    : [];
  return createCore(conn, STOCK_DOCS.inventory, payload, confirm, notes, () => undefined);
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

/** Какие флаги ордерной схемы склада касаются документа: сторона склада → поле флага и даты начала. */
const ORDER_SCHEME: Record<Exclude<StockKind, "inventory">, Array<[string, string]>> = {
  surplus: [["Склад_Key", "ПриОтраженииИзлишковНедостач"]],
  writeoff: [["Склад_Key", "ПриОтраженииИзлишковНедостач"]],
  transfer: [
    ["СкладОтправитель_Key", "ПриОтгрузке"],
    ["СкладПолучатель_Key", "ПриПоступлении"],
  ],
};

/**
 * Склады документа с ордерной схемой на его дату. При ордерной схеме «в наличии» меняет складской ордер, а не
 * распоряжение, и сверка по строкам ложно показала бы расхождение с советом отменить проведение (ревью выпуска 3).
 * На копии рабочей базы ордерных складов нет — поэтому не сверяем, а не угадываем движения. Схема без флагов — нет.
 */
async function orderSchemeWarehouses(
  conn: Connection,
  kind: Exclude<StockKind, "inventory">,
  doc: ODataEntity,
): Promise<string[]> {
  const out: string[] = [];
  for (const [side, f] of ORDER_SCHEME[kind])
    out.push(...(await orderedWarehouses(conn, [String(doc[side] ?? "")], f, String(doc["Date"] ?? ""))));
  return out;
}

/**
 * Какие из складов на дату работают по ордерной схеме для стороны f (ПриОтгрузке, ПриПоступлении,
 * ПриОтраженииИзлишковНедостач): названия в кавычках. Схема справочника без флага — ордерных нет.
 */
export async function orderedWarehouses(
  conn: Connection,
  refs: readonly string[],
  f: string,
  date: string,
): Promise<string[]> {
  const props = new Set(
    (await conn.getMetadata()).entities.get(WAREHOUSES)?.properties.map((p) => p.name) ?? [],
  );
  const flag = `ИспользоватьОрдернуюСхему${f}`;
  if (!props.has(flag)) return [];
  const since = `ДатаНачалаОрдернойСхемы${f}`;
  const day = date.slice(0, 10);
  const out: string[] = [];
  for (const ref of new Set(refs.map((r) => r.toLowerCase()))) {
    if (!ref || ref === EMPTY_REF) continue;
    const w = await conn.client.getEntity<ODataEntity>(
      `${WAREHOUSES}(${odataGuid(ref)})${buildQuery({ select: ["Description", flag, ...(props.has(since) ? [since] : [])] })}`,
    );
    const start = String(w[since] ?? "").slice(0, 10);
    if (w[flag] === true && (!start || start === "0001-01-01" || start <= day))
      out.push(`«${String(w["Description"] ?? ref)}»`);
  }
  return out;
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
  const kind = (Object.keys(STOCK_DOCS) as StockKind[]).find((k) => STOCK_DOCS[k] === entitySet);
  // Как у сборщика: только база без регистра бухгалтерии; в КА и ERP — проводки, как раньше.
  if (!kind || !(await tradeStockSet(conn, kind))) return undefined;
  if (kind === "inventory")
    return { movementsNote: "Пересчёт товаров не делает движений по складу: проверено состояние документа." };
  const guid = guidOf(docRef);
  try {
    const records = await stockRecords(conn, entitySet, guid);
    let expected: StockMove[] | undefined = [];
    if (post) {
      const doc = await conn.client.getEntity<ODataEntity>(`${entitySet}(${odataGuid(guid)})?$format=json`);
      const ordered = await orderSchemeWarehouses(conn, kind, doc);
      if (ordered.length)
        return {
          movementsNote:
            `Склад с ордерной схемой (${ordered.join(", ")}): товар в наличии меняют складские ордера, а не этот ` +
            "документ — сверка движений не выполнялась. Проверено состояние документа.",
          movements: { register: STOCK_REGISTER, found: records.length },
        };
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
