import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { odataGuid } from "../odata/query.js";
import { fail } from "./_shared.js";
import { createCore, docDate, patchLinesOrPreview, patchOrPreview, stripForCopy } from "./write.js";
import {
  STOCK_DOCS,
  seriesSavedNote,
  stockRows,
  tradeStockSet,
  type StockKind,
  type StockLine,
} from "./trade-warehouse.js";

/**
 * Правка, копирование и изменение полей документов Управления торговлей. Строки собираются той же функцией, что при
 * создании (stockRows): поля, которые 1С заполняет при проведении (виды запасов, ключи аналитики), не передаются —
 * 1С пересчитает их сама. Только непроведённые документы: правка проведённого меняла бы движения.
 */

/** Документ УТ, который умеют править эти инструменты; undefined — сборщик Бухгалтерии, как раньше. */
export async function tradeEditKind(conn: Connection, entitySet: string): Promise<StockKind | undefined> {
  const kind = (Object.keys(STOCK_DOCS) as StockKind[]).find((k) => STOCK_DOCS[k] === entitySet);
  return kind && (await tradeStockSet(conn, kind)) ? kind : undefined;
}

const guidOf = (s: string) => s.replace(/[{}']/g, "").trim();
type Row = Record<string, unknown>;
const rowsOf = (doc: ODataEntity, section: string) => (doc[section] as Row[] | undefined) ?? [];
const renumber = (rows: Row[]) => rows.map((r, i) => ({ ...r, LineNumber: i + 1 }));

async function readDoc(conn: Connection, entitySet: string, ref: string): Promise<ODataEntity> {
  return conn.client.getEntity<ODataEntity>(`${entitySet}(${odataGuid(guidOf(ref))})?$format=json`);
}

const POSTED =
  "Документ проведён. Сначала отмените проведение (write.document.post_document с post=false), затем меняйте его.";
const INVENTORY_LINES =
  "Строки пересчёта товаров (учёт и факт) этими инструментами не правятся: создайте новый пересчёт " +
  "(write.warehouse.create_inventory), а этот пометьте на удаление.";

/** Дописывает заметку в успешный ответ. */
function withNote(result: CallToolResult, note: string | undefined): CallToolResult {
  if (!note || result.isError || !result.structuredContent) return result;
  const sc = result.structuredContent as Record<string, unknown>;
  const notes = [...((sc["notes"] as string[] | undefined) ?? []), note];
  return { ...result, structuredContent: { ...sc, notes } };
}

/** После правки перемещения с сериями — сохранила ли 1С серии (без статуса указания она их очищает). */
async function seriesCheck(
  conn: Connection,
  entitySet: string,
  ref: string,
  sent: number,
  result: CallToolResult,
): Promise<CallToolResult> {
  if (!sent || result.isError || (result.structuredContent as Record<string, unknown>)?.["dryRun"])
    return result;
  return withNote(result, seriesSavedNote(sent, await readDoc(conn, entitySet, ref)));
}

const ROW_CHECK = ["Номенклатура_Key", "Количество", "КоличествоУпаковок", "Цена", "Сумма"] as const;
const same = (a: unknown, b: unknown) =>
  typeof a === "number" || typeof b === "number"
    ? Math.abs(Number(a) - Number(b)) < 1e-6
    : String(a) === String(b);

/**
 * Сверка строк после правки: перечитанные строки совпадают с отправленными по позиции, количеству и цене. Общая
 * проверка правки табличные части не сверяет (verified=false) — для документа УТ сверяем сами.
 */
export async function rowsCheck(
  conn: Connection,
  entitySet: string,
  ref: string,
  sent: readonly Row[],
  result: CallToolResult,
): Promise<CallToolResult> {
  const sc = result.structuredContent as Record<string, unknown> | undefined;
  if (result.isError || !sc || sc["dryRun"] || sc["replayed"]) return result;
  const stored = rowsOf(await readDoc(conn, entitySet, ref), "Товары");
  const mismatches: string[] = [];
  if (stored.length !== sent.length)
    mismatches.push(`строк в 1С ${stored.length}, отправлено ${sent.length}`);
  sent.forEach((row, i) => {
    for (const k of ROW_CHECK)
      if (k in row && !same(row[k], stored[i]?.[k])) mismatches.push(`строка ${i + 1}: ${k}`);
  });
  if (mismatches.length)
    return {
      ...result,
      structuredContent: {
        ...sc,
        verified: false,
        note: `1С сохранила строки не так, как отправлено: ${mismatches.slice(0, 10).join("; ")}. Проверьте документ в 1С.`,
      },
    };
  return { ...result, structuredContent: { ...sc, verified: true } };
}

function priceRefusal(kind: StockKind, lines: readonly StockLine[]): string | undefined {
  if (kind === "surplus") {
    const missing = lines.map((l, i) => (l.price ? 0 : i + 1)).filter(Boolean);
    if (missing.length)
      return `Оприходование излишков: укажите price в строках ${missing.join(", ")} — коннектор цену не подставляет.`;
  }
  if (kind !== "transfer" && lines.some((l) => l.seriesRef))
    return "Серии в оприходовании и списании Управления торговлей пока не поддержаны: 1С сбросила бы серию без ошибки.";
  return undefined;
}

/** Замена всех строк документа. */
export async function replaceTradeLines(
  conn: Connection,
  entitySet: string,
  kind: StockKind,
  ref: string,
  lines: StockLine[],
  confirm: boolean,
): Promise<CallToolResult> {
  if (kind === "inventory") return fail(INVENTORY_LINES);
  const refusal = priceRefusal(kind, lines);
  if (refusal) return fail(`${refusal} Ничего не записано.`);
  const doc = await readDoc(conn, entitySet, ref);
  if (doc["Posted"] === true) return fail(POSTED);
  const built = stockRows(kind, lines);
  // Серии перемещения заменяются вместе со строками: прежние без новых строк означали бы чужие серии.
  const fields: Row = { Товары: built.Товары, ...(kind === "transfer" ? { Серии: built.Серии ?? [] } : {}) };
  // Как у Бухгалтерии: замена строк — без журнала операций (у инструмента нет operationId), сохранённое
  // перечитывается.
  const result = await patchOrPreview(conn, entitySet, ref, fields, confirm);
  return seriesCheck(
    conn,
    entitySet,
    ref,
    built.Серии?.length ?? 0,
    await rowsCheck(conn, entitySet, ref, built.Товары, result),
  );
}

/** Добавление строки: прежние строки остаются со всеми полями, которые заполнила 1С. */
export async function addTradeLine(
  conn: Connection,
  entitySet: string,
  kind: StockKind,
  ref: string,
  line: StockLine,
  confirm: boolean,
): Promise<CallToolResult> {
  if (kind === "inventory") return fail(INVENTORY_LINES);
  const refusal = priceRefusal(kind, [line]);
  if (refusal) return fail(`${refusal} Ничего не записано.`);
  const doc = await readDoc(conn, entitySet, ref);
  if (doc["Posted"] === true) return fail(POSTED);
  const before = rowsOf(doc, "Товары");
  const built = stockRows(kind, [line]);
  const fields: Row = { Товары: renumber([...before, ...built.Товары]) };
  const series = rowsOf(doc, "Серии");
  if (built.Серии) fields["Серии"] = renumber([...series, ...built.Серии]);
  const result = await rowsCheck(
    conn,
    entitySet,
    ref,
    fields["Товары"] as Row[],
    await patchLinesOrPreview(conn, entitySet, ref, fields, confirm, {
      section: "Товары",
      before: before.length,
      expected: before.length + 1,
    }),
  );
  return seriesCheck(conn, entitySet, ref, built.Серии ? series.length + built.Серии.length : 0, result);
}

/** Удаление строки по номеру. Серии перемещения уходят вместе со строкой своей позиции. */
export async function removeTradeLine(
  conn: Connection,
  entitySet: string,
  kind: StockKind,
  ref: string,
  lineNumber: number,
  confirm: boolean,
): Promise<CallToolResult> {
  if (kind === "inventory") return fail(INVENTORY_LINES);
  const doc = await readDoc(conn, entitySet, ref);
  if (doc["Posted"] === true) return fail(POSTED);
  const rows = rowsOf(doc, "Товары");
  if (lineNumber > rows.length) return fail(`В документе всего ${rows.length} строк(и).`);
  const removed = rows[lineNumber - 1]!;
  const kept = rows.filter((_, i) => i + 1 !== lineNumber);
  if (!kept.length)
    return fail("Нельзя удалить последнюю строку — в документе должна остаться хотя бы одна позиция.");
  const fields: Row = { Товары: renumber(kept) };
  const item = String(removed["Номенклатура_Key"] ?? "");
  const series = rowsOf(doc, "Серии");
  const ofItem = series.filter((s) => String(s["Номенклатура_Key"]) === item);
  if (ofItem.length) {
    // Серии не привязаны к строке: у той же позиции в другой строке они не разносятся однозначно.
    if (kept.some((r) => String(r["Номенклатура_Key"]) === item))
      return fail(
        "У этой позиции есть серии, а сама позиция — ещё в другой строке: какие серии убрать, не определить. " +
          "Замените строки целиком (write.document.update_document_lines). Ничего не записано.",
      );
    fields["Серии"] = renumber(series.filter((s) => String(s["Номенклатура_Key"]) !== item));
  }
  return rowsCheck(
    conn,
    entitySet,
    ref,
    fields["Товары"] as Row[],
    await patchLinesOrPreview(conn, entitySet, ref, fields, confirm, {
      section: "Товары",
      before: rows.length,
      expected: kept.length,
    }),
  );
}

/** Поля шапки, которые можно менять (update_entity) и задавать копии (copy_document): белый список по документу. */
const EDITABLE: Record<StockKind, readonly string[]> = {
  transfer: ["Статус", "СкладОтправитель_Key", "СкладПолучатель_Key", "ПеремещениеПодДеятельность"],
  surplus: ["Склад_Key", "ВидЦены_Key", "СтатьяДоходов_Key", "ПересчетТоваров_Key"],
  writeoff: ["Склад_Key", "СтатьяРасходов", "СтатьяРасходов_Type", "ПересчетТоваров_Key"],
  inventory: ["Склад_Key", "Статус"],
};
const COMMON = ["Комментарий", "Date", "Подразделение_Key", "Ответственный_Key"] as const;
const editable = (kind: StockKind) => new Set<string>([...COMMON, ...EDITABLE[kind]]);
const notEditable = (kind: StockKind, fields: Record<string, unknown>) => {
  const allowed = editable(kind);
  return Object.keys(fields).filter((k) => !allowed.has(k));
};
const whitelistRefusal = (kind: StockKind, bad: string[]) =>
  `Поля ${bad.join(", ")} в этом документе Управления торговлей коннектор не меняет: часть их 1С заполняет сама ` +
  `при проведении. Можно: ${[...editable(kind)].join(", ")}; строки — инструментами строк документа. Ничего не записано.`;

/** Изменение полей шапки по белому списку, только у непроведённого документа. */
export async function updateTradeFields(
  conn: Connection,
  entitySet: string,
  kind: StockKind,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
): Promise<CallToolResult> {
  const bad = notEditable(kind, fields);
  if (bad.length) return fail(whitelistRefusal(kind, bad));
  const doc = await readDoc(conn, entitySet, ref);
  if (doc["Posted"] === true) return fail(POSTED);
  return patchOrPreview(conn, entitySet, ref, fields, confirm);
}

/**
 * Поля, которые копия не берёт у образца: основание (новый документ по тому же пересчёту или заказу задвоил бы его),
 * виды запасов и ключи аналитики (1С заполнит при проведении по новой дате), служебные поля автора.
 */
const NOT_COPIED = new Set([
  "ПересчетТоваров_Key",
  "ДокументОснование",
  "ЗаказНаПеремещение_Key",
  "Основание",
  "ВидыЗапасов",
  "Автор_Key",
  "ДополнительныеРеквизиты",
  "ШтрихкодыУпаковок",
  "ОтборПересчета_Base64Data",
]);
const NOT_COPIED_ROW = new Set(["АналитикаУчетаНоменклатуры_Key", "ЗаказНаПеремещение_Key", "КодСтроки"]);
/** Поля строки, которые можно поправить в копии; количество упаковок и сумма следуют за количеством и ценой. */
const ROW_EDITABLE = new Set(["Количество", "Цена"]);

export interface TradeCopyInput {
  date?: string | undefined;
  fields?: Record<string, unknown> | undefined;
  lines?: Array<{ lineNumber: number; fields: Record<string, unknown> }> | undefined;
}

/** Копия складского документа УТ: непроведённая, без основания, с новым номером и датой. */
export async function copyTradeDocument(
  conn: Connection,
  entitySet: string,
  kind: StockKind,
  ref: string,
  input: TradeCopyInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const bad = notEditable(kind, input.fields ?? {});
  if (bad.length) return fail(whitelistRefusal(kind, bad));
  const badRow = (input.lines ?? []).flatMap((l) =>
    Object.keys(l.fields).filter((k) => !ROW_EDITABLE.has(k)),
  );
  if (badRow.length)
    return fail(
      `Поля строк ${[...new Set(badRow)].join(", ")} в копии не меняются: можно ${[...ROW_EDITABLE].join(", ")}. ` +
        "Ничего не записано.",
    );
  if (kind === "inventory" && input.lines?.length) return fail(`${INVENTORY_LINES} Ничего не записано.`);
  const source = stripForCopy(await readDoc(conn, entitySet, ref));
  const payload: Row = {};
  for (const [k, v] of Object.entries(source)) if (!NOT_COPIED.has(k)) payload[k] = v;
  const rows = rowsOf(source, "Товары").map((r) =>
    Object.fromEntries(Object.entries(r).filter(([k]) => !NOT_COPIED_ROW.has(k))),
  );
  for (const l of input.lines ?? []) {
    const row = rows[l.lineNumber - 1];
    if (!row) return fail(`В документе-образце нет строки ${l.lineNumber}. Ничего не записано.`);
    Object.assign(row, l.fields);
    if ("Количество" in l.fields && "КоличествоУпаковок" in row)
      row["КоличествоУпаковок"] = row["Количество"];
    if ("Сумма" in row)
      row["Сумма"] = Math.round(Number(row["Количество"] ?? 0) * Number(row["Цена"] ?? 0) * 100) / 100;
  }
  payload["Товары"] = rows;
  payload["Date"] = await docDate(conn, input.date);
  payload["Posted"] = false;
  Object.assign(payload, input.fields ?? {});
  const series = rowsOf(payload, "Серии").length;
  const notes = [
    "Копия — без основания (пересчёта, заказа) и видов запасов образца: 1С заполнит их при проведении.",
    ...(input.lines?.some((l) => "Количество" in l.fields) && series
      ? ["Серии скопированы как в образце: если менялось количество, проверьте серии в 1С."]
      : []),
  ];
  return createCore(conn, entitySet, payload, confirm, notes, (created) => seriesSavedNote(series, created));
}
