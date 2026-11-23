import type { McpServer } from "@modelcontextprotocol/server";
import { checkResponsible } from "../odata/users.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { InputError } from "../errors.js";
import { pageCached } from "../odata/cache.js";
import { callPageStart } from "../odata/call-context.js";
import {
  MAX_ITEM_GROUPS,
  nomenclatureSet,
  readPageAttributes,
  selectByItem,
  selectionNotes,
} from "../odata/inventory.js";
import {
  EMPTY_REF,
  MECHANISMS,
  PRICE_DOC,
  dependentKinds,
  detectMechanism,
  laterPrices,
  priceKey,
  readPriceKinds,
  resolvePriceKind,
  slicePrices,
  type PriceKind,
} from "../odata/pricing.js";
import { localToday } from "../odata/settlements-aging.js";
import { AggregateOverflowError } from "../odata/aggregate.js";
import { currentWriteOperationId } from "../odata/write-operation-context.js";
import { createResultSchema, getPricesResultSchema } from "../schemas/output.js";
import { MAX_RESPONSE_CHARS, databaseField, fitToChars, guard, ok, renderJson } from "./_shared.js";
import { confirmField, createCore } from "./write.js";
import type { PreviewView } from "./preview-view.js";
import { OPERATION_ID_CHARS, RESPONSE_RESERVE } from "./preview-fit.js";

const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;
const MOMENT_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?$/;
/**
 * Строк в документе. Ответ больше не растёт со строками: что не помещается, сжимается (fitPrices). Предел — время
 * чтения и записи: на копии рабочей УТ предпросмотр 1000 строк — около 15 с.
 */
const MAX_LINES = 1000;

/**
 * До стольких позиций цены читаются по ссылкам пачками по 20, как в 1.1. Больше — по виду цен целиком, одним запросом:
 * 1000 строк пачками — 100 запросов цен, в облачной 1С это минуты.
 */
const BY_REFS_UP_TO = 80;

/** Отклонение цены: изменение больше стольких процентов текущей. Ровно 20 % — не отклонение. */
const DEVIATION_PCT = 20;

/** Название позиции в сжатой строке. */
const COMPACT_NAME = 60;

const mechanismField = z
  .enum(["legacy", "2.5", "bp"])
  .optional()
  .describe(
    "Механизм цен: legacy — прежний УТ (табличная часть «Товары»), 2.5 — ценообразование 2.5, bp — Бухгалтерия (типы цен). Без параметра " +
      "определяется по последним проведённым документам и записям регистров",
  );

const momentField = (label: string) =>
  z
    .string()
    .regex(MOMENT_RE, "Дата в формате YYYY-MM-DD или YYYY-MM-DDTHH:mm:ss")
    .describe(`${label} (YYYY-MM-DD или YYYY-MM-DDTHH:mm:ss)`);

/** Сейчас по местному времени: дата без времени (00:00:00) раньше ночной выгрузки цен того же дня. */
function localNow(now = new Date()): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${localToday(now)}T${p(now.getHours())}:${p(now.getMinutes())}:${p(now.getSeconds())}`;
}
const moment = (v: string) => (v.length === 10 ? `${v}T00:00:00` : v);
const clean = (ref: string) => ref.replace(/[{}]/g, "").toLowerCase();
const decimals = (n: number) => (String(n).split(".")[1] ?? "").length;

/** Добавить поля к ответу createCore: и в structuredContent, и в текст. */
function withFields(result: Awaited<ReturnType<typeof createCore>>, extra: Record<string, unknown>) {
  if (result.isError || !result.structuredContent) return result;
  const structuredContent = { ...result.structuredContent, ...extra };
  return {
    ...result,
    content: [{ type: "text" as const, text: renderJson(structuredContent) }],
    structuredContent,
  };
}

/** Чтение по виду цен целиком; записей больше предела чтения — тем же чтением по ссылкам. */
async function overflowFallback<T>(wide: () => Promise<T>, byRefs: () => Promise<T>): Promise<T> {
  try {
    return await wide();
  } catch (e) {
    if (e instanceof AggregateOverflowError) return byRefs();
    throw e;
  }
}

interface PriceRow {
  item: string;
  ref: string;
  article?: string;
  characteristicRef?: string;
  current: number | null;
  currentSince?: string;
  price: number;
  change?: number;
  laterPriceFrom?: string;
  deletionMark?: boolean;
}

const cents = (n: number) => Math.round(n * 100);
/** Изменение в процентах к текущей цене; без текущей или при нулевой — нет. */
const changePct = (p: PriceRow) =>
  p.current ? Math.round(((p.price - p.current) / p.current) * 1000) / 10 : undefined;
/** Больше 20 % текущей — в копейках, без ошибок округления: ровно 20 % не отклонение. */
const overThreshold = (p: PriceRow) =>
  p.current !== null && Math.abs(cents(p.price) - cents(p.current)) * 100 > DEVIATION_PCT * cents(p.current);
const isDeviation = (p: PriceRow) =>
  !!p.deletionMark || !!p.laterPriceFrom || p.current === null || overThreshold(p);
/** Порядок отклонений в сжатом ответе: сначала то, что меняет смысл записи, затем крупные изменения. */
const severity = (p: PriceRow) =>
  p.deletionMark ? 0 : p.laterPriceFrom ? 1 : p.current === 0 ? 2 : overThreshold(p) ? 3 : 4;

/** Сводка по всем строкам документа — и когда prices в ответе неполный. */
function priceSummary(prices: readonly PriceRow[]) {
  const pcts = prices.map(changePct).filter((v): v is number => v !== undefined);
  const count = (f: (p: PriceRow) => boolean) => prices.filter(f).length;
  return {
    lines: prices.length,
    withoutCurrent: count((p) => p.current === null),
    up: count((p) => p.current !== null && cents(p.price) > cents(p.current)),
    down: count((p) => p.current !== null && cents(p.price) < cents(p.current)),
    unchanged: count((p) => p.current !== null && cents(p.price) === cents(p.current)),
    deviationThresholdPct: DEVIATION_PCT,
    overThreshold: count(overThreshold),
    ...(pcts.some((v) => v > 0) ? { maxIncreasePct: Math.max(...pcts) } : {}),
    ...(pcts.some((v) => v < 0) ? { maxDecreasePct: Math.min(...pcts) } : {}),
    laterPrices: count((p) => !!p.laterPriceFrom),
    deletionMarked: count((p) => !!p.deletionMark),
    deviations: count(isDeviation),
  };
}

const compactRow = (p: PriceRow, i: number) => {
  const pct = changePct(p);
  return {
    line: i + 1,
    item: p.item.length > COMPACT_NAME ? `${p.item.slice(0, COMPACT_NAME - 1)}…` : p.item,
    ...(p.article ? { article: p.article } : {}),
    current: p.current,
    price: p.price,
    ...(p.change !== undefined ? { change: p.change } : {}),
    ...(pct !== undefined ? { changePct: pct } : {}),
    ...(p.laterPriceFrom ? { laterPriceFrom: p.laterPriceFrom } : {}),
    ...(p.deletionMark ? { deletionMark: true } : {}),
  };
};

/**
 * Ответ в пределе при любом числе строк. Помещается — как в 1.1. Нет — по шагам, пока не поместится: строки документа
 * в view (они повторяют prices) → сжатые строки prices → в prices только отклонения, по важности, сколько войдёт.
 * Сводка summary всегда по всем строкам; неполный prices помечен pricesComplete=false.
 */
function fitPrices(
  result: Awaited<ReturnType<typeof createCore>>,
  prices: PriceRow[],
  created: boolean,
  maxResponseChars: number,
) {
  const sc = result.structuredContent as Record<string, unknown> | undefined;
  if (result.isError || !sc) return result;
  // Запись идёт вне контекста чтения (withToolCall): предел — из настройки, как pageCharsLimit у чтения.
  const limit = Math.min(maxResponseChars, MAX_RESPONSE_CHARS);
  const budget = (limit >= MAX_RESPONSE_CHARS ? limit : limit - 1_000) - RESPONSE_RESERVE;
  const size = (o: Record<string, unknown>) => renderJson(o).length;
  // Как в 1.1, пока ответ помещается в сам предел: запас нужен только сжатым ответам (ревью: при пределе 10 000
  // ответ 15–25 строк, помещавшийся в 1.1, терял таблицу view).
  if (size(sc) + OPERATION_ID_CHARS <= limit) return result;
  const notes = [...((sc["notes"] as string[] | undefined) ?? [])];
  let extra: Record<string, unknown> = {};
  const view = sc["view"] as PreviewView | undefined;
  if (view?.tables) {
    const tables = Object.fromEntries(
      Object.entries(view.tables).filter(([, t]) => t.length !== prices.length),
    );
    notes.push("Строки документа в view не повторены: они в prices.");
    extra = { view: { ...view, tables } };
    if (size({ ...sc, ...extra, notes }) <= budget) return withFields(result, { ...extra, notes });
  }
  const rows = prices.map(compactRow);
  const compactNote =
    "Строки prices сжаты: без ref, currentSince и характеристики, название до 60 знаков; line — номер строки во входе.";
  if (size({ ...sc, ...extra, prices: rows, notes: [...notes, compactNote] }) <= budget)
    return withFields(result, {
      ...extra,
      prices: rows,
      pricesFormat: "compact",
      notes: [...notes, compactNote],
    });
  const deviations = rows
    .map((r, i) => ({ r, p: prices[i]! }))
    .filter(({ p }) => isDeviation(p))
    .sort(
      (a, b) =>
        severity(a.p) - severity(b.p) ||
        Math.abs(b.r.changePct ?? 0) - Math.abs(a.r.changePct ?? 0) ||
        a.r.line - b.r.line,
    )
    .map(({ r }) => r);
  const where = created ? "в созданном документе в 1С" : "в документе в 1С после создания — он непроведённый";
  const kinds = `изменение больше ${DEVIATION_PCT} %, нет текущей цены, более поздняя цена, пометка удаления`;
  const note = (shown: number) =>
    `Строк ${prices.length} — все в ответ не помещаются. ` +
    (deviations.length
      ? `В prices только отклонения (${kinds}): ${shown} из ${deviations.length}` +
        `${shown < deviations.length ? ", самые важные первыми" : ""}. `
      : `Отклонений (${kinds}) нет, prices пуст. `) +
    `Счёт по всем строкам — summary; все строки видны ${where}.`;
  const fixed = { ...sc, ...extra, prices: [], notes: [...notes, compactNote, note(deviations.length)] };
  const kept = deviations.length ? fitToChars(deviations, size(fixed) + 100, budget) : [];
  return withFields(result, {
    ...extra,
    prices: kept,
    pricesFormat: "compact",
    pricesComplete: false,
    pricesTotal: prices.length,
    pricesShown: kept.length,
    deviationsTotal: deviations.length,
    notes: [...notes, compactNote, note(kept.length)],
  });
}

export function registerPricingTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "write.pricing.create_price_setting",
    {
      title: "Установить цены номенклатуры",
      description:
        "Создаёт документ «Установка цен номенклатуры» по одному виду цен — непроведённым; цены вступают в силу после " +
        "проведения (write.document.post_document, отдельный предпросмотр). Предпросмотр показывает по каждой позиции " +
        "текущую цену и новую (prices) и предупреждает: у вида цен есть зависимые виды (наценка на этот вид) — 1С их " +
        "при записи через OData не пересчитает; у позиции уже есть цена с более поздней датой — новая её не перекроет; " +
        "цена на тот же момент — отказ. date — момент цены, по умолчанию сейчас (по часам компьютера с коннектором); дата " +
        "без времени — начало дня. До 1000 строк на документ; не помещаются в ответ — строки prices сжимаются, а то и остаются только " +
        "отклонения (изменение больше 20 %, нет текущей цены, более поздняя цена, пометка удаления), счёт по всем строкам — summary. " +
        "Механизм определяется по базе (mechanism — явно): Управление торговлей 11 — прежний (legacy) или " +
        "ценообразование 2.5 (на живой базе не проверено); Бухгалтерия 3.0 — по типам цен (bp, вид цен — тип цен " +
        "номенклатуры). КА 2 и ERP 2 — по тем же полям, что УТ, без проверки.",
      inputSchema: z.object({
        database: databaseField,
        priceKind: z.string().trim().min(1).max(200).describe("Вид цен: Ref_Key или точное название"),
        lines: z
          .array(
            z.object({
              itemRef: z.string().trim().regex(GUID, "Ref_Key номенклатуры (GUID)"),
              price: z
                .number()
                .positive()
                .finite()
                .describe("Новая цена, не больше двух знаков после запятой"),
              characteristicRef: z
                .string()
                .trim()
                .regex(GUID, "Ref_Key характеристики (GUID)")
                .optional()
                .describe("Характеристика, если цены по характеристикам"),
            }),
          )
          .min(1)
          .max(MAX_LINES),
        date: momentField("Момент цены; по умолчанию сейчас").optional(),
        mechanism: mechanismField,
        comment: z.string().trim().max(500).optional().describe("Комментарий документа"),
        responsibleRef: z
          .string()
          .trim()
          .regex(GUID, "Ref_Key пользователя (GUID)")
          .optional()
          .describe(
            "Ответственный (Ref_Key пользователя 1С — найдите через read.catalog.list_users; помеченный, недействительный и служебный — отказ)",
          ),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, priceKind, lines, date, mechanism, comment, responsibleRef, confirm }) =>
      guard("write.pricing.create_price_setting", async () => {
        const conn = ctx.db(database);
        // Ответственный — первым: неверная ссылка не должна ждать разбора вида цен и строк.
        const responsible = await checkResponsible(conn, responsibleRef);
        const mech = await detectMechanism(conn, mechanism);
        const s = MECHANISMS[mech];
        const kinds = await readPriceKinds(conn, mech);
        if (!s.characteristic && lines.some((l) => l.characteristicRef))
          throw new InputError("В этой базе цены задаются без характеристик: уберите characteristicRef.");
        const kind = resolvePriceKind(priceKind, kinds);
        if (kind.deleted) throw new InputError(`Вид цен «${kind.name}» помечен на удаление.`);
        // Статусы вида цен в УТ: «Действует», «НеДействует» или пусто. Отказ — только на явном «НеДействует».
        if (kind.status === "НеДействует")
          throw new InputError(`Вид цен «${kind.name}» не действует (статус ${kind.status}).`);
        if (!kind.currencyRef) throw new InputError(`У вида цен «${kind.name}» не задана валюта цены.`);
        // Момент по умолчанию — время предпросмотра, как дата документов (docDate): подтверждение и безопасный повтор
        // после неизвестного исхода (1.1) отправляют ту же запись, даже если между ними прошла полночь.
        const op = currentWriteOperationId();
        const previewedAt = op ? (await conn.client.operationEntry(op))?.createdAt : undefined;
        const at = moment(date ?? localNow(previewedAt ? new Date(previewedAt) : undefined));

        // Строки: повтор позиции и характеристики, цены с лишними знаками — отказ на входе.
        const seen = new Set<string>();
        for (const l of lines) {
          const key = priceKey({
            itemRef: clean(l.itemRef),
            characteristicRef: l.characteristicRef && clean(l.characteristicRef),
            kindRef: kind.ref,
          });
          if (seen.has(key))
            throw new InputError(`Позиция ${l.itemRef} повторяется в строках: одна цена на позицию.`);
          seen.add(key);
          if (decimals(l.price) > 2)
            throw new InputError(`Цена ${l.price} — больше двух знаков после запятой.`);
        }
        const itemRefs = lines.map((l) => clean(l.itemRef));
        const nomSet = await nomenclatureSet(conn);
        if (!nomSet) throw new InputError("Справочник номенклатуры не опубликован: позиции не проверить.");
        const many = new Set(itemRefs).size > BY_REFS_UP_TO;
        // До 200 позиций — по ссылкам (до 80 строк — те же запросы, что в 1.1), больше — справочник целиком.
        const attrs = await readPageAttributes(conn, nomSet, itemRefs);
        const unknown = itemRefs.filter((r) => !attrs.has(r) || attrs.get(r)!.isFolder);
        if (unknown.length)
          throw new InputError(
            `Нет в справочнике номенклатуры (или это папки): ${unknown.slice(0, 10).join(", ")}.`,
          );

        // Что действует сейчас и что установлено позже.
        // Много строк — по виду цен целиком, один запрос; слишком много записей у вида — пачками, как раньше.
        const kindWide = <T>(wide: () => Promise<T>, byRefs: () => Promise<T>) =>
          many ? overflowFallback(wide, byRefs) : byRefs();
        const current = new Map(
          (
            await kindWide(
              () => slicePrices(conn, mech, { kindRefs: [kind.ref], at, lean: true }),
              () => slicePrices(conn, mech, { kindRefs: [kind.ref], itemRefs, at }),
            )
          )
            .filter((r) => seen.has(priceKey(r)))
            .map((r) => [priceKey(r), r]),
        );
        const later = await kindWide(
          () => laterPrices(conn, mech, { kindRef: kind.ref, after: at, lean: true }),
          () => laterPrices(conn, mech, { kindRef: kind.ref, itemRefs, after: at }),
        );
        const sameMoment = later.filter((r) => r.period === at && seen.has(priceKey(r)));
        if (sameMoment.length)
          throw new InputError(
            `На момент ${at} у ${sameMoment.length} позиций уже есть цена этого вида: в регистре один момент — одна ` +
              "цена. Укажите другой момент (date).",
          );
        const laterByKey = new Map<string, string>();
        for (const r of later) {
          const k = priceKey(r);
          if (seen.has(k) && (!laterByKey.has(k) || r.period < laterByKey.get(k)!))
            laterByKey.set(k, r.period);
        }

        const prices: PriceRow[] = lines.map((l) => {
          const key = priceKey({
            itemRef: clean(l.itemRef),
            characteristicRef: l.characteristicRef && clean(l.characteristicRef),
            kindRef: kind.ref,
          });
          const a = attrs.get(clean(l.itemRef));
          const was = current.get(key);
          return {
            item: a?.name ?? l.itemRef,
            ref: clean(l.itemRef),
            ...(a?.article ? { article: a.article } : {}),
            ...(l.characteristicRef ? { characteristicRef: clean(l.characteristicRef) } : {}),
            ...(was ? { current: was.price, currentSince: was.period } : { current: null }),
            price: l.price,
            ...(was ? { change: Math.round((l.price - was.price) * 100) / 100 } : {}),
            ...(laterByKey.has(key) ? { laterPriceFrom: laterByKey.get(key) } : {}),
            ...(a?.deletionMark ? { deletionMark: true } : {}),
          };
        });

        const warnings: string[] = [];
        const dependents = dependentKinds(kind, kinds);
        if (dependents.length)
          warnings.push(
            `От вида цен «${kind.name}» зависят: ${dependents.map((k) => k.name).join(", ")}. При записи через OData ` +
              "1С их не пересчитает: установите их цены отдельно или пересчитайте в 1С.",
          );
        const laterCount = prices.filter((p) => "laterPriceFrom" in p).length;
        if (laterCount)
          warnings.push(
            `У ${laterCount} позиций уже есть цена этого вида с более поздней датой (laterPriceFrom): новая цена ` +
              "действовала бы только до неё.",
          );
        const deleted = prices.filter((p) => p.deletionMark).length;
        if (deleted) warnings.push(`${deleted} позиций помечены на удаление.`);
        if (mech === "2.5")
          warnings.push(
            "Ценообразование 2.5: заполнение документа на живой базе не проверено, проверьте его в 1С.",
          );

        // УТ: статус, строка вида цен, вид цен и «цена изменена вручную» в строке — как в живых документах.
        // Бухгалтерия: тип цен в шапке, в строке — номенклатура, цена и валюта.
        const payload: Record<string, unknown> = {
          Date: at,
          ...(s.trade ? { Статус: "Согласован", Согласован: true } : {}),
          ...(s.headerKind ? { [s.headerKind]: kind.ref } : {}),
          ...(comment ? { Комментарий: comment } : {}),
          ...(responsible ? { Ответственный_Key: clean(responsible) } : {}),
          ...(s.trade ? { ВидыЦен: [{ LineNumber: 1, ВидЦены_Key: kind.ref }] } : {}),
          [s.rows]: lines.map((l, i) => ({
            LineNumber: i + 1,
            Номенклатура_Key: clean(l.itemRef),
            ...(s.characteristic
              ? { [s.characteristic]: l.characteristicRef ? clean(l.characteristicRef) : EMPTY_REF }
              : {}),
            ...Object.fromEntries(s.emptyKeys.map((k) => [k, EMPTY_REF])),
            ...(s.headerKind ? {} : { [s.kindKey]: kind.ref }),
            Цена: l.price,
            ...(s.trade ? { ЦенаИзмененаВручную: true } : {}),
            Валюта_Key: kind.currencyRef,
          })),
        };
        const result = await createCore(conn, PRICE_DOC, payload, confirm, warnings, () => undefined);
        // Строки документа в предпросмотре — только в prices: вдвоём с payload они не помещались в норму ответа.
        const compact =
          confirm || result.isError || !result.structuredContent
            ? result
            : withFields(result, {
                payload: {
                  ...((result.structuredContent as Record<string, unknown>)["payload"] as Record<
                    string,
                    unknown
                  >),
                  [s.rows]: `${lines.length} строк — см. prices`,
                },
              });
        const full = withFields(compact, {
          mechanism: mech,
          priceKind: { name: kind.name, ref: kind.ref },
          date: at,
          posted: false,
          prices,
          summary: priceSummary(prices),
          ...(confirm ? { next: "Цены вступят в силу после проведения: write.document.post_document." } : {}),
        });
        return fitPrices(full, prices, confirm, ctx.behavior.maxResponseChars);
      }),
  );

  server.registerTool(
    "read.analytics.get_prices",
    {
      title: "Цены номенклатуры",
      description:
        "Действующие цены по видам цен на момент (срез последних по регистру цен; в Бухгалтерии — по типам цен). priceKinds — " +
        "виды цен (Ref_Key или точное название; без параметра — все действующие). Отбор позиций как в " +
        "get_warehouse_stock: group или groups (с подгруппами), query (название или артикул), items (Ref_Key). " +
        "asOf — момент (YYYY-MM-DD — на начало дня, YYYY-MM-DDTHH:mm:ss — точно; без параметра — сейчас). Позиции " +
        "без цены вида в ответ не входят: полный список номенклатуры — read.catalog.list_items. " +
        "count — всего строк; длинный список читайте с offset=nextOffset.",
      inputSchema: z.object({
        database: databaseField,
        priceKinds: z.array(z.string().trim().min(1).max(200)).min(1).max(20).optional(),
        group: z
          .string()
          .trim()
          .min(1)
          .max(500)
          .optional()
          .describe("Группа номенклатуры: Ref_Key, название или путь"),
        groups: z
          .array(z.string().trim().min(1).max(500))
          .min(1)
          .max(MAX_ITEM_GROUPS)
          .optional()
          .describe("Несколько групп номенклатуры (вместо group)"),
        query: z.string().trim().min(1).max(200).optional().describe("Части слов названия или артикул"),
        items: z
          .array(z.string().trim().regex(GUID, "Ref_Key номенклатуры (GUID)"))
          .min(1)
          .max(500)
          .optional(),
        asOf: momentField("Момент цен").optional(),
        mechanism: mechanismField,
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько строк вернуть"),
      }),
      outputSchema: getPricesResultSchema,
    },
    ({ database, priceKinds, group, groups: groupList, query, items, asOf, mechanism, limit }) =>
      guard("read.analytics.get_prices", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const mech = await detectMechanism(conn, mechanism);
        const allKinds = await readPriceKinds(conn, mech);
        const kinds: PriceKind[] = priceKinds
          ? [...new Map(priceKinds.map((q) => resolvePriceKind(q, allKinds)).map((k) => [k.ref, k])).values()]
          : allKinds.filter((k) => !k.folder && !k.deleted && k.status !== "НеДействует");
        if (!kinds.length) throw new InputError("В базе нет действующих видов цен.");
        const at = asOf ? moment(asOf) : undefined;
        const itemRefs = items?.map(clean);
        const key = `prices:${mech}:${at ?? ""}:${kinds
          .map((k) => k.ref)
          .sort()
          .join(",")}:${(itemRefs ?? []).sort().join(",")}`;
        const rows = await pageCached(conn, key, callPageStart() === 0, () =>
          slicePrices(conn, mech, {
            kindRefs: kinds.map((k) => k.ref),
            ...(itemRefs ? { itemRefs } : {}),
            ...(at ? { at } : {}),
          }),
        );
        const nomSet = await nomenclatureSet(conn);
        const lines = rows.map((r) => ({ ...r }));
        const sel = await selectByItem(conn, nomSet, lines, { group, groups: groupList, query });
        const kindName = new Map(kinds.map((k) => [k.ref, k.name]));
        const sorted = sel.lines
          .filter((r) => kindName.has(r.kindRef))
          .sort(
            (a, b) =>
              a.itemRef.localeCompare(b.itemRef) ||
              (a.characteristicRef ?? "").localeCompare(b.characteristicRef ?? "") ||
              a.kindRef.localeCompare(b.kindRef),
          );
        const shown = sorted.slice(0, limit);
        const page = shown.slice(callPageStart());
        const attributes = sel.attributes;
        if (nomSet) {
          const unread = page.map((l) => l.itemRef).filter((r) => !attributes.has(r));
          for (const [ref, a] of await readPageAttributes(conn, nomSet, unread)) attributes.set(ref, a);
        }
        const notes = await selectionNotes(conn, nomSet, sel, { group, groups: groupList, query }, "цены");
        if (mech === "2.5") notes.push("Ценообразование 2.5: чтение цен на живой базе не проверено.");
        return ok({
          database: conn.cfg.name,
          mechanism: mech,
          ...(at ? { asOf: at } : {}),
          priceKinds: kinds.map((k) => ({ name: k.name, ref: k.ref })),
          count: sorted.length,
          ...(notes.length ? { note: notes.join(" ") } : {}),
          items: shown.map((r) => {
            const a = attributes.get(r.itemRef);
            return {
              item: a?.name ?? r.itemRef,
              ref: r.itemRef,
              ...(a?.article ? { article: a.article } : {}),
              ...(r.characteristicRef && r.characteristicRef !== EMPTY_REF
                ? { characteristicRef: r.characteristicRef }
                : {}),
              priceKind: kindName.get(r.kindRef)!,
              priceKindRef: r.kindRef,
              price: r.price,
              since: r.period,
              ...(r.recorderRef ? { recorderRef: r.recorderRef } : {}),
              ...(sel.matchedGroup?.has(r.itemRef)
                ? { matchedGroupRef: sel.matchedGroup.get(r.itemRef)! }
                : {}),
            };
          }),
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
