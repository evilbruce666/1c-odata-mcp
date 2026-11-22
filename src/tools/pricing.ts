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
  readItemAttributes,
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
import { currentWriteOperationId } from "../odata/write-operation-context.js";
import { createResultSchema, getPricesResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok, renderJson } from "./_shared.js";
import { confirmField, createCore } from "./write.js";

const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;
const MOMENT_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?$/;
/**
 * Строк в документе. Предпросмотр — одна строка ответа на позицию: на 500 строках он выходил за норму ответа, и клиент
 * терял operationId (ревью); 100 строк с названиями по 100 символов и более поздними ценами — около 44,5 тыс. из 45.
 * В ночных документах базы — до 32 строк; больше — несколькими документами.
 */
const MAX_LINES = 80;

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
        "без времени — начало дня. До 80 строк на документ. " +
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
        const attrs = await readItemAttributes(conn, nomSet, itemRefs);
        const unknown = itemRefs.filter((r) => !attrs.has(r) || attrs.get(r)!.isFolder);
        if (unknown.length)
          throw new InputError(
            `Нет в справочнике номенклатуры (или это папки): ${unknown.slice(0, 10).join(", ")}.`,
          );

        // Что действует сейчас и что установлено позже.
        const current = new Map(
          (await slicePrices(conn, mech, { kindRefs: [kind.ref], itemRefs, at })).map((r) => [
            priceKey(r),
            r,
          ]),
        );
        const later = await laterPrices(conn, mech, { kindRef: kind.ref, itemRefs, after: at });
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

        const prices = lines.map((l) => {
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
        return withFields(compact, {
          mechanism: mech,
          priceKind: { name: kind.name, ref: kind.ref },
          date: at,
          posted: false,
          prices,
          ...(confirm ? { next: "Цены вступят в силу после проведения: write.document.post_document." } : {}),
        });
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
