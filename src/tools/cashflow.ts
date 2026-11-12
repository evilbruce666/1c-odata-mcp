import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import {
  ok,
  fail,
  guard,
  databaseField,
  organizationField,
  dateField,
  offsetField,
  pageOf,
  snapshotField,
} from "./_shared.js";
import { and, cmp, contains, odataGuid } from "../odata/query.js";
import { CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { resolveOrganization } from "../odata/orgs.js";
import { resolveNames, num } from "../odata/accounting.js";
import { counterpartyRefsByKind, resolveCashflowItems } from "../odata/refilters.js";
import { TAX_PAYMENT_OPERATION, contractField, operationField } from "../odata/doc-fields.js";
import {
  collectDocuments,
  fetchAllForAggregation,
  emptyMeta,
  addMeta,
  type ScanMeta,
} from "../odata/aggregate.js";
import {
  paymentsBreakdownResultSchema,
  dealHistoryResultSchema,
  taxesPaidResultSchema,
} from "../schemas/output.js";

/** Нормализует GUID-строку для сравнения: без скобок, нижний регистр. */
function normGuid(v: unknown): string {
  return v ? String(v).replace(/[{}]/g, "").trim().toLowerCase() : "";
}

// Деньги копим в ЦЕЛЫХ КОПЕЙКАХ (float-сложение тысяч сумм даёт дрейф) и делим в конце.
const toCents = (v: unknown): number => Math.round(num(v) * 100);
const fromCents = (c: number): number => Math.round(c) / 100;

/** Множество имён полей сущности из $metadata (для детекта полиморфных полей). */
async function propsOf(conn: Connection, set: string): Promise<Set<string>> {
  const meta = await conn.getMetadata();
  return new Set((meta.entities.get(set)?.properties ?? []).map((p) => p.name));
}

/**
 * Имя поля контрагента в документе: типизированное `Контрагент_Key` (можно
 * фильтровать на сервере) или полиморфное `Контрагент` (Edm.String — фильтруется
 * только на клиенте, серверный eq/substringof по нему в 1С не работает).
 */
function counterpartyField(props: Set<string>): { field: string; serverFilterable: boolean } | undefined {
  if (props.has("Контрагент_Key")) return { field: "Контрагент_Key", serverFilterable: true };
  if (props.has("Контрагент")) return { field: "Контрагент", serverFilterable: false };
  return undefined;
}

/** OR-цепочка серверного фильтра по статьям ДДС (поле Edm.Guid — фильтруется). */
function cashflowItemFilter(refs: string[]): string | undefined {
  if (refs.length === 0) return undefined;
  if (refs.length === 1) return cmp("СтатьяДвиженияДенежныхСредств_Key", "eq", odataGuid(refs[0]!));
  return `(${refs.map((r) => `СтатьяДвиженияДенежныхСредств_Key eq ${odataGuid(r)}`).join(" or ")})`;
}

const GROUP_BY = ["operation", "month", "counterparty", "cashflowItem", "total"] as const;
const CP_KIND = ["ИП", "ЮрЛицо", "ФизЛицо", "Нерезидент", "Госорган"] as const;

/** Сумма-копилка по ключу разбивки (суммы — в копейках). */
interface Bucket {
  key: string;
  count: number;
  cents: number;
  inCents: number;
  outCents: number;
}

export function registerCashflowTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.analytics.get_payments_breakdown",
    {
      title: "Разбивка платежей за период",
      description:
        "Суммирует проведённые банковские и кассовые документы за период с разбивкой по выбранному " +
        "признаку (вид операции / месяц / контрагент / статья ДДС / итог). Фильтры: направление " +
        "(приход/расход), контрагент (Ref_Key), категория контрагента (ИП/ЮрЛицо/ФизЛицо/Нерезидент/" +
        "Госорган), статья ДДС (по коду/имени/Ref), подстрока в назначении платежа (напр. «аренда»), " +
        "вид операции. Отвечает на вопросы вроде «сколько процентов получили за год от такого-то банка», " +
        "«сколько заплатили ИП за год», «куда уходили деньги по статьям ДДС», «приход по месяцам». " +
        "Период — даты YYYY-MM-DD. Контрагент в банковских документах хранится полиморфно — фильтр " +
        "применяется на стороне коннектора; статья ДДС / назначение / период / организация — на сервере 1С. " +
        "Документ «Оплата платёжной картой» сюда не входит: это отражение в учёте, деньги приходят банковским поступлением; пустой ответ по контрагенту-покупателю не значит, что оплат нет (карточка 62). " +
        "Суммы — как зачислено: поступление, из которого банк удержал комиссию, показано за её вычетом. Отдельного поля " +
        "«карта или СБП» нет, различить можно только по формулировке банка в purposeContains. Оборот до комиссии и удержанную " +
        "комиссию по банкам и месяцам даёт read.accounting.get_account_card по счёту 57.03 с groupBy month, counterparty, " +
        "corrAccount: весь кредит — оборот, кредит с корреспонденцией 91.02 — комиссия. Комиссию, которую банк списывает " +
        "отдельным документом (обычно СБП), там не видно: она здесь, direction out, purposeContains с её формулировкой.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
        direction: z
          .enum(["in", "out", "both"])
          .default("both")
          .describe("in — приход, out — расход, both — оба (по умолчанию)"),
        counterpartyRef: z.string().optional().describe("Ref_Key контрагента (из find_counterparty)"),
        counterpartyKind: z
          .enum(CP_KIND)
          .optional()
          .describe(
            "Категория контрагента: ИП, ЮрЛицо, ФизЛицо (без ИП), Нерезидент, Госорган. " +
              "Берётся из справочника по полям ЮридическоеФизическоеЛицо / ИндивидуальныйПредприниматель и т.п.",
          ),
        cashflowItem: z
          .string()
          .optional()
          .describe("Статья ДДС: код (напр. «00-000006»), часть наименования (напр. «Аренда») или Ref_Key"),
        purposeContains: z
          .string()
          .optional()
          .describe("Подстрока в назначении платежа (напр. «процент», «депозит», «аренда»)"),
        operationType: z
          .string()
          .optional()
          .describe("Точный вид операции (ВидОперации), напр. «ПрочееПоступление», «Депозит»"),
        groupBy: z
          .enum(GROUP_BY)
          .default("operation")
          .describe(
            "Разбивка: operation (вид операции), month (месяц), counterparty (контрагент), " +
              "cashflowItem (статья ДДС), total (без разбивки)",
          ),
      }),
      outputSchema: paymentsBreakdownResultSchema,
    },
    ({
      database,
      organization,
      from,
      to,
      direction,
      counterpartyRef,
      counterpartyKind,
      cashflowItem,
      purposeContains,
      operationType,
      groupBy,
    }) =>
      guard("read.analytics.get_payments_breakdown", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const available = await conn.available();

        const kindSet = counterpartyKind ? await counterpartyRefsByKind(conn, counterpartyKind) : undefined;
        const itemMatches = cashflowItem ? await resolveCashflowItems(conn, cashflowItem) : [];
        if (cashflowItem && itemMatches.length === 0) {
          return fail(`Статья ДДС «${cashflowItem}» не найдена в справочнике.`);
        }
        const itemRefs = itemMatches.map((m) => m.ref);

        const inKeys: string[] = [...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn];
        const outKeys: string[] = [...DOCUMENTS.bankOut, ...DOCUMENTS.cashOut];
        const wantKeys: string[] =
          direction === "in" ? inKeys : direction === "out" ? outKeys : [...inKeys, ...outKeys];
        const inSets = new Set(inKeys);

        const cpRef = counterpartyRef ? normGuid(counterpartyRef) : undefined;
        const opType = operationType?.trim();

        // План по каждому опубликованному виду документа (без даты — её добавит collectDocuments).
        interface Plan {
          set: string;
          isIn: boolean;
          cp: ReturnType<typeof counterpartyField>;
          hasCfi: boolean;
          op: string | undefined;
          baseFilter: string | undefined;
          select: string[];
        }
        const plans: Plan[] = [];
        for (const candidate of wantKeys) {
          const set = resolveEntity([candidate], available);
          if (!set) continue;
          const props = await propsOf(conn, set);
          const hasPurpose = props.has("НазначениеПлатежа");
          if (purposeContains && !hasPurpose) continue;
          const hasCfi = props.has("СтатьяДвиженияДенежныхСредств_Key");
          if (cashflowItem && !hasCfi) continue;
          const cp = counterpartyField(props);
          const op = operationField(props);

          const baseFilter =
            and(
              org ? cmp(DOC_FIELDS.organization, "eq", odataGuid(org.ref)) : undefined,
              cmp(DOC_FIELDS.posted, "eq", "true"),
              purposeContains && hasPurpose ? contains("НазначениеПлатежа", purposeContains) : undefined,
              cpRef && cp?.serverFilterable ? cmp(cp.field, "eq", odataGuid(cpRef)) : undefined,
              hasCfi ? cashflowItemFilter(itemRefs) : undefined,
            ) || undefined;

          const select: string[] = [DOC_FIELDS.date, DOC_FIELDS.amount];
          if (cp) select.push(cp.field);
          if (op) select.push(op);
          if (hasCfi) select.push("СтатьяДвиженияДенежныхСредств_Key");

          plans.push({ set, isIn: inSets.has(candidate), cp, hasCfi, op, baseFilter, select });
        }
        if (plans.length === 0) {
          return fail("Банковские/кассовые документы не опубликованы в OData. Добавьте их в «Состав OData».");
        }

        // Параллельная выборка по видам документов (авто-чанкинг + громкое переполнение внутри).
        const results = await Promise.all(
          plans.map((p) =>
            collectDocuments(conn, p.set, {
              baseFilter: p.baseFilter,
              dateField: DOC_FIELDS.date,
              from,
              to,
              select: p.select,
            }).then((r) => ({ plan: p, rows: r.rows, meta: r.meta })),
          ),
        );

        const buckets = new Map<string, Bucket>();
        const bump = (key: string, cents: number, isIn: boolean): void => {
          const b = buckets.get(key) ?? { key, count: 0, cents: 0, inCents: 0, outCents: 0 };
          b.count += 1;
          b.cents += cents;
          if (isIn) b.inCents += cents;
          else b.outCents += cents;
          buckets.set(key, b);
        };
        let inflowC = 0,
          outflowC = 0,
          docCount = 0;
        let meta: ScanMeta = emptyMeta();
        const cpKeys = new Set<string>();
        const cfiKeys = new Set<string>();

        for (const { plan, rows, meta: m } of results) {
          meta = addMeta(meta, m);
          for (const r of rows) {
            const cpNorm = plan.cp ? normGuid(r[plan.cp.field]) : "";
            if (cpRef && !(plan.cp?.serverFilterable ?? false) && cpNorm !== cpRef) continue;
            if (kindSet && (!cpNorm || !kindSet.has(cpNorm))) continue;
            const operation = plan.op ? String(r[plan.op] ?? "") : "";
            if (opType && operation !== opType) continue;

            const cents = toCents(r[DOC_FIELDS.amount]);
            docCount += 1;
            if (plan.isIn) inflowC += cents;
            else outflowC += cents;

            let key: string;
            if (groupBy === "operation") key = operation || "—";
            else if (groupBy === "month") key = String(r[DOC_FIELDS.date] ?? "").slice(0, 7);
            else if (groupBy === "counterparty") {
              key = cpNorm;
              if (key) cpKeys.add(key);
            } else if (groupBy === "cashflowItem") {
              key = String(r["СтатьяДвиженияДенежныхСредств_Key"] ?? "").toLowerCase();
              if (key) cfiKeys.add(key);
            } else key = "ИТОГО";
            bump(key, cents, plan.isIn);
          }
        }

        const names = new Map<string, string>();
        if (groupBy === "counterparty" && cpKeys.size) {
          const cpSet = resolveEntity(CATALOGS.counterparties, available);
          if (cpSet) for (const [k, v] of await resolveNames(conn, cpSet, cpKeys)) names.set(normGuid(k), v);
        }
        if (groupBy === "cashflowItem" && cfiKeys.size) {
          const itemSet = resolveEntity(CATALOGS.cashflowItems, available);
          if (itemSet)
            for (const [k, v] of await resolveNames(conn, itemSet, cfiKeys)) names.set(normGuid(k), v);
        }

        const labelOf = (b: Bucket): string => {
          if (groupBy === "counterparty") return b.key ? (names.get(b.key) ?? b.key) : "Без контрагента";
          if (groupBy === "cashflowItem") return b.key ? (names.get(b.key) ?? b.key) : "Без статьи ДДС";
          return b.key || "—";
        };
        const groups = [...buckets.values()]
          .map((b) => ({
            label: labelOf(b),
            count: b.count,
            sum: fromCents(b.cents),
            ...(direction === "both" ? { inflow: fromCents(b.inCents), outflow: fromCents(b.outCents) } : {}),
          }))
          // Месяцы — по порядку: это ряд во времени, а не рейтинг.
          .sort((a, b) =>
            groupBy === "month"
              ? a.label.localeCompare(b.label)
              : b.sum - a.sum || a.label.localeCompare(b.label),
          );

        return ok({
          database: conn.cfg.name,
          organization: org?.name,
          period: { from, to },
          direction,
          filters: {
            ...(counterpartyRef ? { counterpartyRef } : {}),
            ...(counterpartyKind ? { counterpartyKind, counterpartyCount: kindSet?.size ?? 0 } : {}),
            ...(cashflowItem
              ? {
                  cashflowItem,
                  resolvedItems: itemMatches.map((m) => ({ code: m.code, name: m.name, ref: m.ref })),
                }
              : {}),
            ...(purposeContains ? { purposeContains } : {}),
            ...(operationType ? { operationType } : {}),
          },
          groupBy,
          documents: docCount,
          inflow: fromCents(inflowC),
          outflow: fromCents(outflowC),
          net: fromCents(inflowC - outflowC),
          groups,
          scan: { documentsScanned: meta.rowsScanned, windows: meta.chunks, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  // === Класс 5: история по сделке/договору ===
  server.registerTool(
    "read.analytics.get_deal_history",
    {
      title: "Движения по сделке/договору",
      description:
        "Хронология всех проведённых банковских и кассовых движений по конкретной сделке либо договору: " +
        "приход, расход, итоги и сальдо. Сделку можно задать кодом из назначения платежа (dealKey, " +
        "напр. «CB12812240037987» — банковский идентификатор депозита/займа/контракта в назначении) " +
        "или ссылкой на договор (contractRef, Ref_Key из find_counterparty + договор). Обязателен ровно " +
        "один из них. Период — опционально (без него — всё время).",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        dealKey: z
          .string()
          .optional()
          .describe(
            "Код сделки/контракта/документа в назначении платежа (напр. «CB12812240037987»). " +
              "Серверный фильтр substringof по НазначениеПлатежа.",
          ),
        contractRef: z
          .string()
          .optional()
          .describe("Ref_Key договора (для документов с ДоговорКонтрагента_Key)"),
        from: dateField("Дата начала").optional(),
        to: dateField("Дата конца").optional(),
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько событий вернуть"),
        offset: offsetField,
        snapshot: snapshotField,
      }),
      outputSchema: dealHistoryResultSchema,
    },
    ({ database, organization, dealKey, contractRef, from, to, limit, offset, snapshot }) =>
      guard("read.analytics.get_deal_history", async () => {
        const t0 = Date.now();
        if (!dealKey && !contractRef) return fail("Укажите dealKey или contractRef.");
        if (dealKey && contractRef) return fail("dealKey и contractRef взаимоисключающие — оставьте один.");
        if (from && to && from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const available = await conn.available();

        const allKeys: string[] = [
          ...DOCUMENTS.bankIn,
          ...DOCUMENTS.cashIn,
          ...DOCUMENTS.bankOut,
          ...DOCUMENTS.cashOut,
        ];
        const inSets = new Set<string>([...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn]);
        const contractSet = resolveEntity(CATALOGS.contracts, available);

        interface Plan {
          set: string;
          isIn: boolean;
          cp: ReturnType<typeof counterpartyField>;
          op: string | undefined;
          filter: string | undefined;
          select: string[];
        }
        const plans: Plan[] = [];
        for (const candidate of allKeys) {
          const set = resolveEntity([candidate], available);
          if (!set) continue;
          const props = await propsOf(conn, set);
          const hasPurpose = props.has("НазначениеПлатежа");
          const contract = contractField(props, contractSet);
          if (dealKey && !hasPurpose) continue;
          if (contractRef && !contract) continue;
          const op = operationField(props);
          const cp = counterpartyField(props);
          const filter =
            and(
              cmp(DOC_FIELDS.posted, "eq", "true"),
              from ? cmp(DOC_FIELDS.date, "ge", `datetime'${from}T00:00:00'`) : undefined,
              to ? cmp(DOC_FIELDS.date, "le", `datetime'${to}T23:59:59'`) : undefined,
              org ? cmp(DOC_FIELDS.organization, "eq", odataGuid(org.ref)) : undefined,
              dealKey ? contains("НазначениеПлатежа", dealKey) : undefined,
              contractRef && contract ? contract.filter(contractRef) : undefined,
            ) || undefined;
          const select: string[] = [DOC_FIELDS.date, DOC_FIELDS.number, DOC_FIELDS.amount];
          if (op) select.push(op);
          if (hasPurpose) select.push("НазначениеПлатежа");
          if (cp) select.push(cp.field);
          plans.push({ set, isIn: inSets.has(candidate), cp, op, filter, select });
        }

        // Сделка/договор узкие по своей природе — одиночная выборка с громким переполнением.
        const results = await Promise.all(
          plans.map((p) =>
            fetchAllForAggregation(
              conn,
              p.set,
              { filter: p.filter, select: p.select, orderby: `${DOC_FIELDS.date} asc` },
              p.set,
            ).then((r) => ({ plan: p, rows: r.rows, meta: r.meta })),
          ),
        );

        interface Event {
          date: string;
          entitySet: string;
          number: string;
          cents: number;
          direction: "in" | "out";
          operation?: string;
          counterpartyRef?: string;
          purpose?: string;
        }
        const events: Event[] = [];
        const cpKeys = new Set<string>();
        let meta: ScanMeta = emptyMeta();
        for (const { plan, rows, meta: m } of results) {
          meta = addMeta(meta, m);
          for (const r of rows) {
            const cpVal = plan.cp ? normGuid(r[plan.cp.field]) : "";
            if (cpVal) cpKeys.add(cpVal);
            events.push({
              date: String(r[DOC_FIELDS.date] ?? ""),
              entitySet: plan.set,
              number: String(r[DOC_FIELDS.number] ?? ""),
              cents: toCents(r[DOC_FIELDS.amount]),
              direction: plan.isIn ? "in" : "out",
              operation: plan.op && r[plan.op] ? String(r[plan.op]) : undefined,
              counterpartyRef: cpVal || undefined,
              purpose: r["НазначениеПлатежа"] ? String(r["НазначениеПлатежа"]) : undefined,
            });
          }
        }

        const cpNames = new Map<string, string>();
        if (cpKeys.size) {
          const cpSet = resolveEntity(CATALOGS.counterparties, available);
          if (cpSet)
            for (const [k, v] of await resolveNames(conn, cpSet, cpKeys)) cpNames.set(k.toLowerCase(), v);
        }

        events.sort(
          (a, b) =>
            a.date.localeCompare(b.date) ||
            a.entitySet.localeCompare(b.entitySet) ||
            String(a.number).localeCompare(String(b.number)),
        );
        const inflowC = events.filter((e) => e.direction === "in").reduce((s, e) => s + e.cents, 0);
        const outflowC = events.filter((e) => e.direction === "out").reduce((s, e) => s + e.cents, 0);
        const allItems = events.map((e) => ({
          date: e.date.slice(0, 10),
          entitySet: e.entitySet,
          number: e.number,
          direction: e.direction,
          amount: fromCents(e.cents),
          operation: e.operation,
          counterparty: e.counterpartyRef ? (cpNames.get(e.counterpartyRef) ?? e.counterpartyRef) : undefined,
          purpose: e.purpose,
        }));
        const head = {
          database: conn.cfg.name,
          organization: org?.name,
          ...(from || to ? { period: { from, to } } : {}),
          filters: { ...(dealKey ? { dealKey } : {}), ...(contractRef ? { contractRef } : {}) },
          events: events.length,
          inflow: fromCents(inflowC),
          outflow: fromCents(outflowC),
          net: fromCents(inflowC - outflowC),
        };
        const { items, page, note } = pageOf(allItems, {
          offset,
          limit,
          head,
          unit: "события",
          snapshot,
          tracksChanges: true,
        });
        return ok({
          ...head,
          ...page,
          items,
          ...(note ? { note: `${note} Итоги посчитаны по всем событиям.` } : {}),
          scan: { documentsScanned: meta.rowsScanned, windows: meta.chunks, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  // === Класс 6: уплаченные налоги/взносы ===
  server.registerTool(
    "read.analytics.get_taxes_paid",
    {
      title: "Уплаченные налоги и взносы за период",
      description:
        "Сумма уплаченных налогов/страховых взносов за период: исходящие банковские списания с " +
        "ВидОперации=«ПеречислениеНалога» (это надёжный маркер уплаты налога в БП 3.0; покрывает " +
        "ЕНП, НДФЛ, страховые взносы, налог УСН/ПСН, налог на прибыль и т.п.). По умолчанию — итог; " +
        "разбивка по статье ДДС / месяцу / получателю (обычно ФНС/Казначейство) / итог. Параметр " +
        "cashflowItem уточняет конкретный вид налога (напр. «УСН», «Страховые», «НДФЛ», «ЕНП»). " +
        "Отвечает на вопросы вроде «сколько налогов мы заплатили в прошлом году», «сколько ушло " +
        "страховых взносов в квартал», «помесячная нагрузка по налогам». Период — даты YYYY-MM-DD.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
        cashflowItem: z
          .string()
          .optional()
          .describe("Уточнить статью ДДС (напр. «УСН», «Страховые», «НДФЛ», «ЕНП», код «00-000010»)"),
        groupBy: z
          .enum(["cashflowItem", "month", "counterparty", "total"])
          .default("cashflowItem")
          .describe("Разбивка: cashflowItem (по налогу) / month / counterparty / total"),
      }),
      outputSchema: taxesPaidResultSchema,
    },
    ({ database, organization, from, to, cashflowItem, groupBy }) =>
      guard("read.analytics.get_taxes_paid", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const available = await conn.available();
        const itemMatches = cashflowItem ? await resolveCashflowItems(conn, cashflowItem) : [];
        if (cashflowItem && itemMatches.length === 0) return fail(`Статья ДДС «${cashflowItem}» не найдена.`);
        const itemRefs = itemMatches.map((m) => m.ref);

        interface Plan {
          set: string;
          cp: ReturnType<typeof counterpartyField>;
          hasCfi: boolean;
          baseFilter: string | undefined;
          select: string[];
        }
        const plans: Plan[] = [];
        for (const candidate of [...DOCUMENTS.bankOut, "Document_ПлатежноеПоручение"]) {
          const set = resolveEntity([candidate], available);
          if (!set) continue;
          const props = await propsOf(conn, set);
          const op = operationField(props);
          if (!op) continue;
          const hasCfi = props.has("СтатьяДвиженияДенежныхСредств_Key");
          if (cashflowItem && !hasCfi) continue;
          const cp = counterpartyField(props);
          const baseFilter =
            and(
              cmp(DOC_FIELDS.posted, "eq", "true"),
              org ? cmp(DOC_FIELDS.organization, "eq", odataGuid(org.ref)) : undefined,
              // КЛЮЧЕВОЙ маркер уплаты налога: в БП и УТ значение своё.
              cmp(op, "eq", `'${TAX_PAYMENT_OPERATION[op]}'`),
              hasCfi ? cashflowItemFilter(itemRefs) : undefined,
            ) || undefined;
          const select: string[] = [DOC_FIELDS.date, DOC_FIELDS.amount];
          if (hasCfi) select.push("СтатьяДвиженияДенежныхСредств_Key");
          if (cp) select.push(cp.field);
          plans.push({ set, cp, hasCfi, baseFilter, select });
        }

        const results = await Promise.all(
          plans.map((p) =>
            collectDocuments(conn, p.set, {
              baseFilter: p.baseFilter,
              dateField: DOC_FIELDS.date,
              from,
              to,
              select: p.select,
            }).then((r) => ({ plan: p, rows: r.rows, meta: r.meta })),
          ),
        );

        const buckets = new Map<string, Bucket>();
        const bump = (key: string, cents: number): void => {
          const b = buckets.get(key) ?? { key, count: 0, cents: 0, inCents: 0, outCents: 0 };
          b.count += 1;
          b.cents += cents;
          buckets.set(key, b);
        };
        let totalC = 0,
          docCount = 0;
        let meta: ScanMeta = emptyMeta();
        const cfiKeys = new Set<string>();
        const cpKeys = new Set<string>();
        for (const { plan, rows, meta: m } of results) {
          meta = addMeta(meta, m);
          for (const r of rows) {
            const cents = toCents(r[DOC_FIELDS.amount]);
            totalC += cents;
            docCount += 1;
            const cfi = plan.hasCfi ? String(r["СтатьяДвиженияДенежныхСредств_Key"] ?? "").toLowerCase() : "";
            const cpVal = plan.cp ? normGuid(r[plan.cp.field]) : "";
            if (cfi) cfiKeys.add(cfi);
            if (cpVal) cpKeys.add(cpVal);
            let key: string;
            if (groupBy === "cashflowItem") key = cfi;
            else if (groupBy === "month") key = String(r[DOC_FIELDS.date] ?? "").slice(0, 7);
            else if (groupBy === "counterparty") key = cpVal;
            else key = "ИТОГО";
            bump(key, cents);
          }
        }

        const names = new Map<string, string>();
        if (groupBy === "cashflowItem" && cfiKeys.size) {
          const itemSet = resolveEntity(CATALOGS.cashflowItems, available);
          if (itemSet)
            for (const [k, v] of await resolveNames(conn, itemSet, cfiKeys)) names.set(k.toLowerCase(), v);
        }
        if (groupBy === "counterparty" && cpKeys.size) {
          const cpSet = resolveEntity(CATALOGS.counterparties, available);
          if (cpSet)
            for (const [k, v] of await resolveNames(conn, cpSet, cpKeys)) names.set(k.toLowerCase(), v);
        }

        const labelOf = (b: Bucket): string => {
          if (groupBy === "cashflowItem") return b.key ? (names.get(b.key) ?? b.key) : "Без статьи ДДС";
          if (groupBy === "counterparty") return b.key ? (names.get(b.key) ?? b.key) : "Без контрагента";
          return b.key || "—";
        };
        const groups = [...buckets.values()]
          .map((b) => ({ label: labelOf(b), count: b.count, sum: fromCents(b.cents) }))
          .sort((a, b) => b.sum - a.sum || a.label.localeCompare(b.label));

        return ok({
          database: conn.cfg.name,
          organization: org?.name,
          period: { from, to },
          filters: {
            ...(cashflowItem
              ? {
                  cashflowItem,
                  resolvedItems: itemMatches.map((m) => ({ code: m.code, name: m.name, ref: m.ref })),
                }
              : {}),
          },
          groupBy,
          documents: docCount,
          total: fromCents(totalC),
          groups,
          scan: { documentsScanned: meta.rowsScanned, windows: meta.chunks, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
