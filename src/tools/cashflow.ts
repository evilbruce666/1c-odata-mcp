import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, contains, odataGuid } from "../odata/query.js";
import { CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { resolveOrganization } from "../odata/orgs.js";
import { resolveNames, num } from "../odata/accounting.js";
import { counterpartyRefsByKind, resolveCashflowItems } from "../odata/refilters.js";

/** Нормализует GUID-строку для сравнения: без скобок, нижний регистр. */
function normGuid(v: unknown): string {
  return v ? String(v).replace(/[{}]/g, "").trim().toLowerCase() : "";
}

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

const GROUP_BY = ["operation", "month", "counterparty", "cashflowItem", "total"] as const;
const CP_KIND = ["ИП", "ЮрЛицо", "ФизЛицо", "Нерезидент", "Госорган"] as const;

interface Bucket {
  key: string;
  count: number;
  sum: number;
  inSum: number;
  outSum: number;
}

export function registerCashflowTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "get_payments_breakdown",
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
        "применяется на стороне коннектора; статья ДДС / назначение / период / организация — на сервере 1С.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        from: z.string().describe("Дата начала периода (YYYY-MM-DD)"),
        to: z.string().describe("Дата конца периода (YYYY-MM-DD)"),
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
      },
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
      guard("get_payments_breakdown", async () => {
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const orgKey = org?.ref;
        const orgName = org?.name;
        const available = await conn.available();

        // Резолв «срезов» справочников — поднимаем ОДИН раз, дальше клиентский фильтр.
        const kindSet = counterpartyKind ? await counterpartyRefsByKind(conn, counterpartyKind) : undefined;
        const itemMatches = cashflowItem ? await resolveCashflowItems(conn, cashflowItem) : [];
        if (cashflowItem && itemMatches.length === 0) {
          return fail(`Статья ДДС «${cashflowItem}» не найдена в справочнике.`);
        }

        const inKeys: string[] = [...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn];
        const outKeys: string[] = [...DOCUMENTS.bankOut, ...DOCUMENTS.cashOut];
        const wantKeys: string[] =
          direction === "in" ? inKeys : direction === "out" ? outKeys : [...inKeys, ...outKeys];
        const inSets = new Set(inKeys);

        const cpRef = counterpartyRef ? normGuid(counterpartyRef) : undefined;
        const opType = operationType?.trim();
        const buckets = new Map<string, Bucket>();
        const bump = (key: string, amount: number, isIn: boolean): void => {
          const b = buckets.get(key) ?? { key, count: 0, sum: 0, inSum: 0, outSum: 0 };
          b.count += 1;
          b.sum += amount;
          if (isIn) b.inSum += amount;
          else b.outSum += amount;
          buckets.set(key, b);
        };

        let inflow = 0,
          outflow = 0,
          docCount = 0,
          truncated = false;
        const usedSets: string[] = [];
        const cpKeys = new Set<string>();
        const cfiKeys = new Set<string>();

        for (const candidate of wantKeys) {
          const set = resolveEntity([candidate], available);
          if (!set) continue;
          const props = await propsOf(conn, set);
          const hasPurpose = props.has("НазначениеПлатежа");
          // Если ищем по назначению, а его нет (кассовые ордера) — этот вид пропускаем.
          if (purposeContains && !hasPurpose) continue;
          const cp = counterpartyField(props);
          const hasOp = props.has("ВидОперации");
          const hasCfi = props.has("СтатьяДвиженияДенежныхСредств_Key");
          // Если просили статью ДДС, а её в этом виде документа нет — пропускаем.
          if (cashflowItem && !hasCfi) continue;

          // Серверный фильтр по статье ДДС (Edm.Guid → можно). Для нескольких
          // совпадений по подстроке собираем OR-цепочку.
          const cfiFilter =
            hasCfi && itemMatches.length
              ? itemMatches.length === 1
                ? cmp("СтатьяДвиженияДенежныхСредств_Key", "eq", odataGuid(itemMatches[0]!.ref))
                : `(${itemMatches
                    .map((m) => `СтатьяДвиженияДенежныхСредств_Key eq ${odataGuid(m.ref)}`)
                    .join(" or ")})`
              : undefined;

          const filter = and(
            cmp(DOC_FIELDS.date, "ge", `datetime'${from}T00:00:00'`),
            cmp(DOC_FIELDS.date, "le", `datetime'${to}T23:59:59'`),
            orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
            cmp(DOC_FIELDS.posted, "eq", "true"),
            purposeContains && hasPurpose ? contains("НазначениеПлатежа", purposeContains) : undefined,
            // Типизированный контрагент можно отфильтровать на сервере.
            cpRef && cp?.serverFilterable ? cmp(cp.field, "eq", odataGuid(cpRef)) : undefined,
            cfiFilter,
          );

          const select: string[] = [DOC_FIELDS.date, DOC_FIELDS.amount];
          if (cp) select.push(cp.field);
          if (hasOp) select.push("ВидОперации");
          if (hasCfi) select.push("СтатьяДвиженияДенежныхСредств_Key");

          const { rows, truncated: t } = await fetchAll(
            conn.client,
            set,
            { filter, select, orderby: `${DOC_FIELDS.date} asc` },
            conn.behavior.pageSize,
            conn.behavior.maxRows,
          );
          if (t) truncated = true;
          usedSets.push(set);
          const isIn = inSets.has(candidate);

          for (const r of rows) {
            const cpVal = cp ? r[cp.field] : undefined;
            const cpNorm = normGuid(cpVal);
            // Клиентская фильтрация по контрагенту для полиморфного поля.
            if (cpRef && !(cp?.serverFilterable ?? false) && cpNorm !== cpRef) continue;
            // Клиентский фильтр по категории контрагента (ИП и т.п.).
            if (kindSet && (!cpNorm || !kindSet.has(cpNorm))) continue;
            if (opType && String(r["ВидОперации"] ?? "") !== opType) continue;

            const amount = num(r[DOC_FIELDS.amount]);
            docCount += 1;
            if (isIn) inflow += amount;
            else outflow += amount;

            let key: string;
            if (groupBy === "operation") key = String(r["ВидОперации"] ?? "—");
            else if (groupBy === "month") key = String(r[DOC_FIELDS.date] ?? "").slice(0, 7);
            else if (groupBy === "counterparty") {
              key = cpNorm;
              if (key) cpKeys.add(key);
            } else if (groupBy === "cashflowItem") {
              key = String(r["СтатьяДвиженияДенежныхСредств_Key"] ?? "").toLowerCase();
              if (key) cfiKeys.add(key);
            } else key = "ИТОГО";
            bump(key, amount, isIn);
          }
        }

        if (usedSets.length === 0) {
          return fail("Банковские/кассовые документы не опубликованы в OData. Добавьте их в «Состав OData».");
        }

        // Имена для меток разбивки.
        const names = new Map<string, string>();
        if (groupBy === "counterparty" && cpKeys.size) {
          const cpSet = resolveEntity(CATALOGS.counterparties, available);
          if (cpSet) {
            const resolved = await resolveNames(conn, cpSet, cpKeys);
            for (const [k, v] of resolved) names.set(normGuid(k), v);
          }
        }
        if (groupBy === "cashflowItem" && cfiKeys.size) {
          const itemSet = resolveEntity(CATALOGS.cashflowItems, available);
          if (itemSet) {
            const resolved = await resolveNames(conn, itemSet, cfiKeys);
            for (const [k, v] of resolved) names.set(normGuid(k), v);
          }
        }

        const round = (n: number): number => Math.round(n * 100) / 100;
        const labelOf = (b: Bucket): string => {
          if (groupBy === "counterparty") return b.key ? (names.get(b.key) ?? b.key) : "Без контрагента";
          if (groupBy === "cashflowItem") return b.key ? (names.get(b.key) ?? b.key) : "Без статьи ДДС";
          return b.key || "—";
        };
        const groups = [...buckets.values()]
          .map((b) => ({
            label: labelOf(b),
            count: b.count,
            sum: round(b.sum),
            ...(direction === "both" ? { inflow: round(b.inSum), outflow: round(b.outSum) } : {}),
          }))
          .sort((a, b) => b.sum - a.sum);

        return ok({
          database: conn.cfg.name,
          organization: orgName,
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
          inflow: round(inflow),
          outflow: round(outflow),
          net: round(inflow - outflow),
          groups,
          ...(truncated
            ? {
                truncated: true,
                note: `Данные усечены лимитом ${conn.behavior.maxRows} строк на вид документа — сузьте период или фильтр.`,
              }
            : {}),
        });
      }),
  );
}
