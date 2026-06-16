import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { resolveOrganization } from "../odata/orgs.js";
import { resolveNames, num } from "../odata/accounting.js";
import { counterpartyRefsByKind, ANALYTICS_PAGE, type CounterpartyKind } from "../odata/refilters.js";

/**
 * Агрегаторы продаж и закупок. У документов реализации/поступления контрагент
 * типизированный (Контрагент_Key — Edm.Guid), фильтр идёт на сервере 1С.
 */

const SALES_GROUP_BY = ["counterparty", "month", "contract", "total"] as const;
const CP_KIND = ["ИП", "ЮрЛицо", "ФизЛицо", "Нерезидент", "Госорган"] as const;

interface Bucket {
  key: string;
  count: number;
  sum: number;
}

interface Args {
  database?: string | undefined;
  organization?: string | undefined;
  from: string;
  to: string;
  counterpartyRef?: string | undefined;
  counterpartyKind?: CounterpartyKind | undefined;
  contractRef?: string | undefined;
  groupBy: (typeof SALES_GROUP_BY)[number];
  limit: number;
}

async function aggregate(
  conn: Connection,
  docCandidates: readonly string[],
  humanName: string,
  args: Args,
): Promise<ReturnType<typeof ok>> {
  const available = await conn.available();
  const set = resolveEntity(docCandidates, available);
  if (!set) {
    return fail(`Документ «${humanName}» не опубликован в OData. Добавьте в «Состав OData».`);
  }
  const org = args.organization ? await resolveOrganization(conn, args.organization) : undefined;
  const kindSet = args.counterpartyKind
    ? await counterpartyRefsByKind(conn, args.counterpartyKind)
    : undefined;

  const filter = and(
    cmp(DOC_FIELDS.date, "ge", `datetime'${args.from}T00:00:00'`),
    cmp(DOC_FIELDS.date, "le", `datetime'${args.to}T23:59:59'`),
    org ? cmp(DOC_FIELDS.organization, "eq", odataGuid(org.ref)) : undefined,
    cmp(DOC_FIELDS.posted, "eq", "true"),
    args.counterpartyRef ? cmp(DOC_FIELDS.counterparty, "eq", odataGuid(args.counterpartyRef)) : undefined,
    args.contractRef ? cmp("ДоговорКонтрагента_Key", "eq", odataGuid(args.contractRef)) : undefined,
  );

  const select = [DOC_FIELDS.date, DOC_FIELDS.amount, DOC_FIELDS.counterparty, "ДоговорКонтрагента_Key"];
  const { rows, truncated } = await fetchAll(
    conn.client,
    set,
    { filter, select, orderby: `${DOC_FIELDS.date} asc` },
    ANALYTICS_PAGE,
    conn.behavior.analyticsMaxRows,
  );

  const buckets = new Map<string, Bucket>();
  const bump = (k: string, amt: number): void => {
    const b = buckets.get(k) ?? { key: k, count: 0, sum: 0 };
    b.count += 1;
    b.sum += amt;
    buckets.set(k, b);
  };
  let total = 0,
    docCount = 0;
  const cpKeys = new Set<string>();
  const contractKeys = new Set<string>();
  for (const r of rows) {
    const cp = String(r[DOC_FIELDS.counterparty] ?? "").toLowerCase();
    if (kindSet && (!cp || !kindSet.has(cp))) continue;
    const amount = num(r[DOC_FIELDS.amount]);
    total += amount;
    docCount += 1;
    let key: string;
    if (args.groupBy === "counterparty") {
      key = cp;
      if (key) cpKeys.add(key);
    } else if (args.groupBy === "month") {
      key = String(r[DOC_FIELDS.date] ?? "").slice(0, 7);
    } else if (args.groupBy === "contract") {
      key = String(r["ДоговорКонтрагента_Key"] ?? "").toLowerCase();
      if (key) contractKeys.add(key);
    } else key = "ИТОГО";
    bump(key, amount);
  }

  const names = new Map<string, string>();
  if (args.groupBy === "counterparty" && cpKeys.size) {
    const cpSet = resolveEntity(CATALOGS.counterparties, available);
    if (cpSet) for (const [k, v] of await resolveNames(conn, cpSet, cpKeys)) names.set(k.toLowerCase(), v);
  }
  if (args.groupBy === "contract" && contractKeys.size) {
    const cSet = resolveEntity(CATALOGS.contracts, available);
    if (cSet)
      for (const [k, v] of await resolveNames(conn, cSet, contractKeys)) names.set(k.toLowerCase(), v);
  }

  const round = (n: number): number => Math.round(n * 100) / 100;
  const labelOf = (b: Bucket): string => {
    if (args.groupBy === "counterparty") return b.key ? (names.get(b.key) ?? b.key) : "Без контрагента";
    if (args.groupBy === "contract") return b.key ? (names.get(b.key) ?? b.key) : "Без договора";
    return b.key || "—";
  };
  const groups = [...buckets.values()]
    .map((b) => ({ label: labelOf(b), count: b.count, sum: round(b.sum) }))
    .sort((a, b) => b.sum - a.sum)
    .slice(0, args.limit);

  return ok({
    database: conn.cfg.name,
    organization: org?.name,
    period: { from: args.from, to: args.to },
    filters: {
      ...(args.counterpartyRef ? { counterpartyRef: args.counterpartyRef } : {}),
      ...(args.counterpartyKind
        ? { counterpartyKind: args.counterpartyKind, counterpartyCount: kindSet?.size ?? 0 }
        : {}),
      ...(args.contractRef ? { contractRef: args.contractRef } : {}),
    },
    groupBy: args.groupBy,
    entitySet: set,
    documents: docCount,
    total: round(total),
    groups,
    ...(truncated
      ? {
          truncated: true,
          note: `Усечено лимитом ${conn.behavior.analyticsMaxRows} строк — сузьте период/фильтр.`,
        }
      : {}),
  });
}

export function registerSalesTools(server: McpServer, ctx: ServerContext): void {
  const commonInput = {
    database: databaseField,
    organization: organizationField,
    from: z.string().describe("Дата начала периода (YYYY-MM-DD)"),
    to: z.string().describe("Дата конца периода (YYYY-MM-DD)"),
    counterpartyRef: z.string().optional().describe("Ref_Key контрагента (точечный фильтр)"),
    counterpartyKind: z
      .enum(CP_KIND)
      .optional()
      .describe("Категория контрагента: ИП, ЮрЛицо, ФизЛицо (без ИП), Нерезидент, Госорган"),
    contractRef: z.string().optional().describe("Ref_Key договора (точечный фильтр)"),
    groupBy: z
      .enum(SALES_GROUP_BY)
      .default("counterparty")
      .describe(
        "Разбивка: counterparty (контрагент), month (месяц), contract (договор), total (без разбивки)",
      ),
    limit: z.number().int().positive().max(1000).default(50).describe("Сколько групп вернуть"),
  };

  server.registerTool(
    "get_sales_breakdown",
    {
      title: "Разбивка продаж за период",
      description:
        "Суммирует проведённые реализации (СуммаДокумента) за период с разбивкой по контрагенту / " +
        "месяцу / договору / итогу. Можно фильтровать конкретным контрагентом, категорией контрагента " +
        "(ИП/ЮрЛицо/ФизЛицо/Нерезидент/Госорган) или договором. Отвечает на вопросы вроде «топ покупателей " +
        "за год», «выручка по месяцам», «сколько продали ИП», «продажи по конкретному договору». " +
        "Период — даты YYYY-MM-DD.",
      inputSchema: commonInput,
    },
    (a) =>
      guard("get_sales_breakdown", async () =>
        aggregate(ctx.db(a.database), DOCUMENTS.sales, "Реализация товаров и услуг", a),
      ),
  );

  server.registerTool(
    "get_purchases_breakdown",
    {
      title: "Разбивка закупок за период",
      description:
        "Суммирует проведённые поступления товаров/услуг (закупки) за период с разбивкой по " +
        "поставщику / месяцу / договору / итогу. Можно фильтровать конкретным поставщиком, категорией " +
        "контрагента или договором. Отвечает на вопросы вроде «топ поставщиков за год», «закупки по " +
        "месяцам», «сколько закупили у ИП». Период — даты YYYY-MM-DD.",
      inputSchema: commonInput,
    },
    (a) =>
      guard("get_purchases_breakdown", async () =>
        aggregate(ctx.db(a.database), DOCUMENTS.purchases, "Поступление товаров и услуг", a),
      ),
  );
}
