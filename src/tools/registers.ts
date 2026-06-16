import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { ACCOUNT_PREFIX, CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { balanceByAccounts, resolveAccounts, resolveNames, num } from "../odata/accounting.js";
import { resolveOrganization } from "../odata/orgs.js";

/** Суммирует поле СуммаДокумента по проведённым документам за период (+орг). */
async function sumDocuments(
  conn: Connection,
  docKeys: readonly string[],
  from: string | undefined,
  to: string | undefined,
  orgKey: string | undefined,
): Promise<{ total: number; perSet: Record<string, number>; usedSets: string[] }> {
  const available = await conn.available();
  const filter = and(
    from ? cmp(DOC_FIELDS.date, "ge", `datetime'${from}T00:00:00'`) : undefined,
    to ? cmp(DOC_FIELDS.date, "le", `datetime'${to}T23:59:59'`) : undefined,
    orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
    cmp(DOC_FIELDS.posted, "eq", "true"),
  );
  const perSet: Record<string, number> = {};
  const usedSets: string[] = [];
  let total = 0;

  for (const key of docKeys) {
    const set = resolveEntity([key], available);
    if (!set) continue;
    usedSets.push(set);
    const { rows } = await fetchAll(
      conn.client,
      set,
      { filter, select: [DOC_FIELDS.date, DOC_FIELDS.amount, DOC_FIELDS.posted] },
      conn.behavior.pageSize,
      conn.behavior.maxRows,
    );
    const s = rows.reduce((acc, r) => acc + num(r[DOC_FIELDS.amount]), 0);
    perSet[set] = s;
    total += s;
  }
  return { total, perSet, usedSets };
}

async function orgKeyOf(conn: Connection, organization?: string): Promise<{ key?: string; name?: string }> {
  if (!organization) return {};
  const org = await resolveOrganization(conn, organization);
  return { key: org.ref, name: org.name };
}

export function registerRegisterTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "get_sales",
    {
      title: "Продажи за период",
      description:
        "Сумма выручки по проведённым документам реализации за период (по полю СуммаДокумента). " +
        "Можно ограничить организацией. Период задаётся датами YYYY-MM-DD.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        from: z.string().describe("Дата начала периода (YYYY-MM-DD)"),
        to: z.string().describe("Дата конца периода (YYYY-MM-DD)"),
      },
    },
    ({ database, organization, from, to }) =>
      guard("get_sales", async () => {
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const r = await sumDocuments(conn, [...DOCUMENTS.sales], from, to, org.key);
        if (r.usedSets.length === 0) {
          return fail("Документы реализации не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          period: { from, to },
          total: r.total,
          byDocument: r.perSet,
        });
      }),
  );

  server.registerTool(
    "get_cashflow",
    {
      title: "Движение денежных средств",
      description:
        "Приход и расход денег за период по проведённым банковским и кассовым документам. " +
        "Возвращает приток, отток и сальдо. Можно ограничить организацией. Период — даты YYYY-MM-DD.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        from: z.string().describe("Дата начала периода (YYYY-MM-DD)"),
        to: z.string().describe("Дата конца периода (YYYY-MM-DD)"),
      },
    },
    ({ database, organization, from, to }) =>
      guard("get_cashflow", async () => {
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const inflow = await sumDocuments(
          conn,
          [...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn],
          from,
          to,
          org.key,
        );
        const outflow = await sumDocuments(
          conn,
          [...DOCUMENTS.bankOut, ...DOCUMENTS.cashOut],
          from,
          to,
          org.key,
        );
        if (inflow.usedSets.length === 0 && outflow.usedSets.length === 0) {
          return fail("Банковские/кассовые документы не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          period: { from, to },
          inflow: inflow.total,
          outflow: outflow.total,
          net: inflow.total - outflow.total,
          byDocument: { ...inflow.perSet, ...outflow.perSet },
        });
      }),
  );

  server.registerTool(
    "get_debtors",
    {
      title: "Дебиторская задолженность",
      description:
        "Кто и сколько должен компании: сальдо счёта 62 (расчёты с покупателями) из регистра " +
        "бухгалтерии Хозрасчетный, сгруппированное по контрагентам. Дебетовое сальдо = долг клиента, " +
        "кредитовое = полученные авансы (вычитается). Можно ограничить организацией. Возвращает только долг > 0.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько контрагентов вернуть"),
      },
    },
    ({ database, organization, limit }) =>
      guard("get_debtors", async () => {
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, ACCOUNT_PREFIX.receivables);
        const rows = await balanceByAccounts(
          conn,
          accounts.map((a) => a.key),
          conn.behavior.maxRows,
          org.key,
        );

        const byCp = new Map<string, number>();
        for (const r of rows) {
          const cp = String(r["ExtDimension1"] ?? "");
          if (!cp) continue;
          byCp.set(cp, (byCp.get(cp) ?? 0) + num(r["СуммаBalanceDr"]) - num(r["СуммаBalanceCr"]));
        }

        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const names = cpSet ? await resolveNames(conn, cpSet, byCp.keys()) : new Map<string, string>();

        const debtors = [...byCp.entries()]
          .filter(([, amount]) => amount > 0.005)
          .map(([ref, amount]) => ({
            counterparty: names.get(ref) ?? ref,
            ref,
            amount: Math.round(amount * 100) / 100,
          }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const total = debtors.reduce((s, d) => s + d.amount, 0);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalReceivable: Math.round(total * 100) / 100,
          count: debtors.length,
          debtors,
        });
      }),
  );

  server.registerTool(
    "get_inventory",
    {
      title: "Остатки товаров",
      description:
        "Текущие остатки товаров/материалов на складах: сальдо счетов 41/10/43 регистра Хозрасчетный, " +
        "сгруппированное по номенклатуре. Возвращает количество и сумму остатка. Можно ограничить организацией. " +
        "(В БП 3.0 учёт ведётся на счетах бухучёта, а не в отдельном регистре остатков.)",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько позиций вернуть"),
      },
    },
    ({ database, organization, limit }) =>
      guard("get_inventory", async () => {
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, ACCOUNT_PREFIX.inventory);
        const rows = await balanceByAccounts(
          conn,
          accounts.map((a) => a.key),
          conn.behavior.maxRows,
          org.key,
        );

        const byItem = new Map<string, { qty: number; amount: number }>();
        for (const r of rows) {
          const item = String(r["ExtDimension1"] ?? "");
          if (!item) continue;
          const cur = byItem.get(item) ?? { qty: 0, amount: 0 };
          cur.qty += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
          cur.amount += num(r["СуммаBalanceDr"]) - num(r["СуммаBalanceCr"]);
          byItem.set(item, cur);
        }

        const nomSet = resolveEntity(CATALOGS.nomenclature, await conn.available());
        const names = nomSet ? await resolveNames(conn, nomSet, byItem.keys()) : new Map<string, string>();

        const items = [...byItem.entries()]
          .filter(([, v]) => Math.abs(v.qty) > 0.0001 || Math.abs(v.amount) > 0.005)
          .map(([ref, v]) => ({
            item: names.get(ref) ?? ref,
            ref,
            quantity: Math.round(v.qty * 1000) / 1000,
            amount: Math.round(v.amount * 100) / 100,
          }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const totalAmount = items.reduce((s, i) => s + i.amount, 0);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalAmount: Math.round(totalAmount * 100) / 100,
          count: items.length,
          items,
        });
      }),
  );
}
