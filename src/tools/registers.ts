import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { ok, fail, guard } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp } from "../odata/query.js";
import { ACCOUNT_PREFIX, CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import {
  balanceByAccounts,
  resolveAccounts,
  resolveNames,
  num,
} from "../odata/accounting.js";

/** Суммирует числовое поле по строкам документов за период. */
async function sumDocuments(
  ctx: ServerContext,
  docKeys: readonly string[],
  from: string | undefined,
  to: string | undefined,
): Promise<{ total: number; perSet: Record<string, number>; usedSets: string[] }> {
  const available = await ctx.available();
  const filter = and(
    from ? cmp(DOC_FIELDS.date, "ge", `datetime'${from}T00:00:00'`) : undefined,
    to ? cmp(DOC_FIELDS.date, "le", `datetime'${to}T23:59:59'`) : undefined,
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
      ctx.client,
      set,
      { filter, select: [DOC_FIELDS.date, DOC_FIELDS.amount, DOC_FIELDS.posted] },
      ctx.cfg.ODATA_PAGE_SIZE,
      ctx.cfg.ODATA_MAX_ROWS,
    );
    const s = rows.reduce(
      (acc, r) => acc + (typeof r[DOC_FIELDS.amount] === "number" ? (r[DOC_FIELDS.amount] as number) : 0),
      0,
    );
    perSet[set] = s;
    total += s;
  }
  return { total, perSet, usedSets };
}

export function registerRegisterTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "get_sales",
    {
      title: "Продажи за период",
      description:
        "Сумма выручки по проведённым документам реализации за период (по полю СуммаДокумента). " +
        "Надёжный источник — документы продаж. Период задаётся датами YYYY-MM-DD.",
      inputSchema: {
        from: z.string().describe("Дата начала периода (YYYY-MM-DD)"),
        to: z.string().describe("Дата конца периода (YYYY-MM-DD)"),
      },
    },
    ({ from, to }) =>
      guard("get_sales", async () => {
        const r = await sumDocuments(ctx, [...DOCUMENTS.sales], from, to);
        if (r.usedSets.length === 0) {
          return fail("Документы реализации не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        return ok({ period: { from, to }, total: r.total, byDocument: r.perSet, source: r.usedSets });
      }),
  );

  server.registerTool(
    "get_cashflow",
    {
      title: "Движение денежных средств",
      description:
        "Приход и расход денег за период по проведённым банковским и кассовым документам. " +
        "Возвращает приток, отток и сальдо. Период — даты YYYY-MM-DD.",
      inputSchema: {
        from: z.string().describe("Дата начала периода (YYYY-MM-DD)"),
        to: z.string().describe("Дата конца периода (YYYY-MM-DD)"),
      },
    },
    ({ from, to }) =>
      guard("get_cashflow", async () => {
        const inflow = await sumDocuments(ctx, [...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn], from, to);
        const outflow = await sumDocuments(ctx, [...DOCUMENTS.bankOut, ...DOCUMENTS.cashOut], from, to);
        if (inflow.usedSets.length === 0 && outflow.usedSets.length === 0) {
          return fail(
            "Банковские/кассовые документы не опубликованы в OData. Добавьте их в «Состав OData».",
          );
        }
        return ok({
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
        "кредитовое = полученные авансы (вычитается). Возвращает только тех, у кого долг > 0.",
      inputSchema: {
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько контрагентов вернуть"),
      },
    },
    ({ limit }) =>
      guard("get_debtors", async () => {
        const accounts = await resolveAccounts(ctx, ACCOUNT_PREFIX.receivables);
        const rows = await balanceByAccounts(ctx, accounts.map((a) => a.key), ctx.cfg.ODATA_MAX_ROWS);

        // Группировка по контрагенту (ExtDimension1). Долг = Дт − Кт.
        const byCp = new Map<string, number>();
        for (const r of rows) {
          const cp = String(r["ExtDimension1"] ?? "");
          if (!cp) continue;
          const net = num(r["СуммаBalanceDr"]) - num(r["СуммаBalanceCr"]);
          byCp.set(cp, (byCp.get(cp) ?? 0) + net);
        }

        const cpSet = resolveEntity(CATALOGS.counterparties, await ctx.available());
        const names = cpSet ? await resolveNames(ctx, cpSet, byCp.keys()) : new Map<string, string>();

        const debtors = [...byCp.entries()]
          .filter(([, amount]) => amount > 0.005)
          .map(([ref, amount]) => ({ counterparty: names.get(ref) ?? ref, ref, amount: Math.round(amount * 100) / 100 }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const total = debtors.reduce((s, d) => s + d.amount, 0);
        return ok({
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
        "сгруппированное по номенклатуре. Возвращает количество и сумму остатка. " +
        "(В БП 3.0 учёт ведётся на счетах бухучёта, а не в отдельном регистре остатков.)",
      inputSchema: {
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько позиций вернуть"),
      },
    },
    ({ limit }) =>
      guard("get_inventory", async () => {
        const accounts = await resolveAccounts(ctx, ACCOUNT_PREFIX.inventory);
        const rows = await balanceByAccounts(ctx, accounts.map((a) => a.key), ctx.cfg.ODATA_MAX_ROWS);

        // Группировка по номенклатуре (ExtDimension1): количество и сумма.
        const byItem = new Map<string, { qty: number; amount: number }>();
        for (const r of rows) {
          const item = String(r["ExtDimension1"] ?? "");
          if (!item) continue;
          const cur = byItem.get(item) ?? { qty: 0, amount: 0 };
          cur.qty += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
          cur.amount += num(r["СуммаBalanceDr"]) - num(r["СуммаBalanceCr"]);
          byItem.set(item, cur);
        }

        const nomSet = resolveEntity(CATALOGS.nomenclature, await ctx.available());
        const names = nomSet ? await resolveNames(ctx, nomSet, byItem.keys()) : new Map<string, string>();

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
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalAmount: Math.round(totalAmount * 100) / 100,
          count: items.length,
          items,
        });
      }),
  );
}
