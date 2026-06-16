import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField, dateField } from "./_shared.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { ACCOUNT_PREFIX, CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { balanceByAccounts, resolveAccounts, resolveNames, num } from "../odata/accounting.js";
import { resolveOrganization } from "../odata/orgs.js";
import { collectDocuments, emptyMeta, addMeta, type ScanMeta } from "../odata/aggregate.js";

// Деньги копим в целых копейках (float-сложение тысяч сумм даёт дрейф).
const toCents = (v: unknown): number => Math.round(num(v) * 100);
const fromCents = (c: number): number => Math.round(c) / 100;

/**
 * Суммирует СуммаДокумента по проведённым документам за период (+орг). Через
 * collectDocuments: полная выборка с авто-чанкингом и громким переполнением
 * (раньше стояло под общим maxRows=1000 → годовые итоги занижались).
 */
async function sumDocuments(
  conn: Connection,
  docKeys: readonly string[],
  from: string,
  to: string,
  orgKey: string | undefined,
): Promise<{ totalCents: number; perSet: Record<string, number>; usedSets: string[]; meta: ScanMeta }> {
  const available = await conn.available();
  const baseFilter =
    and(
      orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
      cmp(DOC_FIELDS.posted, "eq", "true"),
    ) || undefined;
  const perSet: Record<string, number> = {};
  const usedSets: string[] = [];
  let totalCents = 0;
  let meta = emptyMeta();

  for (const key of docKeys) {
    const set = resolveEntity([key], available);
    if (!set) continue;
    usedSets.push(set);
    const { rows, meta: m } = await collectDocuments(conn, set, {
      baseFilter,
      dateField: DOC_FIELDS.date,
      from,
      to,
      select: [DOC_FIELDS.date, DOC_FIELDS.amount],
    });
    meta = addMeta(meta, m);
    const c = rows.reduce((acc, r) => acc + toCents(r[DOC_FIELDS.amount]), 0);
    perSet[set] = fromCents(c);
    totalCents += c;
  }
  return { totalCents, perSet, usedSets, meta };
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
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
      },
    },
    ({ database, organization, from, to }) =>
      guard("get_sales", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
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
          total: fromCents(r.totalCents),
          byDocument: r.perSet,
          scan: { documentsScanned: r.meta.rowsScanned, windows: r.meta.chunks, elapsedMs: Date.now() - t0 },
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
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
      },
    },
    ({ database, organization, from, to }) =>
      guard("get_cashflow", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const [inflow, outflow] = await Promise.all([
          sumDocuments(conn, [...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn], from, to, org.key),
          sumDocuments(conn, [...DOCUMENTS.bankOut, ...DOCUMENTS.cashOut], from, to, org.key),
        ]);
        if (inflow.usedSets.length === 0 && outflow.usedSets.length === 0) {
          return fail("Банковские/кассовые документы не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        const scanned = inflow.meta.rowsScanned + outflow.meta.rowsScanned;
        const windows = inflow.meta.chunks + outflow.meta.chunks;
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          period: { from, to },
          inflow: fromCents(inflow.totalCents),
          outflow: fromCents(outflow.totalCents),
          net: fromCents(inflow.totalCents - outflow.totalCents),
          byDocument: { ...inflow.perSet, ...outflow.perSet },
          scan: { documentsScanned: scanned, windows, elapsedMs: Date.now() - t0 },
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
        "кредитовое = полученные авансы (вычитается). Можно ограничить организацией. По умолчанию " +
        "берётся текущее сальдо; параметром asOf=YYYY-MM-DD можно получить дебиторку на конец " +
        "указанной даты (для аудита/исторических отчётов). Возвращает только долг > 0.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        asOf: dateField("Дата сальдо — на конец этой даты (без параметра — текущее)").optional(),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько контрагентов вернуть"),
      },
    },
    ({ database, organization, asOf, limit }) =>
      guard("get_debtors", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, ACCOUNT_PREFIX.receivables);
        const rows = await balanceByAccounts(
          conn,
          accounts.map((a) => a.key),
          org.key,
          asOf,
        );

        // Сальдо копим в копейках (целые) — без float-дрейфа на тысячах строк.
        const byCp = new Map<string, number>();
        for (const r of rows) {
          const cp = String(r["ExtDimension1"] ?? "");
          if (!cp) continue;
          byCp.set(cp, (byCp.get(cp) ?? 0) + toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]));
        }

        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const names = cpSet ? await resolveNames(conn, cpSet, byCp.keys()) : new Map<string, string>();

        const debtors = [...byCp.entries()]
          .filter(([, cents]) => cents > 0)
          .map(([ref, cents]) => ({
            counterparty: names.get(ref) ?? ref,
            ref,
            amount: fromCents(cents),
          }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const totalCents = [...byCp.values()].filter((c) => c > 0).reduce((s, c) => s + c, 0);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          ...(asOf ? { asOf } : {}),
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalReceivable: fromCents(totalCents),
          count: debtors.length,
          debtors,
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  server.registerTool(
    "get_inventory",
    {
      title: "Остатки товаров",
      description:
        "Остатки товаров/материалов на складах: сальдо счетов 41/10/43 регистра Хозрасчетный, " +
        "сгруппированное по номенклатуре. Возвращает количество и сумму остатка. Можно ограничить " +
        "организацией. По умолчанию — текущие остатки; параметром asOf=YYYY-MM-DD можно получить " +
        "остатки на конец указанной даты (для инвентаризации/аудита). " +
        "(В БП 3.0 учёт ведётся на счетах бухучёта, а не в отдельном регистре остатков.)",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        asOf: dateField("Дата остатков — на конец этой даты (без параметра — текущие)").optional(),
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько позиций вернуть"),
      },
    },
    ({ database, organization, asOf, limit }) =>
      guard("get_inventory", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, ACCOUNT_PREFIX.inventory);
        const rows = await balanceByAccounts(
          conn,
          accounts.map((a) => a.key),
          org.key,
          asOf,
        );

        // Сумма — в копейках (целые), количество — float (округляем до 3 знаков).
        const byItem = new Map<string, { qty: number; cents: number }>();
        for (const r of rows) {
          const item = String(r["ExtDimension1"] ?? "");
          if (!item) continue;
          const cur = byItem.get(item) ?? { qty: 0, cents: 0 };
          cur.qty += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
          cur.cents += toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]);
          byItem.set(item, cur);
        }

        const nomSet = resolveEntity(CATALOGS.nomenclature, await conn.available());
        const names = nomSet ? await resolveNames(conn, nomSet, byItem.keys()) : new Map<string, string>();

        const present = [...byItem.entries()].filter(
          ([, v]) => Math.abs(v.qty) > 0.0001 || Math.abs(v.cents) > 0,
        );
        const items = present
          .map(([ref, v]) => ({
            item: names.get(ref) ?? ref,
            ref,
            quantity: Math.round(v.qty * 1000) / 1000,
            amount: fromCents(v.cents),
          }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const totalCents = present.reduce((s, [, v]) => s + v.cents, 0);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          ...(asOf ? { asOf } : {}),
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalAmount: fromCents(totalCents),
          count: items.length,
          items,
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
