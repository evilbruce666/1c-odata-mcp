import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { ok, fail, guard } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, buildQuery } from "../odata/query.js";
import { DOC_FIELDS, DOCUMENTS, REGISTERS, resolveEntity } from "../config/mapping.js";
import type { ODataEntity } from "../types/odata.js";

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
    "get_inventory",
    {
      title: "Остатки на складах",
      description:
        "Текущие остатки товаров по складам из регистра накопления (виртуальная таблица Balance). " +
        "Возвращает строки с измерениями (Номенклатура/Склад) и количеством. " +
        "ВНИМАНИЕ: точное имя регистра подтверждается из $metadata при первом запуске.",
      inputSchema: {
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько строк вернуть"),
      },
    },
    ({ limit }) =>
      guard("get_inventory", async () => {
        const available = await ctx.available();
        const reg = resolveEntity(REGISTERS.stock, available);
        if (!reg) {
          return fail(
            "Регистр остатков товаров не опубликован в OData. " +
              "Добавьте AccumulationRegister_ТоварыНаСкладах в «Состав OData».",
          );
        }
        // Виртуальная таблица остатков: <Регистр>/Balance
        const path = `${reg}/Balance${buildQuery({ top: Math.min(limit, ctx.cfg.ODATA_MAX_ROWS) })}`;
        const page = await ctx.client.getCollection<ODataEntity>(path);
        return ok({ register: reg, count: page.value.length, rows: page.value });
      }),
  );

  server.registerTool(
    "get_debtors",
    {
      title: "Дебиторская задолженность",
      description:
        "Остатки дебиторской задолженности (сальдо счёта 62) из регистра бухгалтерии Хозрасчетный, " +
        "виртуальная таблица Balance. Показывает, кто и сколько должен компании. " +
        "ЭКСПЕРИМЕНТАЛЬНО: формат виртуальной таблицы регистра бухгалтерии подтверждается " +
        "из живой базы при первом запуске; при ошибке используйте get_customer_history по контрагенту.",
      inputSchema: {
        limit: z.number().int().positive().max(1000).default(200),
      },
    },
    ({ limit }) =>
      guard("get_debtors", async () => {
        const available = await ctx.available();
        const reg = resolveEntity(REGISTERS.accounting, available);
        if (!reg) {
          return fail(
            "Регистр бухгалтерии не опубликован в OData. " +
              "Добавьте AccountingRegister_Хозрасчетный в «Состав OData».",
          );
        }
        const path = `${reg}/Balance${buildQuery({ top: Math.min(limit, ctx.cfg.ODATA_MAX_ROWS) })}`;
        const page = await ctx.client.getCollection<ODataEntity>(path);
        return ok({
          register: reg,
          note: "Сырые строки сальдо. Фильтрацию по счёту 62 и группировку по контрагенту откалибруем после сверки с $metadata живой базы.",
          count: page.value.length,
          rows: page.value,
        });
      }),
  );
}
