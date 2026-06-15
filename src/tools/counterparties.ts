import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, withTruncationNote, databaseField, organizationField } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, contains, odataGuid, odataString, buildQuery } from "../odata/query.js";
import {
  CATALOGS,
  COUNTERPARTY_FIELDS as CF,
  DOC_FIELDS,
  DOCUMENTS,
  resolveEntity,
} from "../config/mapping.js";
import { resolveOrganization } from "../odata/orgs.js";
import { requireEntity } from "../odata/publication.js";
import type { Counterparty, DocumentSummary } from "../types/domain.js";
import type { ODataEntity } from "../types/odata.js";

function toCounterparty(r: ODataEntity): Counterparty {
  return {
    ref: String(r[CF.ref] ?? ""),
    name: String(r[CF.name] ?? ""),
    code: r[CF.code] ? String(r[CF.code]) : undefined,
    inn: r[CF.inn] ? String(r[CF.inn]) : undefined,
    kpp: r[CF.kpp] ? String(r[CF.kpp]) : undefined,
    fullName: r[CF.fullName] ? String(r[CF.fullName]) : undefined,
    isFolder: typeof r[CF.isFolder] === "boolean" ? (r[CF.isFolder] as boolean) : undefined,
  };
}

const counterpartySet = (conn: Connection): Promise<string> =>
  requireEntity(conn, CATALOGS.counterparties, "Справочник «Контрагенты»");

export function registerCounterpartyTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "find_counterparty",
    {
      title: "Поиск контрагента",
      description:
        "Ищет контрагентов по части названия или по ИНН. Возвращает список совпадений " +
        "с Ref_Key (используйте его в get_counterparty и истории взаиморасчётов). " +
        "Пример: «найди контрагента Ромашка» или «контрагент с ИНН 7701234567».",
      inputSchema: {
        database: databaseField,
        query: z.string().min(1).describe("Часть названия или ИНН контрагента"),
        limit: z.number().int().positive().max(100).default(20).describe("Сколько вернуть"),
      },
    },
    ({ database, query, limit }) =>
      guard("find_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await counterpartySet(conn);
        const isInn = /^\d{10,12}$/.test(query.trim());
        const filter = isInn
          ? cmp(CF.inn, "eq", odataString(query.trim()))
          : contains(CF.name, query);

        const { rows, truncated } = await fetchAll(
          conn.client,
          set,
          { filter, select: [CF.ref, CF.name, CF.code, CF.inn, CF.kpp, CF.isFolder] },
          conn.behavior.pageSize,
          Math.min(limit, conn.behavior.maxRows),
        );

        const items = rows.map(toCounterparty).filter((c) => !c.isFolder);
        return ok(withTruncationNote(items, truncated, limit));
      }),
  );

  server.registerTool(
    "get_counterparty",
    {
      title: "Карточка контрагента",
      description:
        "Возвращает полную карточку контрагента по его Ref_Key: реквизиты (ИНН/КПП), " +
        "наименования, код. Ref_Key берётся из find_counterparty.",
      inputSchema: {
        database: databaseField,
        ref: z.string().describe("Ref_Key контрагента (GUID)"),
      },
    },
    ({ database, ref }) =>
      guard("get_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await counterpartySet(conn);
        const path = `${set}(guid'${ref.replace(/[{}']/g, "")}')${buildQuery({})}`;
        const entity = await conn.client.getEntity(path);
        return ok(toCounterparty(entity));
      }),
  );

  const historyInput = {
    database: databaseField,
    organization: organizationField,
    ref: z.string().describe("Ref_Key контрагента"),
    from: z.string().optional().describe("Дата начала периода (YYYY-MM-DD)"),
    to: z.string().optional().describe("Дата конца периода (YYYY-MM-DD)"),
    limit: z.number().int().positive().max(500).default(100),
  };

  async function history(
    conn: Connection,
    docKeys: readonly string[],
    ref: string,
    orgKey: string | undefined,
    from: string | undefined,
    to: string | undefined,
    limit: number,
  ): Promise<{ docs: DocumentSummary[]; truncated: boolean; usedSets: string[] }> {
    const available = await conn.available();
    const guid = odataGuid(ref);
    const docs: DocumentSummary[] = [];
    const usedSets: string[] = [];
    let truncated = false;

    for (const key of docKeys) {
      const set = resolveEntity([key], available);
      if (!set) continue;
      usedSets.push(set);
      const filter = and(
        cmp(DOC_FIELDS.counterparty, "eq", guid),
        orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
        from ? cmp(DOC_FIELDS.date, "ge", `datetime'${from}T00:00:00'`) : undefined,
        to ? cmp(DOC_FIELDS.date, "le", `datetime'${to}T23:59:59'`) : undefined,
      );
      const { rows, truncated: t } = await fetchAll(
        conn.client,
        set,
        { filter, orderby: `${DOC_FIELDS.date} desc` },
        conn.behavior.pageSize,
        limit,
      );
      truncated ||= t;
      for (const r of rows) {
        docs.push({
          ref: String(r[DOC_FIELDS.ref] ?? ""),
          type: set,
          entitySet: set,
          number: r[DOC_FIELDS.number] ? String(r[DOC_FIELDS.number]) : undefined,
          date: r[DOC_FIELDS.date] ? String(r[DOC_FIELDS.date]) : undefined,
          posted: typeof r[DOC_FIELDS.posted] === "boolean" ? (r[DOC_FIELDS.posted] as boolean) : undefined,
          amount: typeof r[DOC_FIELDS.amount] === "number" ? (r[DOC_FIELDS.amount] as number) : undefined,
        });
      }
    }
    docs.sort((a, b) => (b.date ?? "").localeCompare(a.date ?? ""));
    return { docs: docs.slice(0, limit), truncated, usedSets };
  }

  async function orgKeyOf(conn: Connection, organization?: string): Promise<string | undefined> {
    return organization ? (await resolveOrganization(conn, organization)).ref : undefined;
  }

  server.registerTool(
    "get_customer_history",
    {
      title: "История по покупателю",
      description:
        "Документы реализации и счета по контрагенту-покупателю за период. " +
        "Показывает, что и на какие суммы продавали клиенту. Ref_Key — из find_counterparty.",
      inputSchema: historyInput,
    },
    ({ database, organization, ref, from, to, limit }) =>
      guard("get_customer_history", async () => {
        const conn = ctx.db(database);
        const orgKey = await orgKeyOf(conn, organization);
        const r = await history(conn, [...DOCUMENTS.sales, ...DOCUMENTS.customerInvoice], ref, orgKey, from, to, limit);
        if (r.usedSets.length === 0) {
          return fail("Документы продаж не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        return ok({ counterparty: ref, period: { from, to }, ...withTruncationNote(r.docs, r.truncated, limit) });
      }),
  );

  server.registerTool(
    "get_supplier_history",
    {
      title: "История по поставщику",
      description:
        "Документы поступления по контрагенту-поставщику за период. " +
        "Показывает закупки у поставщика. Ref_Key — из find_counterparty.",
      inputSchema: historyInput,
    },
    ({ database, organization, ref, from, to, limit }) =>
      guard("get_supplier_history", async () => {
        const conn = ctx.db(database);
        const orgKey = await orgKeyOf(conn, organization);
        const r = await history(conn, [...DOCUMENTS.purchases], ref, orgKey, from, to, limit);
        if (r.usedSets.length === 0) {
          return fail("Документы поступлений не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        return ok({ counterparty: ref, period: { from, to }, ...withTruncationNote(r.docs, r.truncated, limit) });
      }),
  );
}
