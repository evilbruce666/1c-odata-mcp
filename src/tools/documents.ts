import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { ok, fail, guard, withTruncationNote } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, odataGuid, buildQuery } from "../odata/query.js";
import { DOC_FIELDS } from "../config/mapping.js";
import type { DocumentSummary } from "../types/domain.js";
import type { ODataEntity } from "../types/odata.js";

function toSummary(r: ODataEntity, set: string): DocumentSummary {
  return {
    ref: String(r[DOC_FIELDS.ref] ?? ""),
    type: set,
    entitySet: set,
    number: r[DOC_FIELDS.number] ? String(r[DOC_FIELDS.number]) : undefined,
    date: r[DOC_FIELDS.date] ? String(r[DOC_FIELDS.date]) : undefined,
    posted: typeof r[DOC_FIELDS.posted] === "boolean" ? (r[DOC_FIELDS.posted] as boolean) : undefined,
    deletionMark:
      typeof r[DOC_FIELDS.deletionMark] === "boolean" ? (r[DOC_FIELDS.deletionMark] as boolean) : undefined,
    amount: typeof r[DOC_FIELDS.amount] === "number" ? (r[DOC_FIELDS.amount] as number) : undefined,
  };
}

export function registerDocumentTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "search_documents",
    {
      title: "Поиск документов",
      description:
        "Ищет документы заданного типа по периоду, контрагенту, статусу проведения и сумме. " +
        "Тип документа — техническое имя из list_entities, напр. 'Document_РеализацияТоваровУслуг'. " +
        "Сначала уточните имя через list_entities (class=document), а поля — через describe_entity.",
      inputSchema: {
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        from: z.string().optional().describe("Дата начала (YYYY-MM-DD)"),
        to: z.string().optional().describe("Дата конца (YYYY-MM-DD)"),
        counterpartyRef: z.string().optional().describe("Ref_Key контрагента для фильтра"),
        postedOnly: z.boolean().default(false).describe("Только проведённые документы"),
        minAmount: z.number().optional().describe("Минимальная сумма документа"),
        limit: z.number().int().positive().max(500).default(50),
      },
    },
    ({ entitySet, from, to, counterpartyRef, postedOnly, minAmount, limit }) =>
      guard("search_documents", async () => {
        const available = await ctx.available();
        if (!available.has(entitySet)) {
          return fail(`Документ '${entitySet}' не найден/не опубликован. Проверьте через list_entities.`);
        }
        const filter = and(
          from ? cmp(DOC_FIELDS.date, "ge", `datetime'${from}T00:00:00'`) : undefined,
          to ? cmp(DOC_FIELDS.date, "le", `datetime'${to}T23:59:59'`) : undefined,
          counterpartyRef ? cmp(DOC_FIELDS.counterparty, "eq", odataGuid(counterpartyRef)) : undefined,
          postedOnly ? cmp(DOC_FIELDS.posted, "eq", "true") : undefined,
          typeof minAmount === "number" ? cmp(DOC_FIELDS.amount, "ge", String(minAmount)) : undefined,
        );

        const { rows, truncated } = await fetchAll(
          ctx.client,
          entitySet,
          { filter: filter || undefined, orderby: `${DOC_FIELDS.date} desc` },
          ctx.cfg.ODATA_PAGE_SIZE,
          limit,
        );
        const items = rows.map((r) => toSummary(r, entitySet));
        return ok(withTruncationNote(items, truncated, limit));
      }),
  );

  server.registerTool(
    "get_document",
    {
      title: "Документ целиком",
      description:
        "Возвращает документ по типу и Ref_Key со всеми полями, включая табличные части. " +
        "Ref_Key берётся из search_documents или истории взаиморасчётов.",
      inputSchema: {
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
      },
    },
    ({ entitySet, ref }) =>
      guard("get_document", async () => {
        const available = await ctx.available();
        if (!available.has(entitySet)) {
          return fail(`Документ '${entitySet}' не найден/не опубликован.`);
        }
        const path = `${entitySet}(guid'${ref.replace(/[{}']/g, "")}')${buildQuery({})}`;
        const entity = await ctx.client.getEntity(path);
        return ok(entity);
      }),
  );
}
