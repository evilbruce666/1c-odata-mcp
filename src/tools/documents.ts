import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, guard, withTruncationNote, databaseField, organizationField } from "./_shared.js";
import { fetchAll } from "../odata/pagination.js";
import { and, cmp, contains, odataGuid, odataString, or, buildQuery } from "../odata/query.js";
import { DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import { resolveOrganization } from "../odata/orgs.js";
import { ensurePublished } from "../odata/publication.js";
import {
  matchesIncomingNumber,
  matchesOwnNumber,
  parseNumberQuery,
  type NumberQuery,
} from "../odata/doc-number.js";
import type { DocumentSummary } from "../types/domain.js";
import type { ODataEntity } from "../types/odata.js";
import { searchDocumentsResultSchema, odataEntitySchema } from "../schemas/output.js";

/** Входящий номер — номер документа контрагента (накладной поставщика, платёжки покупателя). */
const INCOMING_NUMBER = "НомерВходящегоДокумента";
const POLY_COUNTERPARTY = "Контрагент";

/**
 * Виды документов простыми словами — чтобы не угадывать технические имена. Без вида и без
 * entitySet поиск идёт по всем видам отсюда (основные документы БП 3.0, а не все 160+ типов базы).
 */
const KINDS = {
  sales: ["sales", "servicesAct"],
  customer_invoices: ["customerInvoice"],
  purchases: ["purchases"],
  supplier_invoices: ["supplierInvoice"],
  bank: ["bankIn", "bankOut", "paymentOrder"],
  cash: ["cashIn", "cashOut"],
  vat_invoices: ["issuedInvoice", "receivedInvoice"],
  warehouse: ["returnFromCustomer", "returnToSupplier", "transfer", "surplus", "writeoff", "inventory"],
} as const satisfies Record<string, ReadonlyArray<keyof typeof DOCUMENTS>>;
type Kind = keyof typeof KINDS;
const KIND_NAMES = Object.keys(KINDS) as [Kind, ...Kind[]];

function toSummary(r: ODataEntity, set: string): DocumentSummary {
  const incoming = typeof r[INCOMING_NUMBER] === "string" ? r[INCOMING_NUMBER].trim() : "";
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
    ...(incoming ? { incomingNumber: incoming } : {}),
  };
}

interface SearchParams {
  from?: string | undefined;
  to?: string | undefined;
  orgKey?: string | undefined;
  counterpartyRef?: string | undefined;
  postedOnly: boolean;
  minAmount?: number | undefined;
  number?: NumberQuery | undefined;
  limit: number;
  /** Явно заданный тип без поля контрагента: искать без фильтра с пометкой (как было), а не пропускать. */
  ignoreMissingCounterparty?: boolean;
}

/** Итог по одному типу: строки или причина, по которой фильтр к типу неприменим. */
type SetResult =
  | { set: string; rows: DocumentSummary[]; truncated: boolean; counterpartyNote?: string }
  | { set: string; skipped: string };

async function searchSet(conn: Connection, set: string, p: SearchParams): Promise<SetResult> {
  const props = new Set(((await conn.getMetadata()).entities.get(set)?.properties ?? []).map((x) => x.name));
  // У перемещения, инвентаризации и др. нет СуммаДокумента — 1С на такой фильтр отвечает 400 (#45).
  // Молча пропустить фильтр нельзя: вернулись бы документы на любую сумму.
  if (typeof p.minAmount === "number" && !props.has(DOC_FIELDS.amount))
    return { set, skipped: `нет поля «${DOC_FIELDS.amount}» — фильтр minAmount неприменим` };
  if (p.orgKey && !props.has(DOC_FIELDS.organization))
    return { set, skipped: "нет поля организации — фильтр по организации неприменим" };

  // Контрагент бывает типизированным (Контрагент_Key — фильтруем на сервере)
  // или полиморфным (Контрагент: Edm.String — в 1С серверный eq/substringof по
  // нему не работает, поэтому фильтруем на стороне коннектора).
  let cpServer: string | undefined;
  let cpClientRef: string | undefined;
  let counterpartyNote: string | undefined;
  if (p.counterpartyRef) {
    if (props.has(DOC_FIELDS.counterparty)) {
      cpServer = cmp(DOC_FIELDS.counterparty, "eq", odataGuid(p.counterpartyRef));
    } else if (props.has(POLY_COUNTERPARTY)) {
      cpClientRef = p.counterpartyRef.replace(/[{}]/g, "").trim().toLowerCase();
      counterpartyNote =
        "Контрагент в этом документе хранится полиморфно — фильтр применён на стороне коннектора.";
    } else if (p.ignoreMissingCounterparty) {
      counterpartyNote = "В этом типе документа нет поля контрагента — фильтр по контрагенту проигнорирован.";
    } else {
      return { set, skipped: "нет поля контрагента — фильтр по контрагенту неприменим" };
    }
  }

  // Номер: на сервере — грубый отбор подстрокой (1С сравнивает без регистра), точно — здесь.
  const hasIncoming = props.has(INCOMING_NUMBER);
  const q = p.number;
  const numberFilter = q
    ? or(
        q.digits !== undefined
          ? `endswith(${DOC_FIELDS.number}, ${odataString(q.digits)})`
          : contains(DOC_FIELDS.number, q.needle),
        hasIncoming ? contains(INCOMING_NUMBER, q.needle) : undefined,
      )
    : undefined;

  const filter = and(
    p.from ? cmp(DOC_FIELDS.date, "ge", `datetime'${p.from}T00:00:00'`) : undefined,
    p.to ? cmp(DOC_FIELDS.date, "le", `datetime'${p.to}T23:59:59'`) : undefined,
    p.orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(p.orgKey)) : undefined,
    cpServer,
    p.postedOnly ? cmp(DOC_FIELDS.posted, "eq", "true") : undefined,
    typeof p.minAmount === "number" ? cmp(DOC_FIELDS.amount, "ge", String(p.minAmount)) : undefined,
    numberFilter,
  );
  // Только поля сводки: без табличных частей ответ 1С в разы меньше.
  const select = [
    DOC_FIELDS.ref,
    DOC_FIELDS.number,
    DOC_FIELDS.date,
    DOC_FIELDS.posted,
    DOC_FIELDS.deletionMark,
    DOC_FIELDS.amount,
    ...(cpClientRef ? [POLY_COUNTERPARTY] : []),
    INCOMING_NUMBER,
  ].filter((f) => props.has(f));

  // При клиентской фильтрации тянем больше строк (до maxRows), затем режем до limit.
  const clientSide = !!cpClientRef || !!q;
  const { rows, truncated } = await fetchAll(
    conn.client,
    set,
    { select, filter: filter || undefined, orderby: `${DOC_FIELDS.date} desc` },
    conn.behavior.pageSize,
    clientSide ? conn.behavior.maxRows : p.limit,
  );
  const matched = rows.filter(
    (r) =>
      (!cpClientRef ||
        String(r[POLY_COUNTERPARTY] ?? "")
          .replace(/[{}]/g, "")
          .trim()
          .toLowerCase() === cpClientRef) &&
      (!q || matchesOwnNumber(r[DOC_FIELDS.number], q) || matchesIncomingNumber(r[INCOMING_NUMBER], q)),
  );
  const summaries = matched.map((r) => {
    const s = toSummary(r, set);
    if (!q) return s;
    const by = matchesOwnNumber(r[DOC_FIELDS.number], q) ? "number" : "incomingNumber";
    return { ...s, matchedBy: by };
  });
  return {
    set,
    rows: summaries,
    truncated: truncated || matched.length > p.limit,
    ...(counterpartyNote ? { counterpartyNote } : {}),
  };
}

export function registerDocumentTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.document.search_documents",
    {
      title: "Поиск документов",
      description:
        "Ищет документы по периоду, номеру, контрагенту, организации, статусу проведения и сумме. " +
        "Тип задаётся видом (kind: продажи, покупки, банк, касса…) или техническим именем entitySet из " +
        "list_entities; без обоих — по всем основным видам документов сразу («все документы за 10 сентября», " +
        "«документ №33»). number — номер как его называет человек: «33» находит 0000-000033, а также " +
        "входящий номер документа контрагента (накладной поставщика). Нумерация в 1С каждый год своя — " +
        "по одному номеру обычно несколько документов разных лет; уточняйте период. Редкие виды " +
        "(зарплата, регламентные операции) — только через entitySet.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        entitySet: z
          .string()
          .optional()
          .describe("Техническое имя документа, напр. Document_РеализацияТоваровУслуг; не вместе с kind"),
        kind: z
          .enum(KIND_NAMES)
          .optional()
          .describe(
            "Вид документа: sales — реализации и акты, customer_invoices — счета покупателям, purchases — " +
              "поступления, supplier_invoices — счета поставщиков, bank — банк (поступления, списания, " +
              "платёжки), cash — касса (ПКО, РКО), vat_invoices — счета-фактуры, warehouse — складские",
          ),
        number: z.string().optional().describe("Номер документа: «33», «№ 0000-000033», входящий «ГК-22537»"),
        from: z.string().optional().describe("Дата начала (YYYY-MM-DD)"),
        to: z.string().optional().describe("Дата конца (YYYY-MM-DD)"),
        counterpartyRef: z.string().optional().describe("Ref_Key контрагента для фильтра"),
        postedOnly: z.boolean().default(false).describe("Только проведённые документы"),
        minAmount: z.number().optional().describe("Минимальная сумма документа"),
        limit: z.number().int().positive().max(500).default(50).describe("Сколько документов вернуть"),
      }),
      outputSchema: searchDocumentsResultSchema,
    },
    ({
      database,
      organization,
      entitySet,
      kind,
      number,
      from,
      to,
      counterpartyRef,
      postedOnly,
      minAmount,
      limit,
    }) =>
      guard("read.document.search_documents", async () => {
        if (entitySet && kind) throw new InputError("Задайте либо entitySet, либо kind, не оба.");
        const conn = ctx.db(database);
        const available = await conn.available();
        if (entitySet) ensurePublished(available, entitySet);
        const q = number === undefined ? undefined : parseNumberQuery(number);
        if (number !== undefined && !q) throw new InputError("Пустой номер документа.");
        const orgKey = organization ? (await resolveOrganization(conn, organization)).ref : undefined;
        const params: SearchParams = {
          from,
          to,
          orgKey,
          counterpartyRef,
          postedOnly,
          minAmount,
          number: q,
          limit,
        };

        if (entitySet) {
          const r = await searchSet(conn, entitySet, { ...params, ignoreMissingCounterparty: true });
          // Явно заданный тип: неприменимый фильтр — отказ, а не молчаливый пропуск.
          if ("skipped" in r)
            throw new InputError(
              `В документе ${entitySet} ${r.skipped}. Уберите фильтр; поля документа — в read.schema.describe_entity.`,
            );
          return ok({
            ...withTruncationNote(r.rows.sort(byNumberMatchThenDate).slice(0, limit), r.truncated, limit),
            ...(r.counterpartyNote ? { counterpartyFilter: r.counterpartyNote } : {}),
          });
        }

        const keys = kind ? KINDS[kind] : (Object.values(KINDS).flat() as Array<keyof typeof DOCUMENTS>);
        const sets = [...new Set(keys.map((k) => resolveEntity(DOCUMENTS[k], available)))].filter(
          (s): s is string => !!s,
        );
        if (sets.length === 0)
          throw new InputError(
            kind
              ? `Документы вида ${kind} в этой базе не опубликованы в OData.`
              : "Основные документы в этой базе не опубликованы в OData.",
          );
        // Виды — по одному: облачная 1С (1cfresh) одновременные запросы не ускоряет, а тормозит —
        // замер на живой базе: 8 запросов подряд 7,5 с, одновременно 32 с.
        const results: Array<SetResult | { set: string; error: string }> = [];
        for (const set of sets) {
          try {
            results.push(await searchSet(conn, set, params));
          } catch (e) {
            results.push({ set, error: e instanceof Error ? e.message : String(e) });
          }
        }
        const found = results.filter((r): r is Extract<SetResult, { rows: unknown }> => "rows" in r);
        const failed = results.filter((r): r is { set: string; error: string } => "error" in r);
        if (found.length === 0 && failed.length > 0)
          throw new InputError(`Поиск не удался ни по одному виду документов: ${failed[0]!.error}`);
        const merged = found.flatMap((r) => r.rows).sort(byNumberMatchThenDate);
        const truncated = found.some((r) => r.truncated) || merged.length > limit;
        return ok({
          ...withTruncationNote(merged.slice(0, limit), truncated, limit),
          searched: found.map((r) => ({ entitySet: r.set, found: r.rows.length })),
          ...(results.some((r) => "skipped" in r)
            ? {
                skipped: results
                  .filter((r): r is { set: string; skipped: string } => "skipped" in r)
                  .map((r) => ({ entitySet: r.set, reason: r.skipped })),
              }
            : {}),
          ...(failed.length ? { failed: failed.map((r) => ({ entitySet: r.set, error: r.error })) } : {}),
          scope:
            "Просмотрены основные виды документов (searched). Прочие типы — зарплата, регламентные " +
            "операции, уведомления — ищутся через entitySet.",
        });
      }),
  );

  server.registerTool(
    "read.document.get_document",
    {
      title: "Документ целиком",
      description:
        "Возвращает документ по типу и Ref_Key со всеми полями, включая табличные части. " +
        "Ref_Key берётся из search_documents или истории взаиморасчётов.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
      }),
      outputSchema: odataEntitySchema,
    },
    ({ database, entitySet, ref }) =>
      guard("read.document.get_document", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const path = `${entitySet}(guid'${ref.replace(/[{}']/g, "")}')${buildQuery({})}`;
        const entity = await conn.client.getEntity(path);
        return ok(entity);
      }),
  );
}

/**
 * Совпавшие по своему номеру — выше совпавших по входящему: «счёт №33» — это обычно свой номер, а
 * входящий «33» бывает у десятков банковских выписок. Внутри — новые сверху.
 */
function byNumberMatchThenDate(a: DocumentSummary, b: DocumentSummary): number {
  const rank = (d: DocumentSummary) => (d["matchedBy"] === "incomingNumber" ? 1 : 0);
  return rank(a) - rank(b) || String(b.date ?? "").localeCompare(String(a.date ?? ""));
}
