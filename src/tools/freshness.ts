import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { collectDocuments } from "../odata/aggregate.js";
import { documentTitle } from "../odata/names.js";
import { listOrganizations, resolveOrganization } from "../odata/orgs.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { addDays, localToday, refOf } from "../odata/settlements-aging.js";
import { dataFreshnessResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok, organizationField } from "./_shared.js";
import { CLOSING_SET, evaluateClosing, type ClosingOp } from "./month-close.js";

/** Окно «что вносят сейчас». */
export const FRESHNESS_WINDOW_DAYS = 30;
/** Виды документов отчёта: то, по чему обычно судят, внесён ли период. */
const KINDS = [
  DOCUMENTS.sales,
  DOCUMENTS.purchases,
  DOCUMENTS.bankIn,
  DOCUMENTS.bankOut,
  DOCUMENTS.cashIn,
  DOCUMENTS.cashOut,
  DOCUMENTS.issuedInvoice,
  DOCUMENTS.commissionReport,
];
/** Регламентных операций на чтение: их десяток-два в месяц, последний закрытый месяц попадает. */
const CLOSING_OPS_TOP = 300;

const day = (r: { [k: string]: unknown } | undefined): string | undefined =>
  r ? String(r[DOC_FIELDS.date] ?? "").slice(0, 10) || undefined : undefined;

/**
 * Один вид документа. Последняя дата и будущая — отдельно: документ «на 30.11» не значит, что всё
 * внесено по 30.11. Обычно два запроса, вид без свежих документов — один. Документы на удаление не
 * считаются.
 */
async function freshnessOf(conn: Connection, set: string, orgKey: string | undefined, today: string) {
  const base = and(
    cmp(DOC_FIELDS.deletionMark, "eq", "false"),
    orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
  );
  const latest = async (filter: string) =>
    (
      await conn.client.getCollection(
        `${set}?$top=1&$select=${DOC_FIELDS.date},${DOC_FIELDS.number}&$orderby=${DOC_FIELDS.date} desc` +
          (filter ? `&$filter=${encodeURIComponent(filter)}` : ""),
      )
    ).value?.[0];
  const newest = await latest(base);
  const newestDay = day(newest);
  const future = newestDay && newestDay > today ? newestDay : undefined;
  const last = future
    ? await latest(and(base, cmp(DOC_FIELDS.date, "le", `datetime'${today}T23:59:59'`)))
    : newest;
  const lastDay = day(last);
  const from = addDays(today, -FRESHNESS_WINDOW_DAYS + 1);
  const inWindow =
    lastDay && lastDay >= from
      ? (
          await collectDocuments(conn, set, {
            baseFilter: base,
            dateField: DOC_FIELDS.date,
            from,
            to: today,
            select: [DOC_FIELDS.date, DOC_FIELDS.posted],
          })
        ).rows
      : [];
  return {
    entitySet: set,
    title: documentTitle(set),
    lastDate: lastDay ?? null,
    ...(last?.[DOC_FIELDS.number] ? { lastNumber: String(last[DOC_FIELDS.number]).trim() } : {}),
    countInWindow: inWindow.length,
    unpostedInWindow: inWindow.filter((r) => r[DOC_FIELDS.posted] !== true).length,
    ...(future ? { futureDatedUntil: future } : {}),
  };
}

/**
 * Последний закрытый месяц каждой организации — по тем же правилам, что check_month_close.
 * Это дата, а не проверка, что закрыты и все месяцы до него.
 */
export function lastClosedMonths(ops: readonly ClosingOp[]): Map<string, string> {
  const byOrg = new Map<string, Map<string, ClosingOp[]>>();
  for (const o of ops) {
    const months = byOrg.get(o.org ?? "") ?? new Map<string, ClosingOp[]>();
    months.set(o.month, [...(months.get(o.month) ?? []), o]);
    byOrg.set(o.org ?? "", months);
  }
  const out = new Map<string, string>();
  for (const [org, months] of byOrg) {
    const closed = [...months].filter(([, m]) => evaluateClosing(m).closed).map(([month]) => month);
    if (closed.length) out.set(org, closed.sort().at(-1)!);
  }
  return out;
}

async function closingStatus(conn: Connection, orgKey: string | undefined, today: string) {
  const set = resolveEntity(CLOSING_SET, await conn.available());
  if (!set) return undefined;
  const filter = and(
    cmp(DOC_FIELDS.deletionMark, "eq", "false"),
    cmp(DOC_FIELDS.date, "le", `datetime'${today}T23:59:59'`),
    orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
  );
  const { rows } = await conn.client
    .getCollection(
      `${set}?$top=${CLOSING_OPS_TOP}&$orderby=${DOC_FIELDS.date} desc` +
        `&$select=${[DOC_FIELDS.date, DOC_FIELDS.organization, "ВидОперации", "Состояние"].join(",")}` +
        `&$filter=${encodeURIComponent(filter)}`,
    )
    .then((p) => ({ rows: p.value ?? [] }));
  const ops: ClosingOp[] = rows.map((r) => ({
    month: String(r[DOC_FIELDS.date] ?? "").slice(0, 7),
    ...(refOf(r[DOC_FIELDS.organization]) ? { org: refOf(r[DOC_FIELDS.organization]) } : {}),
    kind: String(r["ВидОперации"] ?? ""),
    state: String(r["Состояние"] ?? ""),
  }));
  const names = new Map((await listOrganizations(conn)).map((o) => [o.ref, o.name]));
  return {
    months: [...lastClosedMonths(ops)].map(([org, month]) => ({
      ...(org ? { organization: names.get(org) ?? org } : {}),
      lastClosedMonth: month,
    })),
    // Операций больше окна: организация без закрытого месяца в окне могла закрыться раньше него.
    partial: rows.length >= CLOSING_OPS_TOP,
  };
}

export function registerFreshnessTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.system.data_freshness",
    {
      title: "Свежесть данных",
      description:
        "По какую дату внесены данные в базу: по основным видам документов (реализации, поступления, банк, " +
        "касса, счета-фактуры, отчёты комитенту) — дата и номер последнего документа не позже сегодня, сколько " +
        `документов за последние ${FRESHNESS_WINDOW_DAYS} дней и сколько из них не проведено, будущие даты отдельно; ` +
        "последний закрытый месяц. Для вопросов «есть ли данные за октябрь», «по какое число разнесён банк», " +
        "а также чтобы выбрать базу, когда их несколько. Только показывает факты: старая дата не значит, что " +
        "данных не хватает, вид документа может не использоваться. Полная проверка месяца — check_month_close.",
      inputSchema: z.object({ database: databaseField, organization: organizationField }),
      outputSchema: dataFreshnessResultSchema,
    },
    ({ database, organization }) =>
      guard("read.system.data_freshness", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const today = localToday();
        const available = await conn.available();
        const documents: Array<Awaited<ReturnType<typeof freshnessOf>>> = [];
        const notPublished: string[] = [];
        const failed: Array<{ entitySet: string; error: string }> = [];
        // По одному: облачная 1С одновременные запросы только замедляет.
        for (const names of KINDS) {
          const set = resolveEntity(names, available);
          if (!set) {
            notPublished.push(names[0]!);
            continue;
          }
          try {
            documents.push(await freshnessOf(conn, set, org?.ref, today));
          } catch (e) {
            failed.push({ entitySet: set, error: e instanceof Error ? e.message : String(e) });
          }
        }
        const closing = await closingStatus(conn, org?.ref, today);
        return ok({
          database: conn.cfg.name,
          ...(org ? { organization: org.name } : {}),
          asOf: today,
          windowDays: FRESHNESS_WINDOW_DAYS,
          documents,
          ...(closing ? { closing: closing.months } : {}),
          ...(closing?.partial
            ? {
                closingNote: `Просмотрены последние ${CLOSING_OPS_TOP} регламентных операций: более раннее закрытие здесь не видно, точно — check_month_close.`,
              }
            : {}),
          ...(notPublished.length ? { notPublished } : {}),
          ...(failed.length ? { failed } : {}),
          scan: { elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
