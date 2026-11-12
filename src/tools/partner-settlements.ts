import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import type { Connection } from "../context.js";
import { InputError } from "../errors.js";
import { readByRefs, resolveNames } from "../odata/accounting.js";
import { localToday, refOf } from "../odata/settlements-aging.js";
import { callPageStart } from "../odata/call-context.js";
import { resolveOrganization } from "../odata/orgs.js";
import {
  AGE_BUCKETS,
  PARTNERS_CATALOG,
  ageBucket,
  daysBetween,
  readSettlements,
  type AgeBucket,
} from "../odata/partner-settlements.js";
import { partnerSettlementsResultSchema } from "../schemas/output.js";
import { databaseField, dateField, guard, ok } from "./_shared.js";

const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;
const rub = (cents: number) => cents / 100;

interface Totals {
  debt: number;
  prepayment: number;
  overdue: number;
  maxOverdueDays: number;
  ages: Record<AgeBucket, number>;
}
const emptyTotals = (): Totals => ({
  debt: 0,
  prepayment: 0,
  overdue: 0,
  maxOverdueDays: 0,
  ages: Object.fromEntries(AGE_BUCKETS.map((b) => [b, 0])) as Record<AgeBucket, number>,
});

/**
 * Партнёр по Ref_Key. Ref_Key контрагента (его даёт find_counterparty) переводится в партнёра контрагента: иначе отбор
 * давал тихий ноль. Неизвестный Ref_Key — отказ, как у складов в остатках.
 */
async function resolvePartner(
  conn: Connection,
  ref: string,
): Promise<{ ref: string; fromCounterparty: boolean }> {
  const available = await conn.available();
  const partners = resolveEntity([PARTNERS_CATALOG], available);
  if (partners && (await readByRefs(conn, partners, [ref], ["Ref_Key"])).size)
    return { ref, fromCounterparty: false };
  const counterparties = resolveEntity(CATALOGS.counterparties, available);
  if (counterparties) {
    const props = new Set(
      (await conn.getMetadata()).entities.get(counterparties)?.properties.map((p) => p.name),
    );
    if (props.has("Партнер_Key")) {
      const row = (await readByRefs(conn, counterparties, [ref], ["Ref_Key", "Партнер_Key"]))
        .values()
        .next().value;
      const partner = row ? refOf(row["Партнер_Key"]) : undefined;
      if (partner) return { ref: partner, fromCounterparty: true };
    }
  }
  throw new InputError(`Партнёр ${ref} не найден (ни среди партнёров, ни среди контрагентов с партнёром).`);
}

export function registerPartnerSettlementTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.analytics.get_partner_settlements",
    {
      title: "Расчёты с партнёрами (УТ 11, КА 2, ERP 2)",
      description:
        "Кто кому должен по управленческим расчётам с партнёрами: регистры «Расчёты с клиентами / поставщиками по " +
        "срокам» (УТ 11, КА 2, ERP 2; проверено на УТ 11.5). Это не счёт 62/60 бухучёта (в Бухгалтерии — get_debtors / " +
        "get_creditors; в КА и ERP суммы могут расходиться). side=clients — долг клиентов нам, side=suppliers — наш долг " +
        "поставщикам. debt — долг, prepayment — предоплата (аванс), net = debt − prepayment, всё в регламентированной " +
        "валюте. overdue — долг с плановой датой погашения раньше даты расчёта (on: asOf или сегодня), aging — долг по " +
        "срокам просрочки, maxOverdueDays — самая старая просрочка партнёра. Отбор: organization, partner (Ref_Key " +
        "партнёра), overdueOnly — только партнёры с просрочкой. partners — по убыванию net; count — всего партнёров, " +
        "длинный список читайте с offset=nextOffset.",
      inputSchema: z.object({
        database: databaseField,
        side: z
          .enum(["clients", "suppliers"])
          .describe("clients — долги клиентов нам; suppliers — наши долги поставщикам"),
        asOf: dateField("Дата расчётов — на конец этой даты (без параметра — текущие)").optional(),
        organization: z
          .string()
          .trim()
          .min(1)
          .optional()
          .describe("Организация (юрлицо): часть названия, ИНН или Ref_Key из list_organizations"),
        partner: z
          .string()
          .trim()
          .regex(GUID, "Ref_Key партнёра (GUID)")
          .optional()
          .describe("Ref_Key партнёра"),
        overdueOnly: z.boolean().default(false).describe("Только партнёры с просроченным долгом"),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько партнёров вернуть"),
      }),
      outputSchema: partnerSettlementsResultSchema,
    },
    ({ database, side, asOf, organization, partner, overdueOnly, limit }) =>
      guard("read.analytics.get_partner_settlements", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const resolved = partner
          ? await resolvePartner(conn, partner.replace(/[{}]/g, "").toLowerCase())
          : undefined;
        const partnerRef = resolved?.ref.toLowerCase();
        // Местная дата, как у дебиторки Бухгалтерии: по UTC на востоке России утром «сегодня» было бы вчера.
        const on = asOf ?? localToday();
        const { rows, scanned } = await readSettlements(conn, side, asOf);
        const selected = rows.filter(
          (r) =>
            (!org || r.organizationRef?.toLowerCase() === org.ref.toLowerCase()) &&
            (!partnerRef || r.partnerRef?.toLowerCase() === partnerRef),
        );

        const byPartner = new Map<string, Totals>();
        const total = emptyTotals();
        for (const r of selected) {
          const t = byPartner.get(r.partnerRef ?? "") ?? emptyTotals();
          for (const acc of [t, total]) {
            acc.debt += r.debtCents;
            acc.prepayment += r.prepaymentCents;
            // Отрицательная строка долга (сторно) тоже идёт в сроки и просрочку: сумма aging всегда равна debt, а
            // просрочка не больше долга (ревью: +100 и −30 давали debt 70 при overdue 100).
            const bucket = ageBucket(r.dueDate, on);
            acc.ages[bucket] += r.debtCents;
            if (r.dueDate && bucket !== "notDue" && bucket !== "noDueDate") {
              acc.overdue += r.debtCents;
              if (r.debtCents > 0)
                acc.maxOverdueDays = Math.max(acc.maxOverdueDays, daysBetween(r.dueDate, on));
            }
          }
          byPartner.set(r.partnerRef ?? "", t);
        }

        const list = [...byPartner]
          .filter(([, t]) => t.debt !== 0 || t.prepayment !== 0)
          .filter(([, t]) => !overdueOnly || t.overdue > 0)
          // Вторичный ключ обязателен: каждая страница считается заново, а порядок равных сумм не гарантирован.
          .sort(
            ([ka, a], [kb, b]) => b.debt - b.prepayment - (a.debt - a.prepayment) || ka.localeCompare(kb),
          );
        const shown = list.slice(0, limit);
        const page = shown.slice(callPageStart());
        const catalog = resolveEntity([PARTNERS_CATALOG], await conn.available());
        const names =
          catalog && page.length
            ? await resolveNames(
                conn,
                catalog,
                page.flatMap(([ref]) => (ref ? [ref] : [])),
              )
            : new Map<string, string>();

        const notes = [
          `Управленческие расчёты по партнёрам на ${on}${asOf ? "" : " (сегодня)"}; суммы — в регламентированной валюте. ` +
            "Предоплата уменьшает net, но не просрочку.",
        ];
        if (total.ages.noDueDate > 0)
          notes.push("Часть долга без плановой даты погашения — в aging.noDueDate.");
        if (byPartner.has(""))
          notes.push("Часть расчётов без партнёра в ключе аналитики — строка partner: null.");
        if (resolved?.fromCounterparty) {
          notes.push(
            "partner — Ref_Key контрагента: расчёты взяты по его партнёру (в УТ долги ведутся по партнёрам).",
          );
        }
        // Долг и предоплата у одного партнёра сразу — обычно незачтённые авансы (на УТ 11.5 так у большинства крупных).
        const both = list.filter(([, t]) => t.debt > 0 && t.prepayment > 0);
        if (both.length) {
          const offset = both.reduce((sum, [, t]) => sum + Math.min(t.debt, t.prepayment), 0);
          notes.push(
            `У ${both.length} партнёров одновременно долг и предоплата (до ${rub(offset)} перекрываются): часто это ` +
              "незачтённые авансы, а не реальный долг; net учитывает предоплату, overdue и aging — нет.",
          );
        }

        return ok({
          database: conn.cfg.name,
          side,
          on,
          ...(asOf ? { asOf } : {}),
          ...(org ? { organization: org.name } : {}),
          debt: rub(total.debt),
          prepayment: rub(total.prepayment),
          net: rub(total.debt - total.prepayment),
          overdue: rub(total.overdue),
          aging: Object.fromEntries(AGE_BUCKETS.map((b) => [b, rub(total.ages[b])])),
          // Всего, а не на странице: сколько в этом ответе — длина списка и returned.
          count: list.length,
          note: notes.join(" "),
          partners: shown.map(([ref, t]) => ({
            partner: ref ? (names.get(ref) ?? ref) : null,
            ...(ref ? { partnerRef: ref } : {}),
            debt: rub(t.debt),
            prepayment: rub(t.prepayment),
            net: rub(t.debt - t.prepayment),
            overdue: rub(t.overdue),
            ...(t.maxOverdueDays ? { maxOverdueDays: t.maxOverdueDays } : {}),
          })),
          scan: { rowsScanned: scanned, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
