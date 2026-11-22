import type { McpServer } from "@modelcontextprotocol/server";
import { USERS } from "../odata/users.js";
import { buckets } from "../odata/sales-report.js";
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
  requireSettlements,
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
        "срокам просрочки (сторно с минусом может сделать отдельную корзину отрицательной, сумма корзин всегда равна debt), maxOverdueDays — самая старая просрочка партнёра. Отбор: organization, partner (Ref_Key " +
        "партнёра), overdueOnly — только партнёры с просрочкой, managers — расчёты этих менеджеров (Ref_Key пользователей 1С, read.catalog.list_users; менеджер — у объекта расчётов); byManager=true — итоги по менеджерам в byManager. schedule (day, week, month) — платёжный календарь: долг по датам планового погашения вперёд до horizon, просрочка, долг за горизонтом и без даты — отдельно. partners — по убыванию net; count — всего партнёров, " +
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
        managers: z
          .array(z.string().trim().regex(GUID, "Ref_Key пользователя (GUID)"))
          .min(1)
          .max(50)
          .optional()
          .describe(
            "Только расчёты этих менеджеров — Ref_Key пользователей 1С (read.catalog.list_users); менеджер — у объекта расчётов (заказ, договор, документ)",
          ),
        schedule: z
          .enum(["day", "week", "month"])
          .optional()
          .describe(
            "Платёжный календарь (schedule): долг по датам планового погашения вперёд от on — по дням, неделям или месяцам; отдельно просрочка, долг за горизонтом и без даты",
          ),
        horizon: dateField(
          "Конец календаря, только вместе с schedule (по умолчанию: day — 30 дней, week — 13 недель, month — 6 месяцев, считая текущие неделю и месяц)",
        ).optional(),
        byManager: z
          .boolean()
          .default(false)
          .describe(
            "Итоги по менеджерам (byManager): долг, предоплата, net и просрочка каждого; без менеджера — отдельно",
          ),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько партнёров вернуть"),
      }),
      outputSchema: partnerSettlementsResultSchema,
    },
    ({
      database,
      side,
      asOf,
      organization,
      partner,
      overdueOnly,
      managers,
      byManager,
      schedule,
      horizon,
      limit,
    }) =>
      guard("read.analytics.get_partner_settlements", async () => {
        const t0 = Date.now();
        if (horizon && !schedule)
          throw new InputError(
            "horizon — конец платёжного календаря: задайте и schedule (day, week или month).",
          );
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        // Сначала регистр: в базе без него отказ должен называть регистр, а не «партнёр не найден».
        await requireSettlements(conn, side);
        const resolved = partner
          ? await resolvePartner(conn, partner.replace(/[{}]/g, "").toLowerCase())
          : undefined;
        const partnerRef = resolved?.ref.toLowerCase();
        // Местная дата, как у дебиторки Бухгалтерии: по UTC на востоке России утром «сегодня» было бы вчера.
        const on = asOf ?? localToday();
        const { rows, scanned } = await readSettlements(conn, side, asOf, byManager || !!managers?.length);
        const wanted = managers?.length
          ? new Set(managers.map((m) => m.replace(/[{}]/g, "").toLowerCase()))
          : undefined;
        const selected = rows.filter(
          (r) =>
            (!org || r.organizationRef?.toLowerCase() === org.ref.toLowerCase()) &&
            (!partnerRef || r.partnerRef?.toLowerCase() === partnerRef) &&
            (!wanted || (!!r.managerRef && wanted.has(r.managerRef.toLowerCase()))),
        );
        // Итоги по менеджерам — по тем же строкам отбора, что и общий итог.
        const byManagerTotals = new Map<string, Totals>();
        if (byManager)
          for (const r of selected) {
            const t = byManagerTotals.get(r.managerRef ?? "") ?? emptyTotals();
            t.debt += r.debtCents;
            t.prepayment += r.prepaymentCents;
            const bucket = ageBucket(r.dueDate, on);
            if (r.dueDate && bucket !== "notDue" && bucket !== "noDueDate") t.overdue += r.debtCents;
            byManagerTotals.set(r.managerRef ?? "", t);
          }
        const managerRefs = [
          ...byManagerTotals.keys(),
          ...(managers ?? []).map((m) => m.replace(/[{}]/g, "").toLowerCase()),
        ].filter(Boolean);
        const managerNames = managerRefs.length
          ? await resolveNames(conn, USERS, managerRefs).catch(() => new Map<string, string>())
          : new Map<string, string>();

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

        // Платёжный календарь: тот же долг «по срокам», что и в итогах, — по датам погашения вперёд.
        let calendar: Record<string, unknown> | undefined;
        if (schedule) {
          // По умолчанию ровно 30 дней, 13 недель и 6 месяцев, считая текущие: раньше +90 и +182 дня давали 14 недель
          // и седьмой обрывок месяца (ревью 1.1).
          const end = horizon ?? defaultHorizon(on, schedule);
          if (end < on) throw new InputError(`horizon ${end} раньше даты расчёта ${on}.`);
          const parts = buckets(
            on,
            end,
            schedule,
            "возьмите schedule крупнее (week или month) или horizon ближе",
          );
          const sums = parts.map(() => 0);
          let overdueC = 0;
          let beyond = 0;
          let noDate = 0;
          for (const r of selected) {
            if (!r.dueDate) noDate += r.debtCents;
            else if (r.dueDate < on) overdueC += r.debtCents;
            else if (r.dueDate > end) beyond += r.debtCents;
            else {
              const k = parts.findIndex((p) => r.dueDate! >= p.from && r.dueDate! <= p.to);
              if (k >= 0) sums[k]! += r.debtCents;
            }
          }
          calendar = {
            from: on,
            to: end,
            step: schedule,
            overdue: rub(overdueC),
            periods: parts.map((p, k) => ({ period: p.label, from: p.from, to: p.to, debt: rub(sums[k]!) })),
            beyondHorizon: rub(beyond),
            noDueDate: rub(noDate),
          };
          notes.push(
            "schedule — долг по плановым датам погашения из тех же расчётов «по срокам»: сумма overdue, periods, beyondHorizon и noDueDate равна debt.",
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
          ...(calendar ? { schedule: calendar } : {}),
          ...(managers?.length
            ? {
                managerFilter: [...new Set(managers.map((m) => m.replace(/[{}]/g, "").toLowerCase()))].map(
                  (m) => ({ ref: m, name: managerNames.get(m) ?? null }),
                ),
              }
            : {}),
          ...(byManager && callPageStart() === 0
            ? {
                byManager: [...byManagerTotals]
                  .filter(([, t]) => t.debt !== 0 || t.prepayment !== 0)
                  .map(([ref, t]) => ({
                    manager: ref ? (managerNames.get(ref) ?? ref) : null,
                    ...(ref ? { managerRef: ref } : {}),
                    debt: rub(t.debt),
                    prepayment: rub(t.prepayment),
                    net: rub(t.debt - t.prepayment),
                    overdue: rub(t.overdue),
                  }))
                  .sort((a, b) => b.net - a.net || String(a.managerRef).localeCompare(String(b.managerRef))),
              }
            : {}),
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

/** Конец календаря по умолчанию: 30 дней; конец 13-й недели и 6-го месяца, считая текущие. */
export function defaultHorizon(on: string, step: "day" | "week" | "month"): string {
  if (step === "day") return addDays(on, 29);
  const d = new Date(`${on}T00:00:00Z`);
  if (step === "week") return addDays(on, 13 * 7 - 1 - ((d.getUTCDay() + 6) % 7));
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 6, 0)).toISOString().slice(0, 10);
}

/** YYYY-MM-DD плюс n дней. */
function addDays(d: string, n: number): string {
  const x = new Date(`${d}T00:00:00Z`);
  x.setUTCDate(x.getUTCDate() + n);
  return x.toISOString().slice(0, 10);
}
