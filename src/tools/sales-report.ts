import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { InputError } from "../errors.js";
import { resolveNames } from "../odata/accounting.js";
import { callPageStart } from "../odata/call-context.js";
import { loadItemGroups, nomenclatureSet, readAllItems, readItemAttributes } from "../odata/inventory.js";
import { resolveOrganization } from "../odata/orgs.js";
import { PARTNERS_CATALOG } from "../odata/partner-settlements.js";
import {
  ACCOUNTING_DIMS,
  TRADE_DIMS,
  aggregate,
  buckets,
  isTradeReport,
  minusYear,
  previousPeriod,
  readAccountingReport,
  readTradeReport,
  type Aggregated,
  type Dim,
  type Granularity,
  type ReportRow,
  PURCHASES_SPEC,
  SALES_SPEC,
  isPurchasesReport,
} from "../odata/sales-report.js";
import { describeWarehouses } from "../odata/stock.js";
import { lastCostCalculation } from "../odata/trade-cost.js";
import { USERS } from "../odata/users.js";
import { purchasesReportResultSchema, salesReportResultSchema } from "../schemas/output.js";
import { databaseField, dateField, guard, ok } from "./_shared.js";

const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;
const refs = (label: string) =>
  z
    .array(z.string().trim().regex(GUID, `Ref_Key (${label}, GUID)`))
    .min(1)
    .max(50)
    .optional();
const DEPARTMENT_SETS = ["Catalog_СтруктураПредприятия", "Catalog_ПодразделенияОрганизаций"] as const;
const ITEM_GROUP_SETS_BP = ["Catalog_НоменклатурныеГруппы"] as const;

type Kind = "sales" | "purchases";
type Metric = "revenue" | "revenueExVat" | "margin" | "marginExVat" | "quantity" | "discount";

/** Рубли из копеек; undefined — метрика в этом разрезе не считается. */
const rub = (c: number | undefined) => (c === undefined ? undefined : c / 100);

type Sums = Pick<
  Aggregated,
  "quantity" | "revenueCents" | "revenueExVatCents" | "costCents" | "costExVatCents" | "discountCents"
>;

/** Метрики строки. В закупках выручки и маржи нет: amount — сумма закупки, amountExVat — без НДС. */
function metrics(a: Sums, kind: Kind = "sales") {
  const discount = a.discountCents !== undefined ? { discount: rub(a.discountCents) } : {};
  if (kind === "purchases")
    return {
      quantity: Math.round(a.quantity * 1000) / 1000,
      amount: a.revenueCents / 100,
      ...(a.revenueExVatCents !== undefined ? { amountExVat: a.revenueExVatCents / 100 } : {}),
      ...discount,
    };
  const margin = a.costCents === undefined ? undefined : a.revenueCents - a.costCents;
  const marginExVat =
    a.costExVatCents === undefined || a.revenueExVatCents === undefined
      ? undefined
      : a.revenueExVatCents - a.costExVatCents;
  return {
    quantity: Math.round(a.quantity * 1000) / 1000,
    revenue: a.revenueCents / 100,
    ...(a.revenueExVatCents !== undefined ? { revenueExVat: a.revenueExVatCents / 100 } : {}),
    ...(a.costCents !== undefined ? { cost: rub(a.costCents) } : {}),
    ...(a.costExVatCents !== undefined ? { costExVat: rub(a.costExVatCents) } : {}),
    ...(margin !== undefined ? { margin: margin / 100 } : {}),
    ...(marginExVat !== undefined ? { marginExVat: marginExVat / 100 } : {}),
    ...discount,
  };
}

function metricOf(a: Aggregated, m: Metric): number {
  if (m === "quantity") return a.quantity;
  if (m === "revenue") return a.revenueCents / 100;
  if (m === "revenueExVat") return (a.revenueExVatCents ?? 0) / 100;
  if (m === "discount") return (a.discountCents ?? 0) / 100;
  const x = metrics(a) as { margin?: number; marginExVat?: number };
  return x[m] ?? 0;
}

/** Метрика считается по строкам (в итоге не undefined). */
function hasMetric(t: Sums, m: Metric): boolean {
  if (m === "quantity" || m === "revenue") return true;
  if (m === "revenueExVat") return t.revenueExVatCents !== undefined;
  if (m === "discount") return t.discountCents !== undefined;
  if (m === "margin") return t.costCents !== undefined;
  return t.costExVatCents !== undefined && t.revenueExVatCents !== undefined;
}

function total(rows: readonly Aggregated[]): Aggregated {
  const add = (a: number | undefined, b: number | undefined) =>
    a === undefined || b === undefined ? undefined : a + b;
  const first = rows[0];
  return rows.reduce<Aggregated>(
    (t, r) => ({
      key: "",
      dims: {},
      quantity: t.quantity + r.quantity,
      revenueCents: t.revenueCents + r.revenueCents,
      revenueExVatCents: add(t.revenueExVatCents, r.revenueExVatCents),
      costCents: add(t.costCents, r.costCents),
      costExVatCents: add(t.costExVatCents, r.costExVatCents),
      discountCents: add(t.discountCents, r.discountCents),
    }),
    {
      key: "",
      dims: {},
      quantity: 0,
      revenueCents: 0,
      // Пустой отбор: метрики, которых у этого регистра нет, не выводятся и как ноль.
      revenueExVatCents: first ? 0 : 0,
      costCents: first && first.costCents === undefined ? undefined : 0,
      costExVatCents: first && first.costExVatCents === undefined ? undefined : 0,
      discountCents: first && first.discountCents === undefined ? undefined : 0,
    },
  );
}

/** Названия значений разрезов страницы: справочник каждого разреза — одним запросом. */
async function dimNames(
  conn: Connection,
  trade: boolean,
  dims: readonly Dim[],
  page: readonly Aggregated[],
  groups: Map<string, string>,
): Promise<Map<Dim, Map<string, string>>> {
  const available = await conn.available();
  const values = (d: Dim) => [...new Set(page.flatMap((a) => (a.dims[d] ? [a.dims[d]!] : [])))];
  const out = new Map<Dim, Map<string, string>>();
  const safe = (p: Promise<Map<string, string>>) => p.catch(() => new Map<string, string>());
  for (const d of dims) {
    const v = values(d);
    if (!v.length) continue;
    if (d === "item") {
      const set = await nomenclatureSet(conn);
      out.set(
        d,
        set
          ? new Map([...(await readItemAttributes(conn, set, v))].map(([k, a]) => [k, a.name ?? k]))
          : new Map(),
      );
    } else if (d === "itemGroup") {
      if (trade) out.set(d, new Map(v.map((k) => [k, groups.get(k) ?? k])));
      else {
        const set = resolveEntity(ITEM_GROUP_SETS_BP, available);
        out.set(d, set ? await safe(resolveNames(conn, set, v)) : new Map());
      }
    } else if (d === "warehouse")
      out.set(d, new Map([...(await describeWarehouses(conn, v))].map(([k, w]) => [k, w.name ?? k])));
    else if (d === "department") {
      const set = resolveEntity(DEPARTMENT_SETS, available);
      out.set(d, set ? await safe(resolveNames(conn, set, v)) : new Map());
    } else if (d === "manager")
      out.set(d, available.has(USERS) ? await safe(resolveNames(conn, USERS, v)) : new Map());
    else if (d === "partner")
      out.set(
        d,
        available.has(PARTNERS_CATALOG) ? await safe(resolveNames(conn, PARTNERS_CATALOG, v)) : new Map(),
      );
    else if (d === "organization") {
      const set = resolveEntity(CATALOGS.organizations, available);
      out.set(d, set ? await safe(resolveNames(conn, set, v)) : new Map());
    }
  }
  return out;
}

export function registerSalesReportTools(server: McpServer, ctx: ServerContext): void {
  registerReport(server, ctx, "sales");
  registerReport(server, ctx, "purchases");
}

const PURCHASES_DESCRIPTION =
  "Закупки за период в разрезах на выбор (Управление торговлей, регистр «Закупки»; в Бухгалтерии закупки по " +
  "документам — get_purchases_breakdown). groupBy — до трёх разрезов: item (товар), itemGroup (группа товара), warehouse " +
  "(склад), department, manager, partner (поставщик), organization, operation, period. period — day, week, month (по " +
  "умолчанию), quarter. У строки: quantity, amount и amountExVat (сумма закупки с НДС и без), discount (скидка поставщика). " +
  "compareWith — previousPeriod или previousYear: у строки prev и growth (рост суммы в %), итог сравнения — в compare. " +
  "sortBy — amount (по умолчанию), amountExVat, quantity, discount; abc=true — класс A/B/C по накопленной доле. Отборы: " +
  "organization, warehouses, departments, managers, partners (Ref_Key поставщиков). count — всего строк; длинный список — " +
  "с offset=nextOffset.";

function registerReport(server: McpServer, ctx: ServerContext, kind: Kind): void {
  const name = kind === "sales" ? "read.analytics.get_sales_report" : "read.analytics.get_purchases_report";
  const sortEnum =
    kind === "sales"
      ? (["revenue", "revenueExVat", "margin", "marginExVat", "quantity", "discount"] as const)
      : (["amount", "amountExVat", "quantity", "discount"] as const);
  server.registerTool(
    name,
    {
      title: kind === "sales" ? "Продажи в разрезах" : "Закупки в разрезах",
      description:
        kind === "purchases"
          ? PURCHASES_DESCRIPTION
          : "Продажи за период в разрезах на выбор — один инструмент для Управления торговлей и Бухгалтерии. groupBy — до " +
            "трёх разрезов: в УТ item (товар), itemGroup (группа товара), warehouse (склад или магазин), department " +
            "(подразделение), manager (менеджер), partner (клиент), organization, operation (хозяйственная операция), period; " +
            "в Бухгалтерии item, itemGroup (номенклатурная группа), department, organization, period — по счёту 90 (покупателя у " +
            "счёта нет: разрез по контрагенту и договору — get_sales_breakdown). period — day, week, month (по умолчанию), quarter. " +
            "У строки: quantity, revenue и revenueExVat (выручка с НДС и без), cost, costExVat, margin, marginExVat " +
            "(себестоимость и валовая прибыль; до расчёта себестоимости за месяц — предварительные, дату называет note). " +
            "Метрика, которую нельзя разделить по разрезу (в Бухгалтерии себестоимость и НДС без номенклатуры), не выводится — " +
            "сказано в note. compareWith — previousPeriod (такой же период перед этим) или previousYear: у строки prev и " +
            "growth (рост выручки в %), итог сравнения — в compare. sortBy — по чему сортировать (по умолчанию revenue); abc=true — " +
            "класс A/B/C по накопленной доле этой метрики (80 % / 95 %). Отборы: organization, а в УТ ещё warehouses, " +
            "departments, managers, partners (Ref_Key). count — всего строк; длинный список — с offset=nextOffset.",
      inputSchema: z.object({
        database: databaseField,
        from: dateField("Начало периода (включительно)"),
        to: dateField("Конец периода (включительно)"),
        groupBy: z
          .array(
            z.enum([
              "item",
              "itemGroup",
              "warehouse",
              "department",
              "manager",
              "partner",
              "organization",
              "operation",
              "period",
            ]),
          )
          .min(1)
          .max(3)
          .describe('Разрезы по порядку, до трёх: напр. ["department", "period"]'),
        period: z
          .enum(["day", "week", "month", "quarter"])
          .optional()
          .describe("Шаг разреза period (по умолчанию month)"),
        compareWith: z
          .enum(["previousPeriod", "previousYear"])
          .optional()
          .describe("Сравнить с таким же периодом перед этим или с тем же периодом год назад"),
        sortBy: z
          .enum(sortEnum as unknown as [string, ...string[]])
          .default(sortEnum[0])
          .describe("Сортировка по убыванию (и метрика для abc)"),
        abc: z
          .boolean()
          .default(false)
          .describe("Класс A/B/C по накопленной доле sortBy: A — до 80 %, B — до 95 %"),
        organization: z
          .string()
          .trim()
          .min(1)
          .optional()
          .describe("Организация: часть названия, ИНН или Ref_Key"),
        warehouses: refs("склада").describe("УТ: только эти склады"),
        departments: refs("подразделения").describe("Только эти подразделения"),
        managers: refs("пользователя").describe("УТ: только эти менеджеры (read.catalog.list_users)"),
        partners: refs("партнёра").describe(
          kind === "sales" ? "УТ: только эти клиенты (партнёры)" : "Только эти поставщики (партнёры)",
        ),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько строк вернуть"),
      }),
      outputSchema: kind === "sales" ? salesReportResultSchema : purchasesReportResultSchema,
    },
    (a) =>
      guard(name, async () => {
        const t0 = Date.now();
        const conn = ctx.db(a.database);
        const dims = [...new Set(a.groupBy)] as Dim[];
        const trade = kind === "sales" ? await isTradeReport(conn) : await isPurchasesReport(conn);
        if (kind === "purchases" && !trade)
          throw new InputError(
            "Закупки в разрезах — по регистру «Закупки» Управления торговлей; в этой базе его нет. В Бухгалтерии закупки по документам — read.analytics.get_purchases_breakdown.",
          );
        const allowed = trade ? TRADE_DIMS : ACCOUNTING_DIMS;
        const bad = dims.filter((d) => !allowed.includes(d));
        if (bad.length)
          throw new InputError(
            trade
              ? `Разрезы ${bad.join(", ")} в Управлении торговлей не поддержаны.`
              : `Разрезы ${bad.join(", ")} в Бухгалтерии недоступны: у счёта 90 есть только номенклатурная группа, номенклатура, подразделение и организация. По контрагенту и договору — read.analytics.get_sales_breakdown.`,
          );
        if (!trade && (a.warehouses || a.managers || a.partners))
          throw new InputError("Отборы warehouses, managers и partners — для Управления торговлей.");
        const granularity: Granularity | undefined = dims.includes("period")
          ? (a.period ?? "month")
          : undefined;
        const org = a.organization ? await resolveOrganization(conn, a.organization) : undefined;
        const notes: string[] = [];

        const read = async (from: string, to: string) => {
          const parts = buckets(from, to, granularity);
          if (trade)
            return {
              parts,
              ...(await readTradeReport(
                conn,
                dims,
                parts,
                {
                  warehouses: a.warehouses,
                  departments: a.departments,
                  managers: a.managers,
                  partners: a.partners,
                  orgKey: org?.ref,
                },
                kind === "sales" ? SALES_SPEC : PURCHASES_SPEC,
              )),
              missing: [] as string[],
            };
          return {
            parts,
            ...(await readAccountingReport(conn, dims, parts, {
              orgKey: org?.ref,
              departments: a.departments,
            })),
          };
        };
        const cur = await read(a.from, a.to);
        // Группа товара в УТ — папка номенклатуры (справочник читается один раз и кэшируется).
        let groupOf: ((item: string | undefined) => string | undefined) | undefined;
        let groupNames = new Map<string, string>();
        if (trade && dims.includes("itemGroup")) {
          const set = await nomenclatureSet(conn);
          if (set) {
            const items = await readAllItems(conn, set);
            groupOf = (i) => (i ? items.get(i)?.groupRef : undefined);
            groupNames = new Map((await loadItemGroups(conn, set)).map((g) => [g.ref, g.name]));
          }
        }
        const rows = aggregate(cur.rows as ReportRow[], dims, groupOf);
        // Метрики, которой в этом разрезе нет (Бухгалтерия: себестоимость и НДС без субконто разреза), сортировать нечем:
        // раньше все строки шли нулями — порядок по ключу и класс C у всех. Отказ, как на неподдерживаемый разрез.
        const wanted = (
          a.sortBy === "amount" ? "revenue" : a.sortBy === "amountExVat" ? "revenueExVat" : a.sortBy
        ) as Metric | undefined;
        if (wanted && rows.length && !hasMetric(total(rows), wanted))
          throw new InputError(
            `sortBy: ${a.sortBy} — этой метрики в выбранном разрезе нет (${cur.missing.length ? `${cur.missing.map((m) => (m === "vat" ? "НДС (90.03)" : "себестоимость (90.02)")).join(" и ")} — без субконто этого разреза` : "не считается по этому регистру"}). Сортируйте по ${kind === "sales" ? "revenue" : "amount"} или quantity либо возьмите другой разрез.`,
          );

        // Сравнение: тот же разрез за прошлый период; period-разрез сопоставляется по номеру части.
        let prevByKey: Map<string, Aggregated> | undefined;
        let compare: Record<string, unknown> | undefined;
        if (a.compareWith) {
          const range =
            a.compareWith === "previousYear"
              ? { from: minusYear(a.from), to: minusYear(a.to) }
              : previousPeriod(a.from, a.to);
          const prev = await read(range.from, range.to);
          // Части прошлого периода без пары (в марте недель больше, чем в апреле) — своей меткой с пометкой: голая
          // метка могла совпасть с меткой текущей части и сложиться с чужой (ревью 1.1: неделя «30.03» = 1.03 + 30–31.03).
          const labelMap = new Map(
            prev.parts.map((p, i) => [p.label, cur.parts[i]?.label ?? `${p.label} (прошлый период)`]),
          );
          const shifted = (prev.rows as ReportRow[]).map((r) =>
            r.dims.period
              ? { ...r, dims: { ...r.dims, period: labelMap.get(r.dims.period) ?? r.dims.period } }
              : r,
          );
          const prevRows = aggregate(shifted, dims, groupOf);
          prevByKey = new Map(prevRows.map((r) => [r.key, r]));
          const now = new Set(rows.map((r) => r.key));
          const gone = prevRows.filter((r) => !now.has(r.key));
          if (gone.length)
            notes.push(
              `${gone.length} строк разреза были только в прошлом периоде (${kind === "sales" ? "выручка" : "сумма"} ${total(gone).revenueCents / 100}): в items их нет, в compare они учтены.`,
            );
          compare = { from: range.from, to: range.to, ...metrics(total(prevRows), kind) };
        }

        const sortBy = (
          a.sortBy === "amount" ? "revenue" : a.sortBy === "amountExVat" ? "revenueExVat" : a.sortBy
        ) as Metric;
        const sorted = rows.sort(
          (x, y) => metricOf(y, sortBy) - metricOf(x, sortBy) || x.key.localeCompare(y.key),
        );
        // ABC — по накопленной доле положительной метрики (возвраты и минусовые строки — в C).
        const abcOf = new Map<string, string>();
        if (a.abc) {
          const sum = sorted.reduce((t, r) => t + Math.max(metricOf(r, sortBy), 0), 0);
          let acc = 0;
          for (const r of sorted) {
            const v = Math.max(metricOf(r, sortBy), 0);
            const share = sum > 0 ? (acc + v) / sum : 1;
            abcOf.set(r.key, v <= 0 ? "C" : share <= 0.8 || acc === 0 ? "A" : share <= 0.95 ? "B" : "C");
            acc += v;
          }
        }
        const shown = sorted.slice(0, a.limit);
        const page = shown.slice(callPageStart());
        const names = await dimNames(conn, trade, dims, page, groupNames);
        const growth = (now: number, was: number) =>
          was === 0 ? null : Math.round(((now - was) / Math.abs(was)) * 1000) / 10;

        const items = shown.map((r) => {
          const out: Record<string, unknown> = {};
          for (const d of dims) {
            const v = r.dims[d];
            if (d === "period" || d === "operation") out[d] = v ?? null;
            else {
              out[d] = v ? (names.get(d)?.get(v) ?? v) : null;
              if (v) out[`${d}Ref`] = v;
            }
          }
          const p = prevByKey?.get(r.key);
          return {
            ...out,
            ...metrics(r, kind),
            ...(prevByKey
              ? {
                  prev: p ? metrics(p, kind) : null,
                  growth: p ? growth(r.revenueCents, p.revenueCents) : null,
                }
              : {}),
            ...(a.abc ? { abcClass: abcOf.get(r.key) } : {}),
          };
        });

        if (cur.missing.length)
          notes.push(
            `В Бухгалтерии ${cur.missing.map((m) => (m === "vat" ? "НДС (90.03)" : "себестоимость (90.02)")).join(" и ")} ведётся без нужного субконто — в этом разрезе не делится и не выводится; итог по ней — при разрезе без номенклатуры.`,
          );
        const calc =
          kind === "sales" ? await lastCostCalculation(conn, a.to).catch(() => undefined) : undefined;
        notes.push(
          kind === "purchases"
            ? "Регистр «Закупки» Управления торговлей; суммы — в валюте управленческого учёта."
            : trade
              ? `Регистр продаж Управления торговлей; суммы — в валюте управленческого учёта.${calc ? ` Себестоимость — по расчёту на ${calc}, продажи после него — предварительно.` : " Расчёт себестоимости не найден: себестоимость предварительная."}`
              : "Бухгалтерия: выручка — Кт 90.01, НДС — Дт 90.03, себестоимость — Дт 90.02 (себестоимость появляется после закрытия месяца).",
        );
        if (a.compareWith)
          notes.push(
            kind === "sales"
              ? "growth — рост выручки к прошлому периоду в процентах; null — в прошлом периоде продаж не было."
              : "growth — рост суммы закупок к прошлому периоду в процентах; null — в прошлом периоде закупок не было.",
          );

        return ok({
          database: conn.cfg.name,
          configuration: trade ? "trade" : "accounting",
          from: a.from,
          to: a.to,
          groupBy: dims,
          ...(granularity ? { period: granularity } : {}),
          ...(org ? { organization: org.name } : {}),
          ...metrics(total(rows), kind),
          ...(compare ? { compare } : {}),
          count: sorted.length,
          note: notes.join(" "),
          items,
          scan: { rowsScanned: cur.scanned, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
