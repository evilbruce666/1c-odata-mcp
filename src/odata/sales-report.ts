import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { nextDay, num, resolveAccounts, turnoversByAccounts } from "./accounting.js";
import { tradeWithoutAccounting } from "./configuration.js";
import { requireEntity } from "./publication.js";
import { refOf } from "./settlements-aging.js";
import { SALES_REGISTER } from "./item-sales.js";

/**
 * Разрезы продаж (1.1) — одна модель для Управления торговлей и Бухгалтерии.
 *
 * УТ: регистр «Выручка и себестоимость продаж», обороты со свёрткой на стороне 1С (параметр Dimensions — только
 * нужные измерения): на копии УТ 11.5 за месяц 40–180 строк за 0,3 с вместо 6 тыс. строк за 2,4 с, итог тот же.
 * Бухгалтерия: обороты счетов 90 Хозрасчётного — выручка Кт 90.01, НДС Дт 90.03, себестоимость Дт 90.02; разрезы —
 * субконто (номенклатурная группа, номенклатура) и измерения регистра (подразделение, организация). Покупателя у
 * счёта 90 нет — разрез по контрагенту в Бухгалтерии ведёт get_sales_breakdown (по документам).
 * Период — отдельным запросом на каждую часть: поле периода в оборотах OData 1С выбрать не даёт (проверено).
 */

export type Dim =
  | "item"
  | "itemGroup"
  | "warehouse"
  | "department"
  | "manager"
  | "partner"
  | "organization"
  | "operation"
  | "period";
export const TRADE_DIMS: readonly Dim[] = [
  "item",
  "itemGroup",
  "warehouse",
  "department",
  "manager",
  "partner",
  "organization",
  "operation",
  "period",
];
export const ACCOUNTING_DIMS: readonly Dim[] = ["item", "itemGroup", "department", "organization", "period"];
export type Granularity = "day" | "week" | "month" | "quarter";

/** Строка до свёртки: значения разрезов (ссылки или текст) и суммы в копейках. */
export interface ReportRow {
  dims: Partial<Record<Dim, string>>;
  quantity: number;
  revenueCents: number;
  /** undefined — нельзя посчитать в этом разрезе (Бухгалтерия: НДС или себестоимость ведутся без нужного субконто). */
  revenueExVatCents?: number | undefined;
  costCents?: number | undefined;
  costExVatCents?: number | undefined;
  /** Скидки (УТ): ручная и автоматическая; undefined — не ведутся в регистре. */
  discountCents?: number | undefined;
}

export interface Bucket {
  label: string;
  from: string;
  to: string;
}

const MAX_BUCKETS = 62;
const pad = (n: number) => String(n).padStart(2, "0");
const ymd = (d: Date) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
const parse = (s: string) => new Date(`${s}T00:00:00Z`);

/** Части периода: день, неделя (с понедельника), месяц, квартал; первая и последняя обрезаны по from/to. */
export function buckets(from: string, to: string, g: Granularity | undefined): Bucket[] {
  nextDay(from);
  nextDay(to);
  if (from > to) throw new InputError(`Начало периода ${from} позже конца ${to}.`);
  if (!g) return [{ label: `${from}..${to}`, from, to }];
  const out: Bucket[] = [];
  let cur = parse(from);
  const end = parse(to);
  while (cur <= end) {
    const start = new Date(cur);
    let next: Date;
    let label: string;
    if (g === "day") {
      next = new Date(start);
      next.setUTCDate(next.getUTCDate() + 1);
      label = ymd(start);
    } else if (g === "week") {
      const monday = new Date(start);
      monday.setUTCDate(monday.getUTCDate() - ((monday.getUTCDay() + 6) % 7));
      next = new Date(monday);
      next.setUTCDate(next.getUTCDate() + 7);
      label = `${ymd(monday)} (неделя)`;
    } else if (g === "month") {
      next = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + 1, 1));
      label = `${start.getUTCFullYear()}-${pad(start.getUTCMonth() + 1)}`;
    } else {
      const q = Math.floor(start.getUTCMonth() / 3);
      next = new Date(Date.UTC(start.getUTCFullYear(), q * 3 + 3, 1));
      label = `${start.getUTCFullYear()}-Q${q + 1}`;
    }
    const last = new Date(next);
    last.setUTCDate(last.getUTCDate() - 1);
    out.push({ label, from: ymd(start), to: ymd(last > end ? end : last) });
    if (out.length > MAX_BUCKETS)
      throw new InputError(
        `Слишком много частей периода (больше ${MAX_BUCKETS}): возьмите шаг крупнее (неделя, месяц или квартал) или сократите интервал.`,
      );
    cur = next;
  }
  return out;
}

/** Даты, сдвинутые на год назад (29 февраля → 28). */
export function minusYear(d: string): string {
  const [y, m, day] = d.split("-").map(Number) as [number, number, number];
  const leap = (x: number) => (x % 4 === 0 && x % 100 !== 0) || x % 400 === 0;
  const dd = m === 2 && day === 29 && !leap(y - 1) ? 28 : day;
  return `${y - 1}-${pad(m)}-${pad(dd)}`;
}

/**
 * Прошлый период для сравнения. Целые месяцы (с 1-го по последнее число) — столько же месяцев назад: октябрь →
 * сентябрь, IV квартал → III (живая проверка: «31 день назад» давал 31 августа — 30 сентября и сравнивал октябрь с
 * одним днём августа). Иначе — столько же дней прямо перед периодом.
 */
export function previousPeriod(from: string, to: string): { from: string; to: string } {
  const f = parse(from);
  const t = parse(to);
  const endOfMonth = (d: Date) =>
    new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate() === d.getUTCDate();
  if (f.getUTCDate() === 1 && endOfMonth(t)) {
    const months = (t.getUTCFullYear() - f.getUTCFullYear()) * 12 + t.getUTCMonth() - f.getUTCMonth() + 1;
    const pf = new Date(Date.UTC(f.getUTCFullYear(), f.getUTCMonth() - months, 1));
    const pt = new Date(Date.UTC(f.getUTCFullYear(), f.getUTCMonth(), 0));
    return { from: ymd(pf), to: ymd(pt) };
  }
  const days = Math.round((t.getTime() - f.getTime()) / 86_400_000) + 1;
  const pt = new Date(f);
  pt.setUTCDate(pt.getUTCDate() - 1);
  const pf = new Date(pt);
  pf.setUTCDate(pf.getUTCDate() - days + 1);
  return { from: ymd(pf), to: ymd(pt) };
}

export async function isPurchasesReport(conn: Connection): Promise<boolean> {
  return (await tradeWithoutAccounting(conn)) && (await conn.available()).has("AccumulationRegister_Закупки");
}

export async function isTradeReport(conn: Connection): Promise<boolean> {
  return (await tradeWithoutAccounting(conn)) && (await conn.available()).has(SALES_REGISTER);
}

// ─── Управление торговлей ───

const KEY_ITEM = "АналитикаУчетаНоменклатуры";
const KEY_PARTNER = "АналитикаУчетаПоПартнерам";

/** Регистр оборотов УТ для разрезов: продажи или закупки (1.1). */
export interface TradeRegisterSpec {
  set: string;
  label: string;
  /** Партнёр и организация: через ключ аналитики по партнёрам (продажи) или прямыми измерениями (закупки). */
  partner: "key" | "direct";
  quantity: string;
  amount: string;
  amountExVat: string;
  cost?: readonly string[];
  costExVat?: readonly string[];
  discount?: readonly string[];
}

export const SALES_SPEC: TradeRegisterSpec = {
  set: SALES_REGISTER,
  label: "Регистр накопления «Выручка и себестоимость продаж»",
  partner: "key",
  quantity: "КоличествоTurnover",
  amount: "СуммаВыручкиTurnover",
  amountExVat: "СуммаВыручкиБезНДСTurnover",
  cost: ["СтоимостьTurnover", "ДопРасходыTurnover"],
  costExVat: ["СтоимостьБезНДСTurnover", "ДопРасходыБезНДСTurnover"],
  // Скидки (ручная и автоматическая) — в той же валюте, что выручка: на копии УТ за октябрь ручная 1,17 млн.
  discount: ["СуммаРучнойСкидкиTurnover", "СуммаАвтоматическойСкидкиTurnover"],
};

export const PURCHASES_REGISTER = "AccumulationRegister_Закупки";
export const PURCHASES_SPEC: TradeRegisterSpec = {
  set: PURCHASES_REGISTER,
  label: "Регистр накопления «Закупки»",
  partner: "direct",
  quantity: "КоличествоTurnover",
  amount: "СуммаTurnover",
  amountExVat: "СуммаБезНДСTurnover",
  discount: ["СуммаСкидкиTurnover"],
};

export interface TradeFilter {
  warehouses?: readonly string[] | undefined;
  departments?: readonly string[] | undefined;
  managers?: readonly string[] | undefined;
  partners?: readonly string[] | undefined;
  orgKey?: string | undefined;
}

export async function readTradeReport(
  conn: Connection,
  dims: readonly Dim[],
  parts: readonly Bucket[],
  filter: TradeFilter,
  spec: TradeRegisterSpec = SALES_SPEC,
): Promise<{ rows: ReportRow[]; scanned: number }> {
  const register = await requireEntity(conn, [spec.set], spec.label);
  const need = new Set(dims);
  if (filter.warehouses?.length) need.add("warehouse");
  if (filter.departments?.length) need.add("department");
  if (filter.managers?.length) need.add("manager");
  if (filter.partners?.length) need.add("partner");
  if (filter.orgKey) need.add("organization");
  const registerDims = new Set<string>();
  const select: string[] = [];
  const expand = new Set<string>();
  if (need.has("item") || need.has("itemGroup")) {
    registerDims.add(KEY_ITEM);
    select.push(`${KEY_ITEM}_Key`, `${KEY_ITEM}/Номенклатура_Key`);
    expand.add(KEY_ITEM);
  }
  if (spec.partner === "key" && (need.has("partner") || need.has("organization"))) {
    registerDims.add(KEY_PARTNER);
    select.push(`${KEY_PARTNER}_Key`, `${KEY_PARTNER}/Партнер_Key`, `${KEY_PARTNER}/Организация_Key`);
    expand.add(KEY_PARTNER);
  }
  if (spec.partner === "direct" && need.has("partner")) {
    registerDims.add("Партнер");
    select.push("Партнер_Key");
  }
  if (spec.partner === "direct" && need.has("organization")) {
    registerDims.add("Организация");
    select.push("Организация_Key");
  }
  if (need.has("warehouse")) {
    registerDims.add("Склад");
    select.push("Склад", "Склад_Type");
  }
  if (need.has("department")) {
    registerDims.add("Подразделение");
    select.push("Подразделение_Key");
  }
  if (need.has("manager")) {
    registerDims.add("Менеджер");
    select.push("Менеджер_Key");
  }
  if (need.has("operation")) {
    registerDims.add("ХозяйственнаяОперация");
    select.push("ХозяйственнаяОперация");
  }
  const resources = [
    spec.quantity,
    spec.amount,
    spec.amountExVat,
    ...(spec.cost ?? []),
    ...(spec.costExVat ?? []),
    ...(spec.discount ?? []),
  ];
  const lower = (xs?: readonly string[]) =>
    xs?.length ? new Set(xs.map((x) => x.replace(/[{}]/g, "").toLowerCase())) : undefined;
  const fw = lower(filter.warehouses);
  const fd = lower(filter.departments);
  const fm = lower(filter.managers);
  const fp = lower(filter.partners);
  const ok = (v: string | undefined, f: Set<string> | undefined) => !f || (!!v && f.has(v.toLowerCase()));
  const sum = (r: ODataEntity, fs: readonly string[] | undefined) =>
    fs ? Math.round(fs.reduce((t, f) => t + num(r[f]), 0) * 100) : undefined;
  const rows: ReportRow[] = [];
  let scanned = 0;
  for (const b of parts) {
    const params = [
      `StartPeriod=datetime'${b.from}T00:00:00'`,
      `EndPeriod=datetime'${b.to}T23:59:59'`,
      // Свёртка на стороне 1С — только по нужным измерениям.
      `Dimensions='${[...registerDims].join(",")}'`,
    ];
    const { rows: raw } = await fetchAllForAggregation(
      conn,
      `${register}/Turnovers(${params.join(",")})`,
      { select: [...select, ...resources], ...(expand.size ? { expand: [...expand] } : {}) },
      `${b.from}…${b.to}`,
    );
    scanned += raw.length;
    for (const r of raw) {
      const item = (r[KEY_ITEM] ?? {}) as ODataEntity;
      const key = (r[KEY_PARTNER] ?? {}) as ODataEntity;
      const partnerRef = spec.partner === "key" ? refOf(key["Партнер_Key"]) : refOf(r["Партнер_Key"]);
      const orgRef = spec.partner === "key" ? refOf(key["Организация_Key"]) : refOf(r["Организация_Key"]);
      const warehouse = r["Склад_Type"] === "StandardODATA.Catalog_Склады" ? refOf(r["Склад"]) : undefined;
      const v: Partial<Record<Dim, string>> = {
        ...(need.has("item") || need.has("itemGroup") ? { item: refOf(item["Номенклатура_Key"]) } : {}),
        ...(need.has("warehouse") ? { warehouse } : {}),
        ...(need.has("department") ? { department: refOf(r["Подразделение_Key"]) } : {}),
        ...(need.has("manager") ? { manager: refOf(r["Менеджер_Key"]) } : {}),
        ...(need.has("partner") ? { partner: partnerRef } : {}),
        ...(need.has("organization") ? { organization: orgRef } : {}),
        ...(need.has("operation")
          ? { operation: String(r["ХозяйственнаяОперация"] ?? "") || undefined }
          : {}),
        ...(need.has("period") ? { period: b.label } : {}),
      };
      if (!ok(v.warehouse, fw) || !ok(v.department, fd) || !ok(v.manager, fm) || !ok(v.partner, fp)) continue;
      if (filter.orgKey && v.organization?.toLowerCase() !== filter.orgKey.toLowerCase()) continue;
      rows.push({
        dims: v,
        quantity: num(r[spec.quantity]),
        revenueCents: Math.round(num(r[spec.amount]) * 100),
        revenueExVatCents: Math.round(num(r[spec.amountExVat]) * 100),
        costCents: sum(r, spec.cost),
        costExVatCents: sum(r, spec.costExVat),
        discountCents: sum(r, spec.discount),
      });
    }
  }
  return { rows, scanned };
}

// ─── Бухгалтерия ───

const TYPE_ITEM_GROUP = "StandardODATA.Catalog_НоменклатурныеГруппы";
const TYPE_ITEM = "StandardODATA.Catalog_Номенклатура";

/** Субконто строки оборотов заданного типа. */
function subconto(r: ODataEntity, type: string): string | undefined {
  for (const i of [1, 2, 3]) if (r[`ExtDimension${i}_Type`] === type) return refOf(r[`ExtDimension${i}`]);
  return undefined;
}

export interface AccountingFilter {
  orgKey?: string | undefined;
  departments?: readonly string[] | undefined;
}

/**
 * Выручка, НДС и себестоимость по счетам 90 в разрезе. Если у счёта НДС (90.03) или себестоимости (90.02) нет
 * нужного субконто (номенклатура обычно только у 90.01.1), его сумма не делится по разрезу — тогда метрика
 * не считается (undefined), а не кладётся целиком в «без разреза».
 */
export async function readAccountingReport(
  conn: Connection,
  dims: readonly Dim[],
  parts: readonly Bucket[],
  filter: AccountingFilter,
): Promise<{ rows: ReportRow[]; scanned: number; missing: Array<"vat" | "cost"> }> {
  const [revenue, vat, cost] = await Promise.all([
    resolveAccounts(conn, ["90.01"]),
    resolveAccounts(conn, ["90.03"]),
    resolveAccounts(conn, ["90.02"]),
  ]);
  if (!revenue.length)
    throw new InputError("Счёт 90.01 не найден в плане счетов: продажи по нему не посчитать.");
  const kind = new Map<string, "revenue" | "vat" | "cost">([
    ...revenue.map((a) => [a.key, "revenue"] as const),
    ...vat.map((a) => [a.key, "vat"] as const),
    ...cost.map((a) => [a.key, "cost"] as const),
  ]);
  const fd = filter.departments?.length
    ? new Set(filter.departments.map((d) => d.replace(/[{}]/g, "").toLowerCase()))
    : undefined;
  const missing = new Set<"vat" | "cost">();
  const rows: ReportRow[] = [];
  let scanned = 0;
  for (const b of parts) {
    const { rows: raw } = await turnoversByAccounts(conn, [...kind.keys()], b.from, b.to, filter.orgKey);
    scanned += raw.length;
    for (const r of raw) {
      const k = kind.get(String(r["Account_Key"]));
      if (!k) continue;
      const department = refOf(r["Подразделение_Key"]);
      if (fd && !(department && fd.has(department.toLowerCase()))) continue;
      const item = subconto(r, TYPE_ITEM);
      const group = subconto(r, TYPE_ITEM_GROUP);
      // Метрика без нужного субконто не делится по разрезу.
      if (k !== "revenue" && dims.includes("item") && !item && num(r["СуммаTurnoverDr"]) !== 0)
        missing.add(k);
      if (k !== "revenue" && dims.includes("itemGroup") && !group && num(r["СуммаTurnoverDr"]) !== 0)
        missing.add(k);
      const v: Partial<Record<Dim, string>> = {
        ...(dims.includes("item") ? { item } : {}),
        ...(dims.includes("itemGroup") ? { itemGroup: group } : {}),
        ...(dims.includes("department") ? { department } : {}),
        ...(dims.includes("organization") ? { organization: refOf(r["Организация_Key"]) } : {}),
        ...(dims.includes("period") ? { period: b.label } : {}),
      };
      const cents = (f: string) => Math.round(num(r[f]) * 100);
      if (k === "revenue")
        rows.push({
          dims: v,
          // Только Кт 90.01: возвраты текущего года 1С проводит сторно (отрицательным Кт), а Дт 90.01 — это и
          // реформация баланса в конце года (Дт 90.01 Кт 90.09): вычитание обнулило бы выручку декабря.
          quantity: num(r["КоличествоTurnoverCr"]),
          revenueCents: cents("СуммаTurnoverCr"),
          revenueExVatCents: cents("СуммаTurnoverCr"),
          costCents: 0,
          costExVatCents: 0,
        });
      else if (k === "vat")
        // НДС — Дт 90.03 (Кт — реформация).
        rows.push({
          dims: v,
          quantity: 0,
          revenueCents: 0,
          revenueExVatCents: -cents("СуммаTurnoverDr"),
          costCents: 0,
          costExVatCents: 0,
        });
      else {
        // Себестоимость — Дт 90.02 (Кт — реформация).
        const c = cents("СуммаTurnoverDr");
        rows.push({
          dims: v,
          quantity: 0,
          revenueCents: 0,
          revenueExVatCents: 0,
          costCents: c,
          costExVatCents: c,
        });
      }
    }
  }
  if (missing.has("vat") || missing.has("cost"))
    for (const r of rows) {
      if (missing.has("vat")) r.revenueExVatCents = undefined;
      if (missing.has("cost")) {
        r.costCents = undefined;
        r.costExVatCents = undefined;
      }
    }
  return { rows, scanned, missing: [...missing] };
}

// ─── Свёртка ───

export interface Aggregated {
  key: string;
  dims: Partial<Record<Dim, string>>;
  quantity: number;
  revenueCents: number;
  revenueExVatCents?: number | undefined;
  costCents?: number | undefined;
  costExVatCents?: number | undefined;
  discountCents?: number | undefined;
}

/** Свёртка по выбранным разрезам; itemOf — отображение товара в разрез «группа» (для itemGroup). */
export function aggregate(
  rows: readonly ReportRow[],
  dims: readonly Dim[],
  itemGroupOf?: (item: string | undefined) => string | undefined,
): Aggregated[] {
  const out = new Map<string, Aggregated>();
  const add = (a: number | undefined, b: number | undefined) =>
    a === undefined || b === undefined ? undefined : a + b;
  for (const r of rows) {
    const v: Partial<Record<Dim, string>> = {};
    for (const d of dims) {
      const val =
        d === "itemGroup" && itemGroupOf ? (r.dims.itemGroup ?? itemGroupOf(r.dims.item)) : r.dims[d];
      if (val !== undefined) v[d] = val;
    }
    const key = dims.map((d) => v[d] ?? "").join("|");
    const cur = out.get(key);
    if (!cur)
      out.set(key, {
        key,
        dims: v,
        quantity: r.quantity,
        revenueCents: r.revenueCents,
        revenueExVatCents: r.revenueExVatCents,
        costCents: r.costCents,
        costExVatCents: r.costExVatCents,
        discountCents: r.discountCents,
      });
    else {
      cur.quantity += r.quantity;
      cur.revenueCents += r.revenueCents;
      cur.revenueExVatCents = add(cur.revenueExVatCents, r.revenueExVatCents);
      cur.costCents = add(cur.costCents, r.costCents);
      cur.costExVatCents = add(cur.costExVatCents, r.costExVatCents);
      cur.discountCents = add(cur.discountCents, r.discountCents);
    }
  }
  return [...out.values()].filter(
    (a) => Math.abs(a.quantity) > 0.0001 || a.revenueCents !== 0 || (a.costCents ?? 0) !== 0,
  );
}
