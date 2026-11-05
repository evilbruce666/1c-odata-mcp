import type { ODataEntity } from "../types/odata.js";
import { EMPTY_GUID } from "./accounting.js";

/**
 * Расчёты с контрагентами по срокам: разбор строк сальдо 62/60 (субконто «Контрагенты»,
 * «Договоры», «Документы расчётов») и возраст долга по каждому документу. Чистые функции —
 * запросы к 1С делает вызывающий (src/tools/debts.ts).
 */

export type SettlementSide = "receivable" | "payable";

/** Строка сальдо в копейках со знаком «в нашу сторону»: > 0 — долг, < 0 — аванс. */
export interface SettlementRow {
  counterparty: string;
  contract?: string;
  contractSet?: string;
  document?: string;
  documentSet?: string;
  cents: number;
}

export interface DocInfo {
  /** YYYY-MM-DD */
  date: string;
  number?: string;
}

export type DebtStatus = "overdue" | "not_due" | "no_term" | "no_document" | "future_dated";

export interface AgedDocument {
  ref?: string;
  entitySet?: string;
  number?: string;
  date?: string;
  contract?: string;
  termDays?: number;
  termSource?: "contract" | "default";
  dueDate?: string;
  overdueDays?: number;
  ageDays?: number;
  cents: number;
  status: DebtStatus;
}

export const OVERDUE_BUCKETS = ["1-30", "31-60", "61-90", "90+"] as const;
/** Долг старше этого числа дней — подозрительно старый: часто незакрытые взаиморасчёты, а не реальный долг. */
export const STALE_DAYS = 365;
export type OverdueBucket = (typeof OVERDUE_BUCKETS)[number];
/** Возраст долга от даты документа — считается и там, где срок оплаты не задан. */
export const AGE_BUCKETS = ["0-30", "31-60", "61-90", "90+"] as const;
export type AgeBucket = (typeof AGE_BUCKETS)[number];

export interface CounterpartyAging {
  netCents: number;
  debtCents: number;
  advanceCents: number;
  overdueCents: number;
  notDueCents: number;
  noTermCents: number;
  noDocumentCents: number;
  /** Документы с датой позже asOf — аномалия учёта (ошибка в дате), не долг «в срок». */
  futureDatedCents: number;
  /** Долг по документам старше STALE_DAYS. */
  staleCents: number;
  oldestDate?: string;
  buckets: Record<OverdueBucket, number>;
  ageBuckets: Record<AgeBucket, number>;
  maxOverdueDays?: number;
  documents: AgedDocument[];
}

export const refOf = (v: unknown): string | undefined =>
  typeof v === "string" && v !== "" && v !== EMPTY_GUID ? v : undefined;

/** «StandardODATA.Document_РеализацияТоваровУслуг» → «Document_РеализацияТоваровУслуг». */
export function entitySetOfType(t: unknown): string | undefined {
  if (typeof t !== "string" || !t) return undefined;
  const name = t.slice(t.lastIndexOf(".") + 1);
  return name && name !== "Undefined" ? name : undefined;
}

const cents = (v: unknown): number => {
  const n = typeof v === "number" ? v : typeof v === "string" ? Number(v) : 0;
  return Number.isFinite(n) ? Math.round(n * 100) : 0;
};

/**
 * Строки сальдо (виртуальная таблица Balance) → строки расчётов. Субконто узнаём по типу,
 * а не по номеру: у 62/60 в БП это контрагент, договор и документ расчётов, но в другой
 * конфигурации порядок может отличаться. Строки без контрагента пропускаются.
 */
export function settlementRows(rows: readonly ODataEntity[], side: SettlementSide): SettlementRow[] {
  const out: SettlementRow[] = [];
  for (const r of rows) {
    let counterparty: string | undefined;
    let contract: string | undefined;
    let contractSet: string | undefined;
    let document: string | undefined;
    let documentSet: string | undefined;
    for (const i of [1, 2, 3]) {
      const ref = refOf(r[`ExtDimension${i}`]);
      const set = entitySetOfType(r[`ExtDimension${i}_Type`]);
      if (!ref || !set) continue;
      if (set === "Catalog_Контрагенты" && !counterparty) counterparty = ref;
      else if (set.startsWith("Catalog_Договор") && !contract) [contract, contractSet] = [ref, set];
      else if (set.startsWith("Document_") && !document) [document, documentSet] = [ref, set];
    }
    // Старые конфигурации отдают субконто без _Type: контрагент — первое субконто.
    counterparty ??= refOf(r["ExtDimension1"]);
    if (!counterparty) continue;
    const dr = cents(r["СуммаBalanceDr"]);
    const cr = cents(r["СуммаBalanceCr"]);
    out.push({
      counterparty,
      ...(contract ? { contract, contractSet } : {}),
      ...(document ? { document, documentSet } : {}),
      cents: side === "receivable" ? dr - cr : cr - dr,
    });
  }
  return out;
}

/** Целые дни между датами YYYY-MM-DD (b − a), без часовых поясов. */
export function daysBetween(a: string, b: string): number {
  const p = (s: string) => {
    const [y, m, d] = s.slice(0, 10).split("-").map(Number) as [number, number, number];
    return Date.UTC(y, m - 1, d);
  };
  return Math.round((p(b) - p(a)) / 86_400_000);
}

export function addDays(ymd: string, days: number): string {
  const [y, m, d] = ymd.slice(0, 10).split("-").map(Number) as [number, number, number];
  return new Date(Date.UTC(y, m - 1, d + days)).toISOString().slice(0, 10);
}

/** Сегодня по локальному календарю (без сдвига в UTC). */
export function localToday(now: Date = new Date()): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())}`;
}

export function ageBucket(days: number): AgeBucket {
  if (days <= 30) return "0-30";
  if (days <= 60) return "31-60";
  if (days <= 90) return "61-90";
  return "90+";
}

export function overdueBucket(days: number): OverdueBucket {
  if (days <= 30) return "1-30";
  if (days <= 60) return "31-60";
  if (days <= 90) return "61-90";
  return "90+";
}

/**
 * Возраст долга одного контрагента на дату asOf. Долг (cents > 0) раскладывается по
 * документам расчётов: срок = дата документа + дни из договора (или defaultTermDays, если
 * в договоре срок не задан и пользователь его назвал). Авансы (cents < 0) считаются
 * отдельно и в просрочку не идут. Без документа расчётов или без его даты возраст не
 * определить — такая сумма идёт в noDocument, а не угадывается.
 */
export function ageCounterparty(
  rows: readonly SettlementRow[],
  docs: ReadonlyMap<string, DocInfo>,
  terms: ReadonlyMap<string, number | undefined>,
  asOf: string,
  defaultTermDays?: number,
): CounterpartyAging {
  const res: CounterpartyAging = {
    netCents: 0,
    debtCents: 0,
    advanceCents: 0,
    overdueCents: 0,
    notDueCents: 0,
    noTermCents: 0,
    noDocumentCents: 0,
    futureDatedCents: 0,
    staleCents: 0,
    buckets: { "1-30": 0, "31-60": 0, "61-90": 0, "90+": 0 },
    ageBuckets: { "0-30": 0, "31-60": 0, "61-90": 0, "90+": 0 },
    documents: [],
  };
  for (const r of rows) {
    res.netCents += r.cents;
    if (r.cents < 0) {
      res.advanceCents -= r.cents;
      continue;
    }
    if (r.cents === 0) continue;
    res.debtCents += r.cents;
    const info = r.document ? docs.get(r.document) : undefined;
    const base = {
      ...(r.document ? { ref: r.document, entitySet: r.documentSet } : {}),
      ...(r.contract ? { contract: r.contract } : {}),
      cents: r.cents,
    };
    if (!info) {
      res.noDocumentCents += r.cents;
      res.documents.push({ ...base, status: "no_document" });
      continue;
    }
    const ageDays = daysBetween(info.date, asOf);
    if (ageDays < 0) {
      res.futureDatedCents += r.cents;
      res.documents.push({
        ...base,
        date: info.date,
        ...(info.number ? { number: info.number } : {}),
        ageDays,
        status: "future_dated",
      });
      continue;
    }
    res.ageBuckets[ageBucket(ageDays)] += r.cents;
    if (ageDays > STALE_DAYS) res.staleCents += r.cents;
    if (!res.oldestDate || info.date < res.oldestDate) res.oldestDate = info.date;
    const fromContract = r.contract ? terms.get(r.contract) : undefined;
    const termDays = fromContract ?? defaultTermDays;
    const doc = { ...base, date: info.date, ...(info.number ? { number: info.number } : {}), ageDays };
    if (termDays === undefined) {
      res.noTermCents += r.cents;
      res.documents.push({ ...doc, status: "no_term" });
      continue;
    }
    const dueDate = addDays(info.date, termDays);
    const overdueDays = daysBetween(dueDate, asOf);
    const termSource = fromContract !== undefined ? ("contract" as const) : ("default" as const);
    if (overdueDays > 0) {
      res.overdueCents += r.cents;
      res.buckets[overdueBucket(overdueDays)] += r.cents;
      res.maxOverdueDays = Math.max(res.maxOverdueDays ?? 0, overdueDays);
      res.documents.push({ ...doc, termDays, termSource, dueDate, overdueDays, status: "overdue" });
    } else {
      res.notDueCents += r.cents;
      res.documents.push({ ...doc, termDays, termSource, dueDate, status: "not_due" });
    }
  }
  // Сначала самая глубокая просрочка, затем самые старые документы.
  res.documents.sort(
    (a, b) =>
      (b.overdueDays ?? -Infinity) - (a.overdueDays ?? -Infinity) ||
      (b.ageDays ?? -Infinity) - (a.ageDays ?? -Infinity) ||
      b.cents - a.cents,
  );
  return res;
}

/**
 * Срок оплаты из договора в днях: только если в договоре стоит флаг «Установлен срок
 * оплаты». Без флага — undefined (срок не задан), даже если число дней заполнено.
 */
/**
 * Вид дней срока из договора, если он не календарный. Срок считается в календарных днях: рабочие
 * без производственного календаря не посчитать, поэтому такой договор только называется в ответе.
 * Значения поля на живых базах не встречались (оно пустое), поэтому сравнение только со словом
 * «календарный»; пустое считаем обычным сроком.
 */
export function nonCalendarTermKind(contract: ODataEntity | undefined): string | undefined {
  if (!contract || contract["УстановленСрокОплаты"] !== true) return undefined;
  const kind = String(contract["ВидДнейСрокаОплаты"] ?? "").trim();
  return kind && !/календ/i.test(kind) ? kind : undefined;
}

export function contractTermDays(contract: ODataEntity | undefined): number | undefined {
  if (!contract || contract["УстановленСрокОплаты"] !== true) return undefined;
  const n = Number(contract["СрокОплаты"]);
  return Number.isFinite(n) && n >= 0 ? Math.trunc(n) : undefined;
}
