import type { Connection } from "../context.js";
import { CATALOGS, REGISTERS, TAX_NOTICES, TAX_REPORTS, resolveEntity } from "../config/mapping.js";
import { accountingChart } from "../config/accounting.js";
import { readByRefs, num } from "./accounting.js";
import { fetchAll } from "./pagination.js";
import { describeObject, requireEntity } from "./publication.js";
import { and, cmp, odataGuid } from "./query.js";
import { taxRules, type TaxRule } from "./tax-calendar-rules.js";
import {
  taxRegimeReader,
  type TaxRegime,
  type TaxRegimeKey,
  type TaxRegimeResult,
  vatPayerFromSettings,
} from "./tax-regime.js";
import { turnoversPath } from "./item-sales.js";
import { readVatThreshold, type VatThresholdFact, type VatThresholdResult } from "./vat-threshold.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Налоговый календарь: что заплатить и сдать в окне дат и что из этого уже есть в 1С.
 *
 * Суммы и сроки начисленного — только из 1С (регистр начислений на ЕНС). Справочник добавляет то, чего в 1С ещё нет:
 * декларации, уведомления, будущие взносы и авансы — без сумм. «Оплачено» не бывает: единый налоговый платёж
 * распределяет ФНС (п. 8 ст. 45 НК РФ), в 1С этого нет.
 */

/** unknown — за этот период есть начисление, счёт которого не прочитан: этот ли это платёж, неизвестно. */
export type PaymentStatus = "accrued" | "not_accrued" | "unknown";
export type ReportStatus = "sent" | "prepared" | "not_found" | "unknown";
export type NoticeStatus = "prepared" | "not_required" | "not_found" | "unknown";

export interface CalendarItem {
  kind: "payment" | "report" | "notice";
  title: string;
  due: string;
  /** Откуда срок: из начисления 1С или из статьи НК. */
  dueSource: "1c" | "law";
  /** Срок по справочнику, если срок в 1С другой. */
  lawDue?: string;
  status: PaymentStatus | ReportStatus | NoticeStatus;
  amount?: number;
  account?: string;
  periodEnd?: string;
  basis?: string;
  condition?: string;
  /** Все документы отчёта за период: первичный и корректирующие, по дате подписи. */
  reports?: Array<{ signed?: string; sent?: string; corrective: boolean }>;
  notices?: Array<{ date: string; amount: number }>;
  note?: string;
}

export interface CalendarRegime {
  regimes: TaxRegimeKey[];
  source: "settings" | "inferred" | "user" | "unknown";
  asOf: string;
  system?: string;
  calculatedFor?: string;
}

export interface TaxCalendar {
  taxpayer: "ip" | "organization";
  from: string;
  to: string;
  regime: CalendarRegime;
  items: CalendarItem[];
  warnings: string[];
  notes: string[];
  /** Что не прочитано и почему календарь без этого неполон. */
  missing: string[];
  rulesCoverage: { years: number[]; missing: number[]; checkedAt: string };
  notCovered: string[];
  /** Доход по книгам 1С рядом с порогом НДС — для УСН, пока организация не плательщик НДС. */
  vatThreshold?: VatThresholdFact[];
}

export interface CalendarInput {
  orgRef: string;
  today: string;
  from: string;
  to: string;
  /** Режим со слов пользователя: заменяет определение по базе. */
  regime?: TaxRegimeKey[] | undefined;
}

/** Начисление и его сторно стоят на разных датах (перенесённой и нет) — читаем шире окна. */
const READ_MARGIN_DAYS = 120;
/** Сторно патента находит своё начисление на соседней дате: срок с переносом и без него. */
const REVERSAL_DAYS = 31;
const MAX_ROWS = 2000;
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

const text = (v: unknown): string => (typeof v === "string" ? v : "");
const date = (v: unknown): string | undefined => {
  const s = text(v);
  return !s || s.startsWith("0001-01-01") ? undefined : s.slice(0, 10);
};
const shift = (ymd: string, days: number) => {
  const d = new Date(`${ymd}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
};
const daysBetween = (a: string, b: string) => Math.abs(Date.parse(a) - Date.parse(b)) / 86_400_000;
const startOfDay = (ymd: string) => `datetime'${ymd}T00:00:00'`;
const endOfDay = (ymd: string) => `datetime'${ymd}T23:59:59'`;
const money = (n: number) => Math.round(n * 100) / 100;
const rub = (n: number) =>
  money(n)
    .toLocaleString("ru-RU")
    .replace(/\u00a0/g, " ");
const reason = (e: unknown) => (e instanceof Error ? e.message : String(e));

export interface Accrual {
  accountKey: string;
  due: string;
  periodEnd: string;
  amount: number;
  patent: boolean;
}

/**
 * Сворачивает записи регистра в «налог — срок — сумма». Сумма к сроку — «Начислено»: у патента «СуммаНалога» равна
 * всему налогу, а доли — в «Начислено». Сторно гасит начисление того же счёта и патента: на той же дате или, для
 * остатка, на соседней (начисление патента — с переносом с выходного, сторно уведомлением об уменьшении — без).
 * Без патента сворачиваются только строки одного периода и срока.
 */
export function netAccruals(rows: readonly ODataEntity[]): { due: Accrual[]; excess: Accrual[] } {
  const groups = new Map<string, Map<string, Accrual>>();
  for (const r of rows) {
    if (r["Active"] === false) continue;
    const accountKey = text(r["СчетУчета_Key"]);
    const due = date(r["СрокУплаты"]);
    const periodEnd = date(r["ПериодРасчета"]) ?? "";
    if (!due) continue;
    const patent = text(r["ДокументПатент_Key"]).replace(EMPTY_GUID, "");
    const group = `${accountKey}|${patent}`;
    // Разные периоды — разные обязанности, даже на одном счёте и сроке: налог за год и аванс за 1 квартал оба
    // платятся 28 апреля. У патента период строки сторно другой, поэтому внутри патента ключ — только срок.
    const line = patent ? due : `${due}|${periodEnd}`;
    const byDue = groups.get(group) ?? new Map<string, Accrual>();
    groups.set(group, byDue);
    const known = byDue.get(line);
    const amount = num(r["Начислено"]);
    if (!known) byDue.set(line, { accountKey, due, periodEnd, amount, patent: !!patent });
    else {
      // Период строки — по начислению, а не по сторно, которое записано более поздним периодом.
      if (amount > 0 && (known.amount <= 0 || periodEnd > known.periodEnd)) known.periodEnd = periodEnd;
      known.amount += amount;
    }
  }
  const out: { due: Accrual[]; excess: Accrual[] } = { due: [], excess: [] };
  for (const byDue of groups.values()) {
    const rest = [...byDue.values()]
      .map((a) => ({ ...a, amount: money(a.amount) }))
      .filter((a) => a.amount !== 0);
    // Соседнюю дату сторно ищет только внутри патента; у остальных налогов остаток показывается как есть.
    for (const minus of rest.filter((a) => a.amount < 0 && a.patent)) {
      const near = rest
        .filter((a) => a.amount > 0 && daysBetween(a.due, minus.due) <= REVERSAL_DAYS)
        .sort((a, b) => daysBetween(a.due, minus.due) - daysBetween(b.due, minus.due));
      for (const plus of near) {
        const take = Math.min(plus.amount, -minus.amount);
        plus.amount = money(plus.amount - take);
        minus.amount = money(minus.amount + take);
      }
    }
    out.due.push(...rest.filter((a) => a.amount > 0));
    out.excess.push(...rest.filter((a) => a.amount < 0));
  }
  return out;
}

interface Account {
  code: string;
  name: string;
}

interface Sources {
  conn: Connection;
  orgRef: string;
  missing: string[];
}

/** Необязательный источник: не опубликован или не читается — календарь без него, с объяснением. */
async function optional<T>(
  s: Sources,
  candidates: readonly string[],
  lacking: string,
  read: (set: string) => Promise<T>,
): Promise<T | undefined> {
  const set = resolveEntity(candidates, await s.conn.available());
  if (!set) {
    s.missing.push(`${lacking}: опубликуйте ${describeObject(candidates[0]!)} в «Составе OData».`);
    return undefined;
  }
  try {
    return await read(set);
  } catch (e) {
    s.missing.push(`${lacking}: ${describeObject(set)} прочитать не удалось (${reason(e)}).`);
    return undefined;
  }
}

async function readRows(s: Sources, set: string, filter: string, select: string[]): Promise<ODataEntity[]> {
  const { rows, truncated } = await fetchAll(
    s.conn.client,
    set,
    { filter: and(cmp("Организация_Key", "eq", odataGuid(s.orgRef)), filter), select },
    s.conn.behavior.pageSize,
    MAX_ROWS,
  );
  // Неполная выборка дала бы неверные суммы и «не найдено» там, где найдено.
  if (truncated) throw new Error(`строк больше ${MAX_ROWS}: сузьте период`);
  return rows;
}

function paymentWithoutAccrual(rule: TaxRule): string {
  if (rule.obligation === "ip_fixed_contributions")
    return "В 1С пока не начислено: фиксированные взносы 1С начисляет при закрытии декабря. Сумму возьмите из начисления.";
  if (rule.obligation === "ip_extra_contribution")
    return "В 1С не начислено. Коннектор этот взнос не считает.";
  return "В 1С не начислено: период ещё не закрыт или начислять нечего.";
}

/**
 * Была ли УСН в этот период, по ответу о режиме судить нельзя: режима нет вовсе; он выведен косвенно (отсутствие
 * следов УСН не значит, что её не было); или в настройках УСН без объекта налогообложения. Другая известная система
 * — не «неизвестно».
 */
function usnUnknown({ regime, noContributions }: TaxRegimeResult): boolean {
  if (!regime) return true;
  if (regime.regimes.some((k) => k.startsWith("usn"))) return false;
  return regime.source === "inferred" || (regime.system === "Упрощенная" && !noContributions);
}

const ruleItem = (rule: TaxRule, status: CalendarItem["status"]): CalendarItem => ({
  kind: rule.kind,
  title: rule.title,
  due: rule.due,
  dueSource: "law",
  status,
  periodEnd: rule.period.to,
  basis: rule.basis,
  ...(rule.condition ? { condition: rule.condition } : {}),
});

export async function buildTaxCalendar(conn: Connection, input: CalendarInput): Promise<TaxCalendar> {
  const { orgRef, today, from, to } = input;
  const accrualSet = await requireEntity(
    conn,
    REGISTERS.taxAccruals,
    "Начисления налогов на ЕНС со сроками уплаты",
  );
  const orgSet = await requireEntity(conn, CATALOGS.organizations, "Справочник «Организации»");
  const org = await conn.client.getEntity(`${orgSet}(guid'${orgRef}')?$format=json`);
  const taxpayer = text(org["ЮридическоеФизическоеЛицо"]) === "ФизическоеЛицо" ? "ip" : "organization";
  const registered = date(org["ДатаРегистрации"]);

  const s: Sources = { conn, orgRef, missing: [] };
  const warnings: string[] = [];
  const notes: string[] = [];
  const reference = taxRules();
  const readFrom = shift(from, -READ_MARGIN_DAYS);
  const readTo = shift(to, READ_MARGIN_DAYS);
  const inWindow = (due: string) => due >= from && due <= to;

  // ─── Начисления 1С ─────────────────────────────────────────────────────────
  const rows = await readRows(
    s,
    accrualSet,
    and(cmp("СрокУплаты", "ge", startOfDay(readFrom)), cmp("СрокУплаты", "le", endOfDay(readTo))),
    ["СрокУплаты", "ПериодРасчета", "СчетУчета_Key", "ДокументПатент_Key", "Начислено", "Active"],
  );
  const net = netAccruals(rows);
  const keys = [...net.due, ...net.excess].map((a) => a.accountKey);
  // undefined — план счетов не прочитан; это не то же самое, что «таких счетов нет».
  const chart = keys.length
    ? await optional(
        s,
        accountingChart(conn.cfg),
        "Названия налогов не прочитаны, начисления не сопоставить со сроками",
        (set) => readByRefs(conn, set, keys, ["Ref_Key", "Code", "Description"]),
      )
    : undefined;
  const accounts = new Map<string, Account>(
    [...(chart ?? [])].map(([ref, r]) => [ref, { code: text(r["Code"]), name: text(r["Description"]) }]),
  );
  const unidentified = [...net.due, ...net.excess].filter((a) => !accounts.get(a.accountKey)?.code);
  if (chart && unidentified.length)
    s.missing.push(
      `Названия ${accounts.size ? "части налогов" : "налогов"} не прочитаны: план счетов вернул не все счета — ` +
        "эти начисления не сопоставить со сроками.",
    );
  /** Периоды, за которые есть начисление неизвестно по какому счёту: «не начислено» о них сказать нельзя. */
  const unclear = new Set(unidentified.map((a) => a.periodEnd));

  // ─── Режим ────────────────────────────────────────────────────────────────
  // Справочные сроки — кандидаты в окне и рядом: срок начисления в 1С может чуть отличаться.
  const candidates = reference.rules.filter(
    (r) =>
      r.taxpayers.includes(taxpayer) &&
      r.due >= readFrom &&
      r.due <= readTo &&
      // До регистрации обязанностей не было.
      !(registered && r.period.to < registered),
  );
  let regimeAt: (asOf: string) => Promise<TaxRegimeResult>;
  if (input.regime) {
    const chosen: TaxRegime = { regimes: input.regime, source: "settings", asOf: today };
    regimeAt = () => Promise.resolve({ regime: chosen, notes: [] });
  } else {
    try {
      regimeAt = await taxRegimeReader(conn, orgRef, today);
    } catch (e) {
      s.missing.push(`Режим налогообложения прочитать не удалось (${reason(e)}).`);
      regimeAt = () => Promise.resolve({ notes: [] });
    }
  }
  const now = await regimeAt(today).catch((e: unknown): TaxRegimeResult => {
    s.missing.push(`Режим налогообложения прочитать не удалось (${reason(e)}).`);
    return { notes: [] };
  });
  notes.push(...now.notes);
  const regime: CalendarRegime = input.regime
    ? { regimes: input.regime, source: "user", asOf: today }
    : now.regime
      ? {
          regimes: now.regime.regimes,
          source: now.regime.source,
          asOf: today,
          ...(now.regime.system ? { system: now.regime.system } : {}),
          ...(now.regime.calculatedFor ? { calculatedFor: now.regime.calculatedFor } : {}),
        }
      : { regimes: [], source: "unknown", asOf: today };

  const applicable: TaxRule[] = [];
  /** Сроки УСН, про которые не известно, относятся ли они к организации: не показываются, но называются. */
  const undecided: TaxRule[] = [];
  for (const rule of candidates) {
    // Режим — тот, что был в периоде обязанности: декларацию за год сдаёт тот, кто был на УСН в этом году.
    const at = await regimeAt(rule.period.to < today ? rule.period.to : today).catch((): TaxRegimeResult => ({
      notes: [],
    }));
    // Взносы за себя ИП платит при любом режиме, кроме НПД и АУСН.
    const applies = rule.obligation.startsWith("ip_")
      ? !at.noContributions
      : rule.regimes.some((k) => (at.regime?.regimes ?? []).includes(k));
    if (applies) applicable.push(rule);
    else if (!rule.obligation.startsWith("ip_") && usnUnknown(at)) undecided.push(rule);
  }
  if (!regime.regimes.length)
    notes.push(
      regime.system === "Упрощенная" && !now.noContributions
        ? "В настройках 1С — УСН без объекта налогообложения: сроки УСН не показаны. Назовите режим параметром " +
            "regime (usn_income или usn_income_expense)."
        : regime.system
          ? `Режим по настройкам 1С — «${regime.system}»: сроков для него в справочнике нет, кроме взносов ИП за себя. ` +
            "Показано начисленное в 1С."
          : "Режим не определён — сроки УСН из справочника не показаны. Назовите режим параметром regime " +
            "(usn_income, usn_income_expense, psn, osno).",
    );

  // ─── Платежи ──────────────────────────────────────────────────────────────
  const items: CalendarItem[] = [];
  const accrued = new Set<string>();
  for (const a of [...net.due].sort((x, y) => x.due.localeCompare(y.due))) {
    const account = accounts.get(a.accountKey);
    // Начисление — факт: название и статья берутся из справочника, даже если режим не определён.
    const rule = reference.rules.find(
      (r) =>
        r.kind === "payment" &&
        !!account?.code &&
        r.taxpayers.includes(taxpayer) &&
        r.match?.account === account?.code &&
        r.match?.periodEnd === a.periodEnd,
    );
    if (rule) accrued.add(rule.id);
    if (!inWindow(a.due)) continue;
    const differs = !!rule && rule.due !== a.due;
    if (differs)
      warnings.push(
        `${rule.title}: в 1С срок ${a.due}, по справочнику ${rule.due} (${rule.basis}). Проверьте.`,
      );
    items.push({
      kind: "payment",
      title: rule?.title ?? `Начислено в 1С: ${account?.name || "налог без названия счёта"}`,
      due: a.due,
      dueSource: "1c",
      ...(differs ? { lawDue: rule.due } : {}),
      status: "accrued",
      amount: a.amount,
      ...(account?.code ? { account: account.code } : {}),
      periodEnd: a.periodEnd,
      ...(rule ? { basis: rule.basis } : {}),
    });
  }
  for (const a of net.excess.filter((x) => inWindow(x.due)))
    warnings.push(
      `${accounts.get(a.accountKey)?.name || "Налог"}: на ${a.due} в 1С отрицательный остаток начисления ` +
        `${rub(a.amount)} ₽ (уменьшение или сторно). Это не платёж.`,
    );
  for (const rule of applicable.filter((r) => r.kind === "payment" && !accrued.has(r.id) && inWindow(r.due)))
    items.push(
      unclear.has(rule.period.to)
        ? {
            ...ruleItem(rule, "unknown"),
            note: "За этот период в 1С есть начисление, но его счёт не прочитан: этот ли это платёж, неизвестно.",
          }
        : { ...ruleItem(rule, "not_accrued"), note: paymentWithoutAccrual(rule) },
    );

  // Режим не определён и на сегодня — об этом уже сказано выше. Начисленное в 1С показано и без режима.
  const unshown = undecided
    .filter((r) => inWindow(r.due) && !accrued.has(r.id))
    .sort((a, b) => a.due.localeCompare(b.due) || a.title.localeCompare(b.title, "ru"));
  if (unshown.length && (now.regime || input.regime))
    warnings.push(
      "По данным 1С нельзя определить режим УСН в периоде этих обязанностей, поэтому они не показаны — проверьте, " +
        `относятся ли они к вам: ${unshown.map((r) => `«${r.title}» — ${r.due}`).join("; ")}. ` +
        "Если режим известен, назовите его параметром regime.",
    );

  // ─── Отчёты ───────────────────────────────────────────────────────────────
  const reports = applicable.filter((r) => r.kind === "report" && inWindow(r.due));
  if (reports.length) {
    const ends = reports.map((r) => r.period.to).sort();
    const journal = await optional(s, TAX_REPORTS, "Статус отчётов неизвестен", (set) =>
      readRows(
        s,
        set,
        and(
          cmp("ДатаОкончания", "ge", startOfDay(ends[0]!)),
          cmp("ДатаОкончания", "le", endOfDay(ends.at(-1)!)),
          "DeletionMark eq false",
        ),
        ["ИсточникОтчета", "ДатаНачала", "ДатаОкончания", "Вид", "ДатаПодписи", "ДатаОтправки"],
      ),
    );
    for (const rule of reports) {
      const found = (journal ?? []).filter(
        (r) =>
          r["ИсточникОтчета"] === rule.match?.reportSource &&
          date(r["ДатаОкончания"]) === rule.period.to &&
          date(r["ДатаНачала"]) === rule.period.from,
      );
      const documents = found
        .map((r) => {
          const [signed, sent] = [date(r["ДатаПодписи"]), date(r["ДатаОтправки"])];
          return { ...(signed ? { signed } : {}), ...(sent ? { sent } : {}), corrective: num(r["Вид"]) > 0 };
        })
        .sort(
          (a, b) =>
            Number(a.corrective) - Number(b.corrective) || (a.signed ?? "").localeCompare(b.signed ?? ""),
        );
      const status: ReportStatus = !journal
        ? "unknown"
        : !documents.length
          ? "not_found"
          : documents.some((d) => "sent" in d)
            ? "sent"
            : "prepared";
      items.push({ ...ruleItem(rule, status), ...(documents.length ? { reports: documents } : {}) });
    }
    if (reports.length && journal)
      notes.push(
        "prepared — отчёт есть в 1С, но даты отправки нет: отправленный не из 1С так и выглядит. Это не «не сдан».",
      );
  }

  // ─── Уведомления ──────────────────────────────────────────────────────────
  const notices = applicable.filter((r) => r.kind === "notice" && inWindow(r.due));
  if (notices.length) {
    const periods = notices.map((r) => r.period.to).sort();
    const documents = await optional(s, TAX_NOTICES, "Уведомления не проверены", (set) =>
      readRows(
        s,
        set,
        // Уведомление за квартал подают после его конца; запас — на поданные заранее.
        and(cmp("Date", "ge", startOfDay(shift(periods[0]!, -REVERSAL_DAYS * 3))), "DeletionMark eq false"),
        ["Date", "Налоги"],
      ),
    );
    const calculations = await optional(
      s,
      REGISTERS.usnCalculation,
      "Расчёт УСН по кварталам не прочитан",
      (set) =>
        readRows(
          s,
          set,
          and(
            cmp("ПериодРасчета", "ge", startOfDay(periods[0]!)),
            cmp("ПериодРасчета", "le", endOfDay(periods.at(-1)!)),
          ),
          ["ПериодРасчета", "НалогКУплате"],
        ),
    );
    for (const rule of notices) {
      const filed = (documents ?? []).flatMap((d) =>
        (Array.isArray(d["Налоги"]) ? (d["Налоги"] as ODataEntity[]) : [])
          .filter(
            (l) => rule.match?.kbk?.includes(text(l["КодБК"])) && date(l["КонецПериода"]) === rule.period.to,
          )
          .map((l) => ({ date: date(d["Date"]) ?? "", amount: money(num(l["Сумма"])) })),
      );
      const calculated = (calculations ?? []).filter((c) => date(c["ПериодРасчета"]) === rule.period.to);
      const toPay = calculated.reduce((sum, c) => sum + num(c["НалогКУплате"]), 0);
      const account = rule.match?.kbk ? "68.12" : undefined;
      const accrual = money(
        net.due
          .filter((a) => accounts.get(a.accountKey)?.code === account && a.periodEnd === rule.period.to)
          .reduce((sum, a) => sum + a.amount, 0),
      );
      const status: NoticeStatus = filed.length
        ? "prepared"
        : !documents
          ? "unknown"
          : accrual > 0 || toPay > 0
            ? "not_found"
            : calculated.length && !unclear.has(rule.period.to)
              ? "not_required"
              : "unknown";
      // Уведомления прочитаны, а ответа всё равно нет — называется причина.
      const why = unclear.has(rule.period.to)
        ? "За этот квартал есть начисление, но его счёт не прочитан: нужно ли уведомление, неизвестно."
        : !calculations
          ? "Расчёт УСН не прочитан: нужно ли уведомление, неизвестно."
          : "Квартал в 1С не рассчитан: нужно ли уведомление, станет ясно после закрытия квартала.";
      items.push({
        ...ruleItem(rule, status),
        ...(filed.length ? { notices: filed.sort((a, b) => a.date.localeCompare(b.date)) } : {}),
        ...(status === "unknown" && documents ? { note: why } : {}),
      });
      // Одно из уведомлений или их сумма должны совпасть с начислением; какое действует, коннектор не решает.
      const total = filed.reduce((sum, n) => sum + n.amount, 0);
      if (
        accrual > 0 &&
        filed.length &&
        ![...filed.map((n) => n.amount), total].some((n) => Math.abs(n - accrual) < 1)
      )
        warnings.push(
          `${rule.title}: в 1С начислено ${rub(accrual)} ₽, а в уведомлениях — ` +
            `${filed.map((n) => `${rub(n.amount)} ₽ от ${n.date}`).join(", ")}. Проверьте, какая сумма верна.`,
        );
    }
  }

  // ─── Порог НДС на УСН ─────────────────────────────────────────────────────
  let vat: VatThresholdResult | undefined;
  const onUsn = regime.regimes.some((k) => k.startsWith("usn"));
  let vatPayer = now.regime?.vatPayerOnUsn;
  if (onUsn && input.regime)
    vatPayer = await vatPayerFromSettings(conn, orgRef, today).catch((e: unknown) => {
      s.missing.push(`Из настроек 1С не прочитано, плательщик ли НДС (${reason(e)}).`);
      return undefined;
    });
  if (onUsn && vatPayer !== true) {
    vat = await readVatThreshold({
      taxpayer,
      today,
      registered,
      vatPayer,
      thresholds: reference.vatThreshold,
      turnover: (book, since, until) =>
        optional(
          s,
          book === "usn" ? REGISTERS.usnIncomeBook : REGISTERS.patentIncomeBook,
          book === "usn"
            ? "Порог НДС не проверен"
            : "Доход по патенту в порог НДС не прибавлен (нужен, если патент применялся)",
          async (set) => {
            const path = turnoversPath(set, since, until, cmp("Организация_Key", "eq", odataGuid(orgRef)));
            // Графа 4 нужна ради строки: по одной графе 5 при нулевом доходе 1С не вернула бы ничего.
            const select = book === "usn" ? ["Графа4Turnover", "Графа5Turnover"] : ["Графа4Turnover"];
            return (await fetchAll(s.conn.client, path, { select }, MAX_ROWS, MAX_ROWS)).rows;
          },
        ),
    });
    warnings.push(...vat.warnings);
    notes.push(...vat.notes);
  }

  const years: number[] = [];
  for (let y = Number(from.slice(0, 4)); y <= Number(to.slice(0, 4)); y++) years.push(y);
  const uncovered = years.filter((y) => !reference.dueYears.includes(y));
  if (uncovered.length)
    warnings.push(
      `В справочнике нет сроков на ${uncovered.join(", ")} год: за него показано только начисленное в 1С. ` +
        "Обновите коннектор.",
    );
  notes.push(
    "Оплату календарь не показывает: единый налоговый платёж распределяет ФНС. Уплаченное из банка — get_taxes_paid.",
  );

  return {
    taxpayer,
    from,
    to,
    regime,
    items: items.sort((a, b) => a.due.localeCompare(b.due) || a.title.localeCompare(b.title, "ru")),
    warnings,
    notes,
    // Одна и та же книга нужна за два года: причина называется один раз.
    missing: [...new Set(s.missing)],
    rulesCoverage: {
      years: years.filter((y) => reference.dueYears.includes(y)),
      missing: uncovered,
      checkedAt: reference.checkedAt,
    },
    notCovered: reference.notCovered,
    ...(vat?.facts.length ? { vatThreshold: vat.facts } : {}),
  };
}
