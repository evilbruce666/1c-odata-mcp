import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, guard, databaseField, organizationField } from "./_shared.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { accountPrefixes, defaultAccountCodes } from "../config/accounting.js";
import { CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import {
  balanceByAccounts,
  nextDay,
  num,
  readByRefs,
  resolveAccounts,
  resolveNames,
  resolvePostingsSource,
  type Account,
} from "../odata/accounting.js";
import { collectDocuments, fetchAllForAggregation } from "../odata/aggregate.js";
import { listOrganizations, resolveOrganization } from "../odata/orgs.js";
import { addDays, entitySetOfType, refOf } from "../odata/settlements-aging.js";
import { monthCloseResultSchema } from "../schemas/output.js";
import { InputError } from "../errors.js";
import { createdAt, documentTitle, negativeCashHint } from "../odata/names.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Проверка перед закрытием месяца — то, что бухгалтер смотрит руками: выполнено ли закрытие,
 * нет ли непроведённых документов, «красного» сальдо, остатков на закрываемых счетах, пропущенной
 * амортизации и реализаций с НДС без счёта-фактуры. Только чтение: регламентные операции через
 * OData не запускаются, мы читаем их результат.
 */

export type CheckStatus = "ok" | "warning" | "error" | "not_applicable" | "skipped";

export interface CheckResult {
  id: string;
  title: string;
  status: CheckStatus;
  message: string;
  items?: Record<string, unknown>[];
  itemsMore?: number;
}

const ITEMS_LIMIT = 20;
const rub = (c: number): number => Math.round(c) / 100;
const toCents = (v: unknown): number => Math.round(num(v) * 100);

const SEVERITY: Record<CheckStatus, number> = { error: 3, warning: 2, ok: 1, skipped: 0, not_applicable: 0 };

export function overallStatus(checks: readonly CheckResult[]): "ok" | "warnings" | "errors" {
  const worst = Math.max(0, ...checks.map((c) => SEVERITY[c.status]));
  return worst === 3 ? "errors" : worst === 2 ? "warnings" : "ok";
}

function withItems(items: Record<string, unknown>[]): Pick<CheckResult, "items" | "itemsMore"> {
  return {
    ...(items.length ? { items: items.slice(0, ITEMS_LIMIT) } : {}),
    ...(items.length > ITEMS_LIMIT ? { itemsMore: items.length - ITEMS_LIMIT } : {}),
  };
}

/** «2026-09» → границы месяца. */
export function monthBounds(month: string): { from: string; to: string } {
  const [y, m] = month.split("-").map(Number) as [number, number];
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return { from: `${month}-01`, to: `${month}-${String(last).padStart(2, "0")}` };
}

// ─── 1. Закрытие месяца ──────────────────────────────────────────────────────

const CLOSING_SET = ["Document_РегламентнаяОперация"];
/** Без этой операции месяц в БП не закрыт: она формирует финансовый результат. */
export const KEY_CLOSING_OPERATION = "ЗакрытиеСчетов90_91";
const DONE_STATES = new Set(["Выполнено", "Пропущено"]);

export interface ClosingOp {
  month: string;
  org?: string;
  ref?: string;
  kind: string;
  state: string;
}

/** Состояние закрытия одного месяца по регламентным операциям (только для одной организации). */
export function evaluateClosing(ops: readonly ClosingOp[]): {
  closed: boolean;
  failed: ClosingOp[];
  kinds: string[];
} {
  const failed = ops.filter((o) => !DONE_STATES.has(o.state));
  const key = ops.find((o) => o.kind === KEY_CLOSING_OPERATION);
  return {
    closed: !!key && key.state === "Выполнено" && failed.length === 0,
    failed,
    kinds: ops.map((o) => o.kind),
  };
}

async function checkClosing(
  conn: Connection,
  month: string,
  orgs: ReadonlyMap<string, string>,
  orgKey: string | undefined,
): Promise<{
  result: CheckResult;
  earlier: CheckResult;
  closedOrgs: Set<string>;
  closedAt: Map<string, string>;
}> {
  const title = "Закрытие месяца";
  const earlierTitle = "Предыдущие месяцы года";
  const set = resolveEntity(CLOSING_SET, await conn.available());
  if (!set) {
    const skipped = (t: string, id: string): CheckResult => ({
      id,
      title: t,
      status: "skipped",
      message: "Документ «Регламентная операция» не опубликован в OData — закрытие месяца не проверить.",
    });
    return {
      result: skipped(title, "closing"),
      earlier: skipped(earlierTitle, "earlierMonths"),
      closedOrgs: new Set(),
      closedAt: new Map(),
    };
  }
  const { from: yearStart } = monthBounds(`${month.slice(0, 4)}-01`);
  const { to } = monthBounds(month);
  const { rows } = await collectDocuments(conn, set, {
    baseFilter:
      and(
        orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
        cmp(DOC_FIELDS.deletionMark, "eq", "false"),
      ) || undefined,
    dateField: DOC_FIELDS.date,
    from: yearStart,
    to,
    select: [DOC_FIELDS.ref, DOC_FIELDS.date, DOC_FIELDS.organization, "ВидОперации", "Состояние"],
  });
  const ops: ClosingOp[] = rows.map((r) => ({
    month: String(r[DOC_FIELDS.date] ?? "").slice(0, 7),
    ...(refOf(r[DOC_FIELDS.organization]) ? { org: refOf(r[DOC_FIELDS.organization]) } : {}),
    kind: String(r["ВидОперации"] ?? ""),
    state: String(r["Состояние"] ?? ""),
    ...(refOf(r[DOC_FIELDS.ref]) ? { ref: refOf(r[DOC_FIELDS.ref]) } : {}),
  }));

  // Организации, которые закрывают месяцы: выбранная или все, у кого были закрытия в году.
  const withOps = [...new Set(ops.map((o) => o.org).filter((o): o is string => !!o))];
  // Закрытий в году нет совсем — проверяем все организации базы: у каждой месяц не закрыт.
  const orgRefs = orgKey ? [orgKey] : withOps.length ? withOps : [...orgs.keys()];
  const orgName = (ref: string) => orgs.get(ref) ?? ref;
  const closedOrgs = new Set<string>();
  const closedAt = new Map<string, string>();
  const notClosed: Record<string, unknown>[] = [];
  const failedItems: Record<string, unknown>[] = [];
  const earlierItems: Record<string, unknown>[] = [];
  const doneKinds: Record<string, unknown>[] = [];

  for (const org of orgRefs) {
    const forOrg = ops.filter((o) => o.org === org || (!o.org && orgRefs.length === 1));
    const cur = evaluateClosing(forOrg.filter((o) => o.month === month));
    if (cur.closed) closedOrgs.add(org);
    else notClosed.push({ organization: orgName(org) });
    for (const f of cur.failed)
      failedItems.push({ organization: orgName(org), operation: f.kind, state: f.state });
    if (cur.kinds.length) {
      const key = forOrg.find((o) => o.month === month && o.kind === KEY_CLOSING_OPERATION);
      const at = createdAt(key?.ref);
      doneKinds.push({
        organization: orgName(org),
        operations: cur.kinds,
        // Время ПЕРВОГО создания операции закрытия (по ссылке): документы, созданные позже, в
        // закрытие могли не попасть. Перезакрытие ссылку не меняет — время остаётся прежним.
        ...(at ? { closingCreatedAt: at } : {}),
      });
      if (at) closedAt.set(org, at);
    }
    for (let m = 1; m < Number(month.slice(5, 7)); m++) {
      const mm = `${month.slice(0, 4)}-${String(m).padStart(2, "0")}`;
      if (!evaluateClosing(forOrg.filter((o) => o.month === mm)).closed) {
        earlierItems.push({ organization: orgName(org), month: mm });
      }
    }
  }

  let result: CheckResult;
  if (ops.length === 0) {
    result = {
      id: "closing",
      title,
      status: "error",
      message: `В ${month.slice(0, 4)} году нет ни одной регламентной операции — закрытие месяца не выполнялось.`,
      ...withItems(notClosed),
    };
  } else if (failedItems.length) {
    result = {
      id: "closing",
      title,
      status: "error",
      message:
        "Закрытие месяца выполнено с ошибками — откройте обработку «Закрытие месяца» и разберите ошибки.",
      ...withItems(failedItems),
    };
  } else if (notClosed.length) {
    result = {
      id: "closing",
      title,
      status: "error",
      message: `Месяц не закрыт: нет выполненной операции «Закрытие счетов 90, 91».`,
      ...withItems([...notClosed, ...doneKinds]),
    };
  } else {
    result = {
      id: "closing",
      title,
      status: "ok",
      message: "Закрытие месяца выполнено.",
      ...withItems(doneKinds),
    };
  }
  const earlier: CheckResult = earlierItems.length
    ? {
        id: "earlierMonths",
        title: earlierTitle,
        status: "warning",
        message:
          "Есть незакрытые месяцы раньше проверяемого. Закрытие идёт по порядку: пока они не закрыты, " +
          "результат этого месяца неточен.",
        ...withItems(earlierItems),
      }
    : {
        id: "earlierMonths",
        title: earlierTitle,
        status: Number(month.slice(5, 7)) === 1 ? "not_applicable" : "ok",
        message:
          Number(month.slice(5, 7)) === 1
            ? "Январь — предыдущих месяцев в году нет."
            : "Все предыдущие месяцы года закрыты.",
      };
  return { result, earlier, closedOrgs, closedAt };
}

// ─── 2. Непроведённые документы ──────────────────────────────────────────────

/** Документы, которые в БП 3.0 проводятся. Счета на оплату и платёжные поручения не проводятся. */
const POSTABLE_DOCUMENTS: readonly (readonly string[])[] = [
  DOCUMENTS.sales,
  DOCUMENTS.purchases,
  DOCUMENTS.bankIn,
  DOCUMENTS.bankOut,
  DOCUMENTS.cashIn,
  DOCUMENTS.cashOut,
  DOCUMENTS.issuedInvoice,
  DOCUMENTS.receivedInvoice,
  DOCUMENTS.servicesAct,
  DOCUMENTS.returnFromCustomer,
  DOCUMENTS.returnToSupplier,
  DOCUMENTS.transfer,
  DOCUMENTS.surplus,
  DOCUMENTS.writeoff,
  ["Document_АвансовыйОтчет"],
  ["Document_ОперацияБух"],
  ["Document_КорректировкаДолга"],
  ["Document_ТребованиеНакладная"],
  ["Document_ОтчетОРозничныхПродажах"],
  ["Document_ПринятиеКУчетуОС"],
  ["Document_ОтражениеЗарплатыВУчете"],
];

async function checkUnposted(
  conn: Connection,
  bounds: { from: string; to: string },
  orgKey: string | undefined,
): Promise<CheckResult> {
  const available = await conn.available();
  const meta = await conn.getMetadata();
  const items: Record<string, unknown>[] = [];
  const cpRefs: string[] = [];
  const sets = [
    ...new Set(POSTABLE_DOCUMENTS.map((c) => resolveEntity(c, available)).filter((s): s is string => !!s)),
  ];
  const perSet = await Promise.all(
    sets.map(async (set) => {
      const props = new Set(meta.entities.get(set)?.properties.map((p) => p.name) ?? []);
      const select: string[] = [DOC_FIELDS.ref, DOC_FIELDS.number, DOC_FIELDS.date];
      for (const f of [DOC_FIELDS.amount, DOC_FIELDS.counterparty]) if (props.has(f)) select.push(f);
      const { rows } = await collectDocuments(conn, set, {
        baseFilter:
          and(
            orgKey && props.has(DOC_FIELDS.organization)
              ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey))
              : undefined,
            cmp(DOC_FIELDS.posted, "eq", "false"),
            cmp(DOC_FIELDS.deletionMark, "eq", "false"),
          ) || undefined,
        dateField: DOC_FIELDS.date,
        ...bounds,
        select,
      });
      return { set, rows };
    }),
  );
  for (const { set, rows } of perSet) {
    for (const r of rows) {
      const cp = refOf(r[DOC_FIELDS.counterparty]);
      if (cp) cpRefs.push(cp);
      items.push({
        document: set.replace(/^Document_/, ""),
        documentTitle: documentTitle(set),
        number: String(r[DOC_FIELDS.number] ?? "").trim(),
        date: String(r[DOC_FIELDS.date] ?? "").slice(0, 10),
        ref: String(r[DOC_FIELDS.ref]),
        ...(r[DOC_FIELDS.amount] !== undefined ? { amount: num(r[DOC_FIELDS.amount]) } : {}),
        ...(cp ? { counterpartyRef: cp } : {}),
      });
    }
  }
  const cpSet = resolveEntity(CATALOGS.counterparties, available);
  const names = cpSet ? await resolveNames(conn, cpSet, cpRefs.slice(0, 200)) : new Map<string, string>();
  for (const i of items) {
    const ref = i["counterpartyRef"];
    if (typeof ref === "string" && names.get(ref)) i["counterparty"] = names.get(ref);
  }
  items.sort((a, b) => String(a["date"]).localeCompare(String(b["date"])));
  return items.length
    ? {
        id: "unposted",
        title: "Непроведённые документы",
        status: "warning",
        message:
          `${items.length} документ(ов) за месяц не проведены — их суммы не попадают в учёт и закрытие. ` +
          "Проведите или пометьте на удаление.",
        ...withItems(items),
      }
    : {
        id: "unposted",
        title: "Непроведённые документы",
        status: "ok",
        message: "Все документы месяца проведены.",
      };
}

// ─── 3. Красное сальдо ───────────────────────────────────────────────────────

type Side = "debit" | "credit";

/**
 * Строки с сальдо «не на той стороне»: expected — нормальная сторона счёта. Сумма копится по
 * ключу (счёт + аналитика), чтобы одна позиция не дробилась по документам.
 */
export function wrongSideBalances(
  rows: readonly ODataEntity[],
  expected: (accountKey: string) => Side | undefined,
  keyDims: number,
  withQuantity = false,
): Array<{ accountKey: string; dims: string[]; cents: number; quantity?: number }> {
  const acc = new Map<string, { accountKey: string; dims: string[]; cents: number; qty: number }>();
  for (const r of rows) {
    const accountKey = String(r["Account_Key"] ?? "");
    if (!expected(accountKey)) continue;
    const dims = [1, 2, 3].slice(0, keyDims).map((i) => refOf(r[`ExtDimension${i}`]) ?? "");
    const k = [accountKey, ...dims].join("|");
    const cur = acc.get(k) ?? { accountKey, dims, cents: 0, qty: 0 };
    cur.cents += toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]);
    cur.qty += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
    acc.set(k, cur);
  }
  const out: Array<{ accountKey: string; dims: string[]; cents: number; quantity?: number }> = [];
  for (const v of acc.values()) {
    const side = expected(v.accountKey)!;
    const wrongSum = side === "debit" ? v.cents < 0 : v.cents > 0;
    const wrongQty = withQuantity && v.qty < -0.0005;
    if (!wrongSum && !wrongQty) continue;
    out.push({
      accountKey: v.accountKey,
      dims: v.dims,
      cents: v.cents,
      ...(withQuantity ? { quantity: Math.round(v.qty * 1000) / 1000 } : {}),
    });
  }
  return out.sort((a, b) => Math.abs(b.cents) - Math.abs(a.cents));
}

async function namesForDims(
  conn: Connection,
  rows: readonly ODataEntity[],
  refs: Iterable<string>,
): Promise<{ names: Map<string, string>; types: Map<string, string> }> {
  const want = new Set(refs);
  const bySet = new Map<string, Set<string>>();
  const types = new Map<string, string>();
  for (const r of rows) {
    for (const i of [1, 2, 3]) {
      const ref = refOf(r[`ExtDimension${i}`]);
      const set = entitySetOfType(r[`ExtDimension${i}_Type`]);
      if (!ref || !set || !want.has(ref)) continue;
      types.set(ref, set);
      if (set.startsWith("Catalog_")) bySet.set(set, (bySet.get(set) ?? new Set()).add(ref));
    }
  }
  const names = new Map<string, string>();
  for (const [set, rs] of bySet) for (const [k, v] of await resolveNames(conn, set, rs)) names.set(k, v);
  return { names, types };
}

/** План счетов и сальдо на конец месяца — один раз на все проверки остатков. */
export interface BalanceSnapshot {
  byKey: ReadonlyMap<string, Account>;
  rows: readonly ODataEntity[];
}

const startsWithAny = (code: string, prefixes: readonly string[]) => prefixes.some((p) => code.startsWith(p));

/**
 * Счета из настройки плана счетов. При strictDefaults без нужной группы настройка бросает
 * InputError — тогда эта часть проверки пропускается с причиной, а не валит всю проверку месяца.
 */
function configured(label: string, get: () => readonly string[], missing: string[]): readonly string[] {
  try {
    return get();
  } catch (e) {
    if (!(e instanceof InputError)) throw e;
    missing.push(label);
    return [];
  }
}

const notConfigured = (missing: readonly string[]) =>
  `Не проверено: в настройке плана счетов (strictDefaults) нет ${missing.join(", ")} — задайте их в ` +
  "ODATA_ACCOUNTING_CONFIG.";

async function balanceSnapshot(
  conn: Connection,
  asOf: string,
  orgKey: string | undefined,
): Promise<BalanceSnapshot> {
  const ignore: string[] = [];
  const prefixes = [
    ...(["cash", "inventory", "receivables", "payables"] as const).flatMap((g) =>
      configured(`prefixes.${g}`, () => accountPrefixes(conn.cfg, g), ignore),
    ),
    ...CLOSED_ACCOUNTS.map((c) => c.code),
    ...ASSET_ACCOUNTS,
  ];
  const [chart, rows] = await Promise.all([
    resolveAccounts(conn, [...new Set(prefixes)]),
    // Без отбора по счетам: один снимок сальдо вместо запроса на каждую группу счетов.
    balanceByAccounts(conn, null, orgKey, asOf),
  ]);
  return { byKey: new Map(chart.map((a) => [a.key, a])), rows };
}

async function checkRedBalances(conn: Connection, snap: BalanceSnapshot): Promise<CheckResult> {
  const title = "Красное сальдо";
  const { byKey } = snap;
  const group = (prefixes: readonly string[]) =>
    new Set([...byKey.values()].filter((a) => startsWithAny(a.code, prefixes)).map((a) => a.key));
  const missing: string[] = [];
  const cash = group(configured("prefixes.cash", () => accountPrefixes(conn.cfg, "cash"), missing));
  const stock = group(
    configured("prefixes.inventory", () => accountPrefixes(conn.cfg, "inventory"), missing),
  );
  // Расчёты: основной счёт и счёт авансов из настройки плана счетов (62.01/62.02, 60.01/60.02).
  const codes = (g: Parameters<typeof defaultAccountCodes>[1]) =>
    configured(`defaults.${g}`, () => defaultAccountCodes(conn.cfg, g), missing);
  const debitSettl = new Set([...codes("receivable"), ...codes("payableAdvance")]);
  const creditSettl = new Set([...codes("receivableAdvance"), ...codes("payable")]);
  const settlementSide = (k: string): Side | undefined => {
    const code = byKey.get(k)?.code ?? "";
    return debitSettl.has(code) ? "debit" : creditSettl.has(code) ? "credit" : undefined;
  };
  const rowsOf = (keys: ReadonlySet<string>) =>
    snap.rows.filter((r) => keys.has(String(r["Account_Key"] ?? "")));
  const cashRows = rowsOf(cash);
  const stockRows = rowsOf(stock);
  const settlRows = snap.rows.filter((r) => settlementSide(String(r["Account_Key"] ?? "")));

  if (!cash.size && !stock.size && !debitSettl.size && !creditSettl.size && missing.length) {
    return { id: "redBalances", title, status: "skipped", message: notConfigured(missing) };
  }
  const notChecked = missing.length ? ` ${notConfigured(missing)}` : "";
  const red = {
    cash: wrongSideBalances(cashRows, () => "debit", 1),
    stock: wrongSideBalances(stockRows, () => "debit", 2, true),
    settl: wrongSideBalances(settlRows, settlementSide, 2),
  };
  const allRows = [...cashRows, ...stockRows, ...settlRows];
  const { names, types } = await namesForDims(
    conn,
    allRows,
    [...red.cash, ...red.stock, ...red.settl].flatMap((r) => r.dims),
  );
  const label = (dims: string[]) => dims.filter(Boolean).map((d) => names.get(d) ?? d);
  const item = (kind: string, r: (typeof red.cash)[number], hint?: string) => {
    const a = byKey.get(r.accountKey)!;
    const refs = r.dims.filter(Boolean);
    return {
      kind,
      account: a.code,
      ...(refs.length
        ? {
            analytics: label(r.dims),
            analyticsRefs: refs,
            analyticsTypes: refs.map((d) => types.get(d) ?? ""),
          }
        : {}),
      balance: rub(r.cents),
      ...(r.quantity !== undefined && r.quantity < 0 ? { quantity: r.quantity } : {}),
      ...(hint ? { hint } : {}),
    };
  };
  // Минус у строки без банковского счёта / кассы — суммы не разнесены, это не «денег не хватает».
  const cashNoAnalytics = red.cash.filter((r) => !r.dims.some(Boolean));
  const cashNegative = red.cash.filter((r) => r.dims.some(Boolean));
  const code = (r: (typeof red.cash)[number]) => byKey.get(r.accountKey)!.code;
  const errors = [
    ...cashNegative.map((r) => item("деньги в минусе", r, negativeCashHint(code(r), false))),
    ...red.stock.map((r) => item("отрицательный остаток", r)),
  ];
  const warnings = [
    ...cashNoAnalytics.map((r) => item("минус без аналитики", r, negativeCashHint(code(r), true))),
    ...red.settl.map((r) => item("расчёты не на той стороне", r)),
  ];
  if (!errors.length && !warnings.length) {
    return {
      id: "redBalances",
      title,
      status: "ok",
      message: `Красного сальдо по деньгам, товарам и расчётам нет.${notChecked}`,
    };
  }
  const parts = [
    cashNegative.length ? `денег в минусе — ${cashNegative.length}` : "",
    cashNoAnalytics.length ? `минус по деньгам без аналитики — ${cashNoAnalytics.length}` : "",
    red.stock.length ? `отрицательных остатков товаров — ${red.stock.length}` : "",
    red.settl.length ? `расчётов «не на той стороне» — ${red.settl.length}` : "",
  ].filter(Boolean);
  return {
    id: "redBalances",
    title,
    status: errors.length ? "error" : "warning",
    message:
      `На конец месяца: ${parts.join(", ")}. Деньги и товары в минусе — ошибка учёта (не загружена ` +
      "выписка, расход раньше прихода). Расчёты не на той стороне (кредит 62.01, дебет 60.01 и т.п.) — " +
      `обычно незачтённый аванс или оплата не по тому договору.${notChecked}`,
    ...withItems([...errors, ...warnings]),
  };
}

/** Чистое сальдо (Дт − Кт, копейки) по группам счетов: «90», «01» и т.п. */
export function netByTopCode(
  snap: BalanceSnapshot,
  codes: readonly string[],
  exclude: (a: Account) => boolean = () => false,
): Map<string, number> {
  const out = new Map<string, number>();
  for (const r of snap.rows) {
    const a = snap.byKey.get(String(r["Account_Key"] ?? ""));
    if (!a || exclude(a)) continue;
    const top = a.code.split(".")[0]!;
    if (!codes.includes(top)) continue;
    out.set(top, (out.get(top) ?? 0) + toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]));
  }
  return out;
}

// ─── 4. Остатки на закрываемых счетах ─────────────────────────────────────────

/** Должны быть в нуле после закрытия (strict) или обычно в нуле (иначе — незавершённое производство). */
const CLOSED_ACCOUNTS: Array<{ code: string; strict: boolean }> = [
  { code: "90", strict: true },
  { code: "91", strict: true },
  { code: "25", strict: false },
  { code: "26", strict: false },
  { code: "20", strict: false },
  { code: "23", strict: false },
  { code: "44", strict: false },
];

function checkAfterClose(snap: BalanceSnapshot, closed: boolean): CheckResult {
  const title = "Остатки после закрытия";
  if (!closed) {
    return {
      id: "afterClose",
      title,
      status: "skipped",
      message: "Месяц не закрыт — остатки на 90, 91, 20–26, 44 проверяются после закрытия.",
    };
  }
  const items: Record<string, unknown>[] = [];
  let strictFail = false;
  const net = netByTopCode(
    snap,
    CLOSED_ACCOUNTS.map((c) => c.code),
  );
  for (const { code, strict } of CLOSED_ACCOUNTS) {
    const cents = net.get(code) ?? 0;
    if (cents === 0) continue;
    strictFail ||= strict;
    items.push({
      account: code,
      balance: rub(cents),
      severity: strict ? "error" : "warning",
      reason: strict
        ? "после закрытия месяца счёт в целом должен быть в нуле. Обычно это документы, внесённые или перепроведённые после закрытия, — карточка счёта за месяц покажет какие; затем перезакройте месяц"
        : code === "44"
          ? "остаток бывает законным (транспортные расходы на остаток товаров), иначе — счёт не закрыт"
          : code === "20" || code === "23"
            ? "остаток бывает законным (незавершённое производство), иначе — счёт не закрыт"
            : "косвенные расходы закрываются полностью: обычно это затраты, внесённые после закрытия, — карточка счёта покажет какие; затем перезакройте месяц",
    });
  }
  if (!items.length) {
    return {
      id: "afterClose",
      title,
      status: "ok",
      message: "Закрываемые счета (90, 91, 20–26, 44) в нуле.",
    };
  }
  return {
    id: "afterClose",
    title,
    status: strictFail ? "error" : "warning",
    message: `После закрытия остались суммы на ${items.length} счетах из тех, что закрываются в конце месяца.`,
    ...withItems(items),
  };
}

/**
 * Остатки на 90/91/25/26 на начало месяца = хвосты прошлых месяцев: после их закрытия внесли
 * документы. Регламентные операции при этом «Выполнено», поэтому одной проверки их наличия мало.
 */
async function carriedResiduals(
  conn: Connection,
  snap: BalanceSnapshot,
  month: string,
  orgKey: string | undefined,
): Promise<Record<string, unknown>[]> {
  const codes = CLOSED_ACCOUNTS.filter((c) => c.strict || c.code === "25" || c.code === "26").map(
    (c) => c.code,
  );
  const keys = [...snap.byKey.values()]
    .filter((a) => codes.includes(a.code.split(".")[0]!))
    .map((a) => a.key);
  const prevEnd = addDays(monthBounds(month).from, -1);
  const rows = await balanceByAccounts(conn, keys, orgKey, prevEnd);
  const net = netByTopCode({ byKey: snap.byKey, rows }, codes);
  return codes
    .filter((c) => (net.get(c) ?? 0) !== 0)
    .map((c) => ({ account: c, onDate: prevEnd, balance: rub(net.get(c)!) }));
}

// ─── 4б. Документы, созданные после закрытия ─────────────────────────────────

export interface LateDocument {
  ref: string;
  set: string;
  org?: string;
  createdAt: string;
  closingCreatedAt: string;
}

/**
 * Документы с движениями за месяц, созданные позже операции закрытия своей организации (по метке
 * времени в Ref_Key). Регламентные операции не в счёт. unknown — регистраторы без метки времени
 * (ссылка не UUIDv1, например из обмена): про них сказать нельзя.
 */
export function lateDocuments(
  records: readonly ODataEntity[],
  closedAt: ReadonlyMap<string, string>,
): { late: LateDocument[]; unknown: number } {
  const seen = new Set<string>();
  const late: LateDocument[] = [];
  let unknown = 0;
  const single = closedAt.size === 1 ? [...closedAt.values()][0] : undefined;
  for (const r of records) {
    const ref = refOf(r["Recorder"]);
    const set = entitySetOfType(r["Recorder_Type"]);
    if (!ref || !set || set === CLOSING_SET[0] || seen.has(ref)) continue;
    seen.add(ref);
    const org = refOf(r["Организация_Key"]);
    // Своё время закрытия у каждой организации; общее — только для движений без организации.
    const closing = org ? closedAt.get(org) : single;
    if (!closing) continue;
    const at = createdAt(ref);
    if (!at) {
      unknown += 1;
      continue;
    }
    if (at > closing)
      late.push({ ref, set, ...(org ? { org } : {}), createdAt: at, closingCreatedAt: closing });
  }
  return { late: late.sort((a, b) => a.createdAt.localeCompare(b.createdAt)), unknown };
}

async function checkLateDocuments(
  conn: Connection,
  bounds: { from: string; to: string },
  orgKey: string | undefined,
  closedAt: ReadonlyMap<string, string>,
  afterClose: CheckResult,
  orgs: ReadonlyMap<string, string>,
): Promise<CheckResult> {
  const id = "lateDocuments";
  const title = "Документы после закрытия";
  if (afterClose.status === "skipped") {
    return { id, title, status: "skipped", message: "Месяц не закрыт — сравнивать не с чем." };
  }
  // Время операции закрытия при перезакрытии не меняется: если закрываемые счета в нуле, документы,
  // созданные позже, уже учтены — сравнение дало бы ложную тревогу.
  if (afterClose.status === "ok") {
    return {
      id,
      title,
      status: "not_applicable",
      message: "Закрываемые счета после закрытия в нуле — поздние документы на результат не повлияли.",
    };
  }
  if (closedAt.size === 0) {
    return {
      id,
      title,
      status: "skipped",
      message: "Время операции закрытия не определить (ссылка без метки времени).",
    };
  }
  const source = await resolvePostingsSource(conn);
  const { rows } = await fetchAllForAggregation(
    conn,
    source.entitySet,
    {
      filter:
        and(
          cmp("Period", "ge", `datetime'${bounds.from}T00:00:00'`),
          cmp("Period", "le", `datetime'${bounds.to}T23:59:59'`),
          orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
        ) || undefined,
      select: ["Recorder", "Recorder_Type", "Организация_Key"],
    },
    `движения ${bounds.from}..${bounds.to}`,
  );
  const { late, unknown } = lateDocuments(rows, closedAt);
  if (!late.length) {
    return {
      id,
      title,
      status: "ok",
      message:
        "Документов, созданных после закрытия, нет — остатки на закрываемых счетах дали перепроведённые " +
        "документы или ошибка закрытия; карточка счёта покажет, какие движения не закрыты." +
        (unknown ? ` У ${unknown} документов время создания не определить.` : ""),
    };
  }
  const shown = late.slice(0, ITEMS_LIMIT);
  const bySet = new Map<string, string[]>();
  for (const d of shown) bySet.set(d.set, [...(bySet.get(d.set) ?? []), d.ref]);
  const info = new Map<string, ODataEntity>();
  for (const [set, refs] of bySet) {
    for (const [k, v] of await readByRefs(conn, set, refs, [
      DOC_FIELDS.ref,
      DOC_FIELDS.number,
      DOC_FIELDS.date,
    ]))
      info.set(k, v);
  }
  return {
    id,
    title,
    status: "warning",
    message:
      `${late.length} документ(ов) с движениями за месяц созданы позже операции закрытия — их суммы в ` +
      "закрытие не попали. Перезакройте месяц. Если месяц уже перезакрывали после их создания, а остатки " +
      "остались — причина в другом: время операции закрытия при перезакрытии не меняется." +
      (unknown ? ` У ${unknown} документов время создания не определить.` : ""),
    ...withItems(
      shown.map((d) => ({
        documentTitle: documentTitle(d.set),
        number: String(info.get(d.ref)?.[DOC_FIELDS.number] ?? "").trim(),
        date: String(info.get(d.ref)?.[DOC_FIELDS.date] ?? "").slice(0, 10),
        documentRef: d.ref,
        documentCreatedAt: d.createdAt,
        closingCreatedAt: d.closingCreatedAt,
        ...(d.org && orgs.get(d.org) ? { organization: orgs.get(d.org) } : {}),
      })),
    ),
    ...(late.length > ITEMS_LIMIT ? { itemsMore: late.length - ITEMS_LIMIT } : {}),
  };
}

// ─── 5. Амортизация ──────────────────────────────────────────────────────────

const ASSET_ACCOUNTS = ["01", "04"];

function checkDepreciation(snap: BalanceSnapshot, closingKinds: ReadonlySet<string>): CheckResult {
  const title = "Амортизация";
  // Счета выбытия (01.09, 04.09) — не объекты на балансе.
  const net = netByTopCode(snap, ASSET_ACCOUNTS, (a) => /^0[14]\.09/.test(a.code));
  const assets = ASSET_ACCOUNTS.filter((c) => (net.get(c) ?? 0) > 0);
  if (!assets.length) {
    return {
      id: "depreciation",
      title,
      status: "not_applicable",
      message: "Основных средств и НМА на балансе нет.",
    };
  }
  const has = (re: RegExp) => [...closingKinds].some((k) => re.test(k));
  const missing = [
    assets.includes("01") && !has(/Амортизац(?!.*НМА)/) ? "основных средств (01)" : "",
    assets.includes("04") && !has(/Амортизац.*НМА/) ? "НМА (04)" : "",
  ].filter(Boolean);
  return missing.length
    ? {
        id: "depreciation",
        title,
        status: "warning",
        message:
          `На балансе есть ${missing.join(" и ")}, а операции начисления амортизации за месяц нет. ` +
          "Если объекты не амортизируются (земля, полностью самортизированы) — всё в порядке.",
      }
    : { id: "depreciation", title, status: "ok", message: "Амортизация за месяц начислена." };
}

// ─── 6. Счета-фактуры на реализацию с НДС ────────────────────────────────────

/** Ссылки на документы-основания из шапки и табличной части счёта-фактуры. */
export function invoiceBases(sf: ODataEntity): string[] {
  const out: string[] = [];
  const base = refOf(sf["ДокументОснование"]);
  if (base) out.push(base);
  const rows = sf["ДокументыОснования"];
  if (Array.isArray(rows)) {
    for (const row of rows as ODataEntity[]) {
      const ref = refOf(row["ДокументОснование"]);
      if (ref) out.push(ref);
    }
  }
  return out;
}

async function checkVatInvoices(
  conn: Connection,
  bounds: { from: string; to: string },
  orgKey: string | undefined,
): Promise<CheckResult> {
  const title = "Счета-фактуры на реализацию";
  const available = await conn.available();
  const salesSet = resolveEntity(DOCUMENTS.sales, available);
  const sfSet = resolveEntity(DOCUMENTS.issuedInvoice, available);
  if (!salesSet) {
    return {
      id: "vatInvoices",
      title,
      status: "skipped",
      message: "Документы реализации не опубликованы в OData.",
    };
  }
  const meta = await conn.getMetadata();
  const props = new Set(meta.entities.get(salesSet)?.properties.map((p) => p.name) ?? []);
  if (!props.has("ДокументБезНДС")) {
    return {
      id: "vatInvoices",
      title,
      status: "skipped",
      message: "В реализации нет признака «Документ без НДС» — не определить, нужен ли счёт-фактура.",
    };
  }
  const orgFilter = orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined;
  const { rows: sales } = await collectDocuments(conn, salesSet, {
    baseFilter: and(orgFilter, cmp(DOC_FIELDS.posted, "eq", "true")) || undefined,
    dateField: DOC_FIELDS.date,
    ...bounds,
    select: [
      DOC_FIELDS.ref,
      DOC_FIELDS.number,
      DOC_FIELDS.date,
      DOC_FIELDS.amount,
      DOC_FIELDS.counterparty,
      "ДокументБезНДС",
    ],
  });
  const withVat = sales.filter((r) => r["ДокументБезНДС"] !== true);
  if (!withVat.length) {
    return {
      id: "vatInvoices",
      title,
      status: "not_applicable",
      message: sales.length
        ? `Все ${sales.length} реализаций месяца без НДС — счета-фактуры не нужны.`
        : "Проведённых реализаций за месяц нет.",
    };
  }
  if (!sfSet) {
    return {
      id: "vatInvoices",
      title,
      status: "skipped",
      message: `${withVat.length} реализаций с НДС, но «Счёт-фактура выданный» не опубликован в OData — не проверить.`,
    };
  }
  // Счёт-фактуру выставляют в течение 5 дней после отгрузки — берём с запасом в следующий месяц.
  const { rows: sfs } = await fetchAllForAggregation(
    conn,
    sfSet,
    {
      filter:
        and(
          orgFilter,
          cmp(DOC_FIELDS.deletionMark, "eq", "false"),
          cmp(DOC_FIELDS.date, "ge", `datetime'${bounds.from}T00:00:00'`),
          cmp(DOC_FIELDS.date, "lt", `datetime'${addDays(bounds.to, 11)}T00:00:00'`),
        ) || undefined,
    },
    `счета-фактуры ${bounds.from}..${addDays(bounds.to, 10)}`,
  );
  const covered = new Set(sfs.flatMap(invoiceBases));
  const missing = withVat.filter((r) => !covered.has(String(r[DOC_FIELDS.ref])));
  if (!missing.length) {
    return {
      id: "vatInvoices",
      title,
      status: "ok",
      message: `На все ${withVat.length} реализаций с НДС выставлены счета-фактуры.`,
    };
  }
  const cpSet = resolveEntity(CATALOGS.counterparties, available);
  const names = cpSet
    ? await resolveNames(
        conn,
        cpSet,
        missing.map((r) => refOf(r[DOC_FIELDS.counterparty]) ?? ""),
      )
    : new Map<string, string>();
  return {
    id: "vatInvoices",
    title,
    status: "warning",
    message:
      `${missing.length} из ${withVat.length} реализаций с НДС без счёта-фактуры (искали выставленные до ` +
      `${addDays(bounds.to, 10)}). Без счёта-фактуры НДС не попадёт в книгу продаж.`,
    ...withItems(
      missing.map((r) => ({
        number: String(r[DOC_FIELDS.number] ?? "").trim(),
        date: String(r[DOC_FIELDS.date] ?? "").slice(0, 10),
        ref: String(r[DOC_FIELDS.ref]),
        amount: num(r[DOC_FIELDS.amount]),
        counterparty: names.get(refOf(r[DOC_FIELDS.counterparty]) ?? "") ?? "",
      })),
    ),
  };
}

// ─── Инструмент ──────────────────────────────────────────────────────────────

const MONTH_RE = /^\d{4}-(0[1-9]|1[0-2])$/;

export function registerMonthCloseTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.accounting.check_month_close",
    {
      title: "Проверка закрытия месяца",
      description:
        "Проверка учёта за месяц — то, что бухгалтер смотрит перед сдачей отчётности и после закрытия " +
        "месяца: (1) закрытие месяца выполнено и без ошибок, закрыты ли предыдущие месяцы года; " +
        "(2) непроведённые документы месяца; (3) красное сальдо на конец месяца — деньги и товары в минусе, " +
        "расчёты 62/60 не на той стороне; (4) остатки на 90, 91 (должны быть в нуле) и 20–26, 44 после " +
        "закрытия и документы, созданные позже операции закрытия (если остатки есть); (5) амортизация при наличии " +
        "ОС и НМА; (6) реализации с НДС без счёта-фактуры. " +
        "Каждая проверка — статус ok / warning / error / not_applicable / skipped, пояснение и строки " +
        "(до 20). Остатки на 90/91/25/26 на начало месяца — признак, что прошлые месяцы закрыты не полностью. " +
        "closingCreatedAt — когда впервые создана операция закрытия (документы, созданные позже, могли в него " +
        "не попасть; перезакрытие этого времени не меняет). Для перехода в карточку счёта: analyticsRefs с " +
        "analyticsTypes Catalog_Контрагенты — counterpartyRef, иначе dimensionRef. У регламентных операций в БП " +
        "Posted = false — это нормально. Только чтение: закрыть месяц через OData нельзя, только проверить результат.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        month: z.string().regex(MONTH_RE, "Месяц в формате YYYY-MM").describe("Проверяемый месяц, YYYY-MM"),
      }),
      outputSchema: monthCloseResultSchema,
    },
    ({ database, organization, month }) =>
      guard("read.accounting.check_month_close", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = organization ? await resolveOrganization(conn, organization) : undefined;
        const orgs = new Map((await listOrganizations(conn)).map((o) => [o.ref, o.name]));
        const bounds = monthBounds(month);

        const closing = await checkClosing(conn, month, orgs, org?.ref);
        const snap = await balanceSnapshot(conn, bounds.to, org?.ref);
        const carried = await carriedResiduals(conn, snap, month, org?.ref);
        if (carried.length) {
          const prev = closing.earlier;
          closing.earlier = {
            id: prev.id,
            title: prev.title,
            status: "warning",
            message:
              (prev.status === "warning" ? `${prev.message} ` : "") +
              "На начало месяца на закрываемых счетах уже были остатки: прошлые месяцы закрыты не полностью — " +
              "после их закрытия внесли документы. Перезакройте их по порядку, начиная с самого раннего.",
            ...withItems([...(prev.items ?? []), ...carried]),
          };
        }
        const closed = closing.result.status === "ok";
        const afterClose = checkAfterClose(snap, closed);
        const closingKinds = new Set(
          (closing.result.items ?? []).flatMap((i) =>
            Array.isArray(i["operations"]) ? (i["operations"] as string[]) : [],
          ),
        );
        const checks: CheckResult[] = [
          closing.result,
          closing.earlier,
          await checkUnposted(conn, bounds, org?.ref),
          await checkRedBalances(conn, snap),
          afterClose,
          await checkLateDocuments(conn, bounds, org?.ref, closing.closedAt, afterClose, orgs),
          checkDepreciation(snap, closingKinds),
          await checkVatInvoices(conn, bounds, org?.ref),
        ];
        const count = (s: CheckStatus) => checks.filter((c) => c.status === s).length;
        return ok({
          database: conn.cfg.name,
          ...(org ? { organization: org.name } : {}),
          month,
          period: bounds,
          balancesAsOf: `${nextDay(bounds.to)}T00:00:00`,
          status: overallStatus(checks),
          summary: {
            errors: count("error"),
            warnings: count("warning"),
            ok: count("ok"),
            notApplicable: count("not_applicable"),
            skipped: count("skipped"),
          },
          checks,
          scan: { elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
