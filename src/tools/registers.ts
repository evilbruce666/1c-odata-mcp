import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField, dateField } from "./_shared.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { ACCOUNT_PREFIX, CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import {
  balanceByAccounts,
  resolveAccounts,
  resolveNames,
  num,
  turnoversByAccounts,
  BALANCE_AND_TURNOVERS,
  GUID_RE,
  EMPTY_GUID,
  REGISTER_RECORDS,
  accountLookupKey,
  accountsByRef,
  fetchRegistrarPostings,
  readDocumentHeader,
  requireDocumentEntity,
  type Account,
  type PostingsSource,
} from "../odata/accounting.js";
import { resolveOrganization } from "../odata/orgs.js";
import { collectDocuments, emptyMeta, addMeta, type ScanMeta } from "../odata/aggregate.js";
import {
  getSalesResultSchema,
  getCashflowResultSchema,
  getDebtorsResultSchema,
  getInventoryResultSchema,
  getAccountTurnoverResultSchema,
  getDocumentPostingsResultSchema,
} from "../schemas/output.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";

// Деньги копим в целых копейках (float-сложение тысяч сумм даёт дрейф).
const toCents = (v: unknown): number => Math.round(num(v) * 100);
const fromCents = (c: number): number => Math.round(c) / 100;

/**
 * Суммирует СуммаДокумента по проведённым документам за период (+орг). Через
 * collectDocuments: полная выборка с авто-чанкингом и громким переполнением
 * (раньше стояло под общим maxRows=1000 → годовые итоги занижались).
 */
async function sumDocuments(
  conn: Connection,
  docKeys: readonly string[],
  from: string,
  to: string,
  orgKey: string | undefined,
): Promise<{ totalCents: number; perSet: Record<string, number>; usedSets: string[]; meta: ScanMeta }> {
  const available = await conn.available();
  const baseFilter =
    and(
      orgKey ? cmp(DOC_FIELDS.organization, "eq", odataGuid(orgKey)) : undefined,
      cmp(DOC_FIELDS.posted, "eq", "true"),
    ) || undefined;
  const perSet: Record<string, number> = {};
  const usedSets: string[] = [];
  let totalCents = 0;
  let meta = emptyMeta();

  for (const key of docKeys) {
    const set = resolveEntity([key], available);
    if (!set) continue;
    usedSets.push(set);
    const { rows, meta: m } = await collectDocuments(conn, set, {
      baseFilter,
      dateField: DOC_FIELDS.date,
      from,
      to,
      select: [DOC_FIELDS.date, DOC_FIELDS.amount],
    });
    meta = addMeta(meta, m);
    const c = rows.reduce((acc, r) => acc + toCents(r[DOC_FIELDS.amount]), 0);
    perSet[set] = fromCents(c);
    totalCents += c;
  }
  return { totalCents, perSet, usedSets, meta };
}

/** Код счёта или префикс: «51», «60», «90.01», «90.01.1», забалансовые «МЦ.04», «001». */
const ACCOUNT_CODE_RE = /^[0-9A-Za-zА-Яа-яЁё]+(\.[0-9A-Za-zА-Яа-яЁё]+)*$/;

/** Сальдо/обороты ОСВ в целых копейках. */
export interface TurnoverCents {
  openingDr: number;
  openingCr: number;
  turnoverDr: number;
  turnoverCr: number;
  closingDr: number;
  closingCr: number;
}

const zeroTurnover = (): TurnoverCents => ({
  openingDr: 0,
  openingCr: 0,
  turnoverDr: 0,
  turnoverCr: 0,
  closingDr: 0,
  closingCr: 0,
});

const turnoverToRub = (t: TurnoverCents) => ({
  openingDebit: fromCents(t.openingDr),
  openingCredit: fromCents(t.openingCr),
  debitTurnover: fromCents(t.turnoverDr),
  creditTurnover: fromCents(t.turnoverCr),
  closingDebit: fromCents(t.closingDr),
  closingCredit: fromCents(t.closingCr),
});

/**
 * Сворачивает строки «Остатков и оборотов» в ОСВ: итог и разбивка по субсчетам.
 * Dr и Cr копятся РАЗДЕЛЬНО по строкам (строка = счёт × измерения × субконто) —
 * это развёрнутое сальдо, как в стандартной ОСВ по счёту с аналитикой.
 * Всё в целых копейках — без float-дрейфа.
 */
export function aggregateAccountTurnover(
  rows: readonly ODataEntity[],
  accounts: readonly Account[],
): {
  total: TurnoverCents;
  byAccount: Array<{ account: Account; sums: TurnoverCents }>;
  consistent: boolean;
} {
  const f = BALANCE_AND_TURNOVERS.fields;
  const total = zeroTurnover();
  const per = new Map<string, TurnoverCents>();
  for (const r of rows) {
    const key = String(r["Account_Key"] ?? "");
    const acc = per.get(key) ?? zeroTurnover();
    const add = (k: keyof TurnoverCents, field: string): void => {
      const c = toCents(r[field]);
      acc[k] += c;
      total[k] += c;
    };
    add("openingDr", f.openingDr);
    add("openingCr", f.openingCr);
    add("turnoverDr", f.turnoverDr);
    add("turnoverCr", f.turnoverCr);
    add("closingDr", f.closingDr);
    add("closingCr", f.closingCr);
    per.set(key, acc);
  }
  // Субсчета — в порядке плана счетов; без движений и остатков не показываем.
  const byAccount = accounts
    .filter((a) => per.has(a.key))
    .map((a) => ({ account: a, sums: per.get(a.key)! }));
  // Контроль: Сн + ОбДт − ОбКт = Ск (в копейках, по сальдо «Дт − Кт»).
  const consistent =
    total.openingDr - total.openingCr + total.turnoverDr - total.turnoverCr ===
    total.closingDr - total.closingCr;
  return { total, byAccount, consistent };
}

// ─── Проводки документа ───────────────────────────────────────────────────────

/**
 * Строгое число из поля записи регистра: отсутствие/null/не-число — ошибка
 * (с ИМЕНЕМ поля, без значения), а не молчаливый 0.
 */
function strictNumber(v: unknown, field: string, row: number): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) return Number(v);
  throw new Error(`Поле ${field} (строка #${row}) отсутствует или не является числом.`);
}

/** Необязательное число: поля нет в строке — undefined; есть, но не число — ошибка. */
function optionalNumber(r: ODataEntity, field: string, row: number): number | undefined {
  const v = r[field];
  if (v === undefined || v === null) return undefined;
  return strictNumber(v, field, row);
}

/** GUID-поле, где пустая ссылка = «нет значения». */
function refOrUndefined(v: unknown): string | undefined {
  return typeof v === "string" && v !== "" && v !== EMPTY_GUID ? v : undefined;
}

export interface PostingDimension {
  index: number;
  type?: string;
  ref?: string;
  value?: string;
}

export interface PostingSide {
  accountCode: string;
  accountName: string;
  accountRef: string;
  dimensions?: PostingDimension[];
  divisionRef?: string;
}

export interface Posting {
  period: string;
  lineNumber?: number;
  active?: boolean;
  debit: PostingSide | null;
  credit: PostingSide | null;
  amount: number;
  quantityDebit?: number;
  quantityCredit?: number;
  currencyAmountDebit?: number;
  currencyAmountCredit?: number;
  organizationRef?: string;
  content?: string;
}

const GUIDISH = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

function dimensionsOf(r: ODataEntity, source: PostingsSource, side: "Dr" | "Cr"): PostingDimension[] {
  const out: PostingDimension[] = [];
  for (const d of source.extDimensions) {
    if (d.side !== side) continue;
    const v = r[d.field];
    if (v === undefined || v === null || v === "" || v === EMPTY_GUID) continue;
    const t = d.typeField ? r[d.typeField] : undefined;
    const type = typeof t === "string" && t ? t.slice(t.lastIndexOf(".") + 1) : undefined;
    const s = String(v);
    out.push({ index: d.index, ...(type ? { type } : {}), ...(GUIDISH.test(s) ? { ref: s } : { value: s }) });
  }
  return out;
}

/**
 * Сворачивает записи регистра одного регистратора в проводки, итоги и
 * корреспонденции. Деньги — в целых копейках. Итоги и byCorrespondence — только
 * по активным записям (неактивные в 1С не влияют на обороты); в postings — все.
 * Сторона с пустым счётом (забалансовые проводки) — null и не идёт в итог стороны.
 */
export function aggregateDocumentPostings(
  rows: readonly ODataEntity[],
  accounts: ReadonlyMap<string, Account>,
  source: PostingsSource,
): {
  postings: Posting[];
  debitCents: number;
  creditCents: number;
  inactiveCount: number;
  byCorrespondence: Array<{
    debitAccount: string | null;
    creditAccount: string | null;
    amount: number;
    entries: number;
  }>;
} {
  const f = REGISTER_RECORDS.required;
  const o = REGISTER_RECORDS.optional;
  const has = (field: string): boolean => source.properties.has(field);
  let debitCents = 0;
  let creditCents = 0;
  let inactiveCount = 0;
  const corr = new Map<
    string,
    { debitAccount: string | null; creditAccount: string | null; cents: number; entries: number }
  >();

  const sideOf = (r: ODataEntity, field: string, side: "Dr" | "Cr", row: number): PostingSide | null => {
    const raw = r[field];
    if (typeof raw !== "string") {
      throw new Error(`Поле ${field} (строка #${row}) отсутствует или не является ссылкой.`);
    }
    const key = refOrUndefined(raw);
    if (!key) return null;
    const acc = accounts.get(accountLookupKey(key));
    if (!acc) throw new Error(`Счёт ${key} (поле ${field}) не сопоставлен с планом счетов.`);
    const dims = dimensionsOf(r, source, side);
    const division = refOrUndefined(r[side === "Dr" ? o.divisionDr : o.divisionCr]);
    return {
      accountCode: acc.code,
      accountName: acc.description,
      accountRef: acc.key,
      ...(dims.length ? { dimensions: dims } : {}),
      ...(division ? { divisionRef: division } : {}),
    };
  };

  const postings = rows.map((r, i): Posting => {
    const row = i + 1;
    const period = r[f.period];
    if (typeof period !== "string" || !period) {
      throw new Error(`Поле ${f.period} (строка #${row}) отсутствует или не является датой.`);
    }
    const cents = Math.round(strictNumber(r[f.amount], f.amount, row) * 100);
    const debit = sideOf(r, f.accountDr, "Dr", row);
    const credit = sideOf(r, f.accountCr, "Cr", row);
    const activeRaw = has(o.active) ? r[o.active] : undefined;
    const active = typeof activeRaw === "boolean" ? activeRaw : undefined;
    const lineRaw = has(o.lineNumber) ? optionalNumber(r, o.lineNumber, row) : undefined;

    if (active === false) {
      inactiveCount += 1;
    } else {
      if (debit) debitCents += cents;
      if (credit) creditCents += cents;
      const dk = debit?.accountCode ?? null;
      const ck = credit?.accountCode ?? null;
      const k = `${dk ?? ""}\u0000${ck ?? ""}`;
      const g = corr.get(k) ?? { debitAccount: dk, creditAccount: ck, cents: 0, entries: 0 };
      g.cents += cents;
      g.entries += 1;
      corr.set(k, g);
    }

    const nz = (field: string): number | undefined => {
      if (!has(field)) return undefined;
      const v = optionalNumber(r, field, row);
      return v === undefined || v === 0 ? undefined : v;
    };
    const qDr = nz(o.quantityDr);
    const qCr = nz(o.quantityCr);
    const cDr = nz(o.currencyAmountDr);
    const cCr = nz(o.currencyAmountCr);
    const org = has(o.organization) ? refOrUndefined(r[o.organization]) : undefined;
    const content =
      has(o.content) && typeof r[o.content] === "string" && r[o.content] ? String(r[o.content]) : undefined;
    return {
      period,
      ...(lineRaw !== undefined ? { lineNumber: lineRaw } : {}),
      ...(active !== undefined ? { active } : {}),
      debit,
      credit,
      amount: fromCents(cents),
      ...(qDr !== undefined ? { quantityDebit: qDr } : {}),
      ...(qCr !== undefined ? { quantityCredit: qCr } : {}),
      ...(cDr !== undefined ? { currencyAmountDebit: cDr } : {}),
      ...(cCr !== undefined ? { currencyAmountCredit: cCr } : {}),
      ...(org ? { organizationRef: org } : {}),
      ...(content ? { content } : {}),
    };
  });

  const byCorrespondence = [...corr.values()]
    .sort(
      (a, b) =>
        b.cents - a.cents ||
        (a.debitAccount ?? "").localeCompare(b.debitAccount ?? "") ||
        (a.creditAccount ?? "").localeCompare(b.creditAccount ?? ""),
    )
    .map((g) => ({
      debitAccount: g.debitAccount,
      creditAccount: g.creditAccount,
      amount: fromCents(g.cents),
      entries: g.entries,
    }));
  return { postings, debitCents, creditCents, inactiveCount, byCorrespondence };
}

const str = (v: unknown): string | undefined => (typeof v === "string" && v !== "" ? v : undefined);
const bool = (v: unknown): boolean | undefined => (typeof v === "boolean" ? v : undefined);

/**
 * Проводки одного документа целиком: проверка вида документа по $metadata → шапка
 * (not_found, если нет) → записи регистра по регистратору (отбор в 1С) → счета
 * пачкой → итоги. Только GET-запросы.
 */
export async function getDocumentPostings(conn: Connection, documentEntity: string, documentRef: string) {
  const t0 = Date.now();
  const ref = documentRef.trim().replace(/^\{|\}$/g, "");
  // До любого запроса к 1С.
  if (!GUID_RE.test(ref))
    throw new InputError(`documentRef должен быть GUID (Ref_Key документа): ${documentRef}`);
  const em = await requireDocumentEntity(conn, documentEntity);
  const doc = await readDocumentHeader(conn, em, ref);
  const { rows, meta, source, filter } = await fetchRegistrarPostings(conn, documentEntity, ref);

  const accountKeys: string[] = [];
  for (const r of rows) {
    for (const fld of [REGISTER_RECORDS.required.accountDr, REGISTER_RECORDS.required.accountCr]) {
      const k = refOrUndefined(r[fld]);
      if (k) accountKeys.push(k);
    }
  }
  const accounts = await accountsByRef(conn, accountKeys);
  const agg = aggregateDocumentPostings(rows, accounts, source);

  // Организация: из шапки; если в шапке нет — единственная организация проводок.
  let orgRef = refOrUndefined(doc["Организация_Key"]);
  if (!orgRef) {
    const fromRows = new Set(agg.postings.map((p) => p.organizationRef).filter(Boolean));
    if (fromRows.size === 1) orgRef = [...fromRows][0];
  }
  let orgName: string | undefined;
  if (orgRef) {
    const orgSet = resolveEntity(CATALOGS.organizations, await conn.available());
    if (orgSet) orgName = (await resolveNames(conn, orgSet, [orgRef])).get(orgRef) || undefined;
  }

  const posted = bool(doc["Posted"]);
  const note =
    // Причину пустого набора по Posted НЕ выводим: в БП 3.0 бывает Posted=false при
    // существующих движениях (напр. РегламентнаяОперация закрытия месяца, Состояние=Выполнено).
    rows.length === 0
      ? "У документа нет записей в регистре Хозрасчетный."
      : agg.inactiveCount > 0
        ? `Неактивных записей: ${agg.inactiveCount} — показаны в postings (active=false), в итоги не входят.`
        : undefined;

  return {
    database: conn.cfg.name,
    document: {
      entitySet: documentEntity,
      ref: str(doc["Ref_Key"]) ?? ref,
      ...(str(doc["Number"]) ? { number: str(doc["Number"]) } : {}),
      ...(str(doc["Date"]) ? { date: str(doc["Date"]) } : {}),
      ...(posted !== undefined ? { posted } : {}),
      ...(bool(doc["DeletionMark"]) !== undefined ? { deletionMark: bool(doc["DeletionMark"]) } : {}),
      ...(orgName ? { organization: orgName } : {}),
      ...(orgRef ? { organizationRef: orgRef } : {}),
      ...(str(doc["ВидОперации"]) ? { operation: str(doc["ВидОперации"]) } : {}),
      ...(str(doc["Состояние"]) ? { state: str(doc["Состояние"]) } : {}),
      ...(str(doc["Комментарий"]) ? { comment: str(doc["Комментарий"]) } : {}),
    },
    postingsCount: agg.postings.length,
    debitTotal: fromCents(agg.debitCents),
    creditTotal: fromCents(agg.creditCents),
    postings: agg.postings,
    byCorrespondence: agg.byCorrespondence,
    ...(note ? { note } : {}),
    source: { entitySet: source.entitySet, filter },
    scan: { rowsScanned: meta.rowsScanned, windows: meta.chunks, elapsedMs: Date.now() - t0 },
  };
}

async function orgKeyOf(conn: Connection, organization?: string): Promise<{ key?: string; name?: string }> {
  if (!organization) return {};
  const org = await resolveOrganization(conn, organization);
  return { key: org.ref, name: org.name };
}

export function registerRegisterTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.analytics.get_sales",
    {
      title: "Продажи за период",
      description:
        "Сумма выручки по проведённым документам реализации за период (по полю СуммаДокумента). " +
        "Можно ограничить организацией. Период задаётся датами YYYY-MM-DD.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
      },
      outputSchema: getSalesResultSchema,
    },
    ({ database, organization, from, to }) =>
      guard("read.analytics.get_sales", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const r = await sumDocuments(conn, [...DOCUMENTS.sales], from, to, org.key);
        if (r.usedSets.length === 0) {
          return fail("Документы реализации не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          period: { from, to },
          total: fromCents(r.totalCents),
          byDocument: r.perSet,
          scan: { documentsScanned: r.meta.rowsScanned, windows: r.meta.chunks, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  server.registerTool(
    "read.analytics.get_cashflow",
    {
      title: "Движение денежных средств",
      description:
        "Приход и расход денег за период по проведённым банковским и кассовым документам. " +
        "Возвращает приток, отток и сальдо. Можно ограничить организацией. Период — даты YYYY-MM-DD.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
      },
      outputSchema: getCashflowResultSchema,
    },
    ({ database, organization, from, to }) =>
      guard("read.analytics.get_cashflow", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const [inflow, outflow] = await Promise.all([
          sumDocuments(conn, [...DOCUMENTS.bankIn, ...DOCUMENTS.cashIn], from, to, org.key),
          sumDocuments(conn, [...DOCUMENTS.bankOut, ...DOCUMENTS.cashOut], from, to, org.key),
        ]);
        if (inflow.usedSets.length === 0 && outflow.usedSets.length === 0) {
          return fail("Банковские/кассовые документы не опубликованы в OData. Добавьте их в «Состав OData».");
        }
        const scanned = inflow.meta.rowsScanned + outflow.meta.rowsScanned;
        const windows = inflow.meta.chunks + outflow.meta.chunks;
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          period: { from, to },
          inflow: fromCents(inflow.totalCents),
          outflow: fromCents(outflow.totalCents),
          net: fromCents(inflow.totalCents - outflow.totalCents),
          byDocument: { ...inflow.perSet, ...outflow.perSet },
          scan: { documentsScanned: scanned, windows, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  server.registerTool(
    "read.analytics.get_debtors",
    {
      title: "Дебиторская задолженность",
      description:
        "Кто и сколько должен компании: сальдо счёта 62 (расчёты с покупателями) из регистра " +
        "бухгалтерии Хозрасчетный, сгруппированное по контрагентам. Дебетовое сальдо = долг клиента, " +
        "кредитовое = полученные авансы (вычитается). Можно ограничить организацией. По умолчанию " +
        "берётся текущее сальдо; параметром asOf=YYYY-MM-DD можно получить дебиторку на конец " +
        "указанной даты (для аудита/исторических отчётов). Возвращает только долг > 0.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        asOf: dateField("Дата сальдо — на конец этой даты (без параметра — текущее)").optional(),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько контрагентов вернуть"),
      },
      outputSchema: getDebtorsResultSchema,
    },
    ({ database, organization, asOf, limit }) =>
      guard("read.analytics.get_debtors", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, ACCOUNT_PREFIX.receivables);
        const rows = await balanceByAccounts(
          conn,
          accounts.map((a) => a.key),
          org.key,
          asOf,
        );

        // Сальдо копим в копейках (целые) — без float-дрейфа на тысячах строк.
        const byCp = new Map<string, number>();
        for (const r of rows) {
          const cp = String(r["ExtDimension1"] ?? "");
          if (!cp) continue;
          byCp.set(cp, (byCp.get(cp) ?? 0) + toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]));
        }

        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const names = cpSet ? await resolveNames(conn, cpSet, byCp.keys()) : new Map<string, string>();

        const debtors = [...byCp.entries()]
          .filter(([, cents]) => cents > 0)
          .map(([ref, cents]) => ({
            counterparty: names.get(ref) ?? ref,
            ref,
            amount: fromCents(cents),
          }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const totalCents = [...byCp.values()].filter((c) => c > 0).reduce((s, c) => s + c, 0);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          ...(asOf ? { asOf } : {}),
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalReceivable: fromCents(totalCents),
          count: debtors.length,
          debtors,
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  server.registerTool(
    "read.accounting.get_account_turnover",
    {
      title: "ОСВ по счёту",
      description:
        "Оборотно-сальдовая ведомость по бухгалтерскому счёту (регистр бухгалтерии Хозрасчетный, " +
        "виртуальная таблица «Остатки и обороты») за период: сальдо на начало Дт/Кт, обороты Дт/Кт, " +
        "сальдо на конец Дт/Кт — итогом и по субсчетам. account — код счёта или префикс " +
        "(«51», «60», «62», «90.01»): берутся все субсчета, код которых начинается с него. " +
        "Сальдо развёрнутое по аналитике (измерения и субконто), как в стандартной ОСВ по счёту. " +
        "Можно ограничить организацией. Период — даты YYYY-MM-DD включительно.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        account: z
          .string()
          .trim()
          .min(1)
          .max(20)
          .regex(ACCOUNT_CODE_RE, "Код счёта: цифры/буквы через точку, напр. 51, 60.01, 90.01.1")
          .describe("Код счёта или префикс субсчетов, напр. 51, 60, 62, 90.01"),
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
      },
      outputSchema: getAccountTurnoverResultSchema,
    },
    ({ database, organization, account, from, to }) =>
      guard("read.accounting.get_account_turnover", async () => {
        const t0 = Date.now();
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, [account]);
        if (accounts.length === 0) {
          throw new InputError(`Счёт "${account}" не найден в плане счетов «Хозрасчётный».`);
        }
        // resolveAccounts листает до maxRows без сигнала об усечке: упёрлись в потолок —
        // список субсчетов мог быть неполным, а неполная ОСВ хуже явной ошибки.
        if (accounts.length >= conn.behavior.maxRows) {
          throw new InputError(
            `Под префикс "${account}" подходит слишком много счетов (≥ ${conn.behavior.maxRows}). ` +
              `Уточните код счёта.`,
          );
        }
        const { rows, meta } = await turnoversByAccounts(
          conn,
          accounts.map((a) => a.key),
          from,
          to,
          org.key,
        );
        const agg = aggregateAccountTurnover(rows, accounts);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          account,
          period: { from, to },
          ...turnoverToRub(agg.total),
          consistent: agg.consistent,
          accounts: agg.byAccount.map(({ account: a, sums }) => ({
            code: a.code,
            description: a.description,
            ref: a.key,
            ...turnoverToRub(sums),
          })),
          ...(rows.length === 0 ? { note: "За период нет ни остатков, ни движений по счёту." } : {}),
          scan: { rowsScanned: meta.rowsScanned, windows: meta.chunks, elapsedMs: Date.now() - t0 },
        });
      }),
  );

  server.registerTool(
    "read.accounting.get_document_postings",
    {
      title: "Проводки документа",
      description:
        "Бухгалтерские проводки, сформированные одним документом (регистратором), из регистра " +
        "бухгалтерии Хозрасчетный: Дт/Кт (код, наименование, Ref_Key счёта, субконто), сумма, " +
        "количество/валютная сумма, итоги Дт/Кт и свод по корреспонденциям счетов. Подходит для " +
        "любого вида документа (Document_*). Только чтение. documentEntity — имя документа из " +
        "list_entities (class=document), documentRef — его Ref_Key (из search_documents).",
      inputSchema: {
        database: databaseField,
        documentEntity: z
          .string()
          .trim()
          .regex(/^Document_[^\s()'/?#&]+$/, "Имя документа вида Document_<Имя>")
          .describe("Имя документа, напр. Document_РегламентнаяОперация"),
        documentRef: z
          .string()
          .trim()
          .regex(
            /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/,
            "GUID",
          )
          .describe("Ref_Key документа (GUID)"),
      },
      outputSchema: getDocumentPostingsResultSchema,
    },
    ({ database, documentEntity, documentRef }) =>
      guard("read.accounting.get_document_postings", async () =>
        ok(await getDocumentPostings(ctx.db(database), documentEntity, documentRef)),
      ),
  );

  server.registerTool(
    "read.analytics.get_inventory",
    {
      title: "Остатки товаров",
      description:
        "Остатки товаров/материалов на складах: сальдо счетов 41/10/43 регистра Хозрасчетный, " +
        "сгруппированное по номенклатуре. Возвращает количество и сумму остатка. Можно ограничить " +
        "организацией. По умолчанию — текущие остатки; параметром asOf=YYYY-MM-DD можно получить " +
        "остатки на конец указанной даты (для инвентаризации/аудита). " +
        "(В БП 3.0 учёт ведётся на счетах бухучёта, а не в отдельном регистре остатков.)",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        asOf: dateField("Дата остатков — на конец этой даты (без параметра — текущие)").optional(),
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько позиций вернуть"),
      },
      outputSchema: getInventoryResultSchema,
    },
    ({ database, organization, asOf, limit }) =>
      guard("read.analytics.get_inventory", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const org = await orgKeyOf(conn, organization);
        const accounts = await resolveAccounts(conn, ACCOUNT_PREFIX.inventory);
        const rows = await balanceByAccounts(
          conn,
          accounts.map((a) => a.key),
          org.key,
          asOf,
        );

        // Сумма — в копейках (целые), количество — float (округляем до 3 знаков).
        const byItem = new Map<string, { qty: number; cents: number }>();
        for (const r of rows) {
          const item = String(r["ExtDimension1"] ?? "");
          if (!item) continue;
          const cur = byItem.get(item) ?? { qty: 0, cents: 0 };
          cur.qty += num(r["КоличествоBalanceDr"]) - num(r["КоличествоBalanceCr"]);
          cur.cents += toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]);
          byItem.set(item, cur);
        }

        const nomSet = resolveEntity(CATALOGS.nomenclature, await conn.available());
        const names = nomSet ? await resolveNames(conn, nomSet, byItem.keys()) : new Map<string, string>();

        const present = [...byItem.entries()].filter(
          ([, v]) => Math.abs(v.qty) > 0.0001 || Math.abs(v.cents) > 0,
        );
        const items = present
          .map(([ref, v]) => ({
            item: names.get(ref) ?? ref,
            ref,
            quantity: Math.round(v.qty * 1000) / 1000,
            amount: fromCents(v.cents),
          }))
          .sort((a, b) => b.amount - a.amount)
          .slice(0, limit);

        const totalCents = present.reduce((s, [, v]) => s + v.cents, 0);
        return ok({
          database: conn.cfg.name,
          organization: org.name,
          ...(asOf ? { asOf } : {}),
          accounts: accounts.map((a) => `${a.code} ${a.description}`),
          totalAmount: fromCents(totalCents),
          count: items.length,
          items,
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
