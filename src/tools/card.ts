import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField, dateField } from "./_shared.js";
import { accountPrefixes } from "../config/accounting.js";
import { CATALOGS, DOC_FIELDS, resolveEntity } from "../config/mapping.js";
import {
  accountLookupKey,
  accountsByRef,
  balanceByAccounts,
  readByRefs,
  resolveAccounts,
  resolveNames,
  type Account,
} from "../odata/accounting.js";
import { cardEntries, fetchMovements, netBalance, type CardEntry, type Dim } from "../odata/account-card.js";
import { addDays, localToday } from "../odata/settlements-aging.js";
import { listOrganizations, resolveOrganization } from "../odata/orgs.js";
import { createdAt, documentTitle, printNumber } from "../odata/names.js";
import { refOf } from "../odata/settlements-aging.js";
import { InputError } from "../errors.js";
import {
  accountCardResultSchema,
  balanceOriginResultSchema,
  reconciliationResultSchema,
} from "../schemas/output.js";

const rub = (c: number): number => Math.round(c) / 100;
const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;
const guidField = (label: string) =>
  z
    .string()
    .trim()
    .regex(GUID, "GUID")
    .transform((s) => s.replace(/[{}]/g, "").toLowerCase())
    .describe(label);
/** Код счёта или префикс: «51», «60», «90.01», забалансовые «МЦ.04». */
const ACCOUNT_CODE_RE = /^[0-9A-Za-zА-Яа-яЁё]+(\.[0-9A-Za-zА-Яа-яЁё]+)*$/;

async function orgOf(conn: Connection, organization?: string) {
  return organization ? await resolveOrganization(conn, organization) : undefined;
}

/** Наименования субконто и номера документов для показанных строк — пачками по виду объекта. */
async function describeRefs(
  conn: Connection,
  entries: readonly CardEntry[],
  extraDims: readonly Dim[] = [],
): Promise<{
  name: (d: Dim) => string;
  doc: (set?: string, ref?: string) => { number?: string; date?: string };
}> {
  const catalogs = new Map<string, Set<string>>();
  const documents = new Map<string, Set<string>>();
  const add = (m: Map<string, Set<string>>, set: string, ref: string) =>
    m.set(set, (m.get(set) ?? new Set()).add(ref));
  const addDim = (d: Dim) => {
    if (d.set?.startsWith("Catalog_")) add(catalogs, d.set, d.ref);
    else if (d.set?.startsWith("Document_")) add(documents, d.set, d.ref);
  };
  for (const e of entries) {
    if (e.recorderSet && e.recorder) add(documents, e.recorderSet, e.recorder);
    for (const d of [...e.dims, ...e.corrDims]) addDim(d);
  }
  for (const d of extraDims) addDim(d);
  const names = new Map<string, string>();
  for (const [set, refs] of catalogs)
    for (const [k, v] of await resolveNames(conn, set, refs)) names.set(k, v);
  const docs = new Map<string, { number?: string; date?: string }>();
  for (const [set, refs] of documents) {
    for (const [k, d] of await readByRefs(conn, set, refs, [
      DOC_FIELDS.ref,
      DOC_FIELDS.number,
      DOC_FIELDS.date,
    ])) {
      const number = String(d[DOC_FIELDS.number] ?? "").trim();
      const date = String(d[DOC_FIELDS.date] ?? "").slice(0, 10);
      docs.set(k, { ...(number ? { number } : {}), ...(date ? { date } : {}) });
    }
  }
  const short = (set?: string) => (set ?? "").replace(/^(Document|Catalog)_/, "");
  return {
    name: (d) => {
      if (d.set?.startsWith("Document_")) {
        const info = docs.get(d.ref);
        return `${short(d.set)}${info?.number ? ` № ${info.number}` : ""}${info?.date ? ` от ${info.date}` : ""}`;
      }
      return names.get(d.ref) || d.ref;
    },
    doc: (_set, ref) => (ref ? (docs.get(ref) ?? {}) : {}),
  };
}

async function codesOf(conn: Connection, known: readonly Account[], entries: readonly CardEntry[]) {
  const byKey = new Map(known.map((a) => [accountLookupKey(a.key), a]));
  const missing = entries
    .map((e) => e.corrAccount)
    .filter((k): k is string => !!k && !byKey.has(accountLookupKey(k)));
  for (const [k, a] of await accountsByRef(conn, missing)) byKey.set(k, a);
  return (key?: string) => (key ? (byKey.get(accountLookupKey(key))?.code ?? key) : undefined);
}

// ─── Карточка счёта ──────────────────────────────────────────────────────────

/**
 * Предел текста ответа карточки — с запасом к норме 50 000 символов (MAX_RESPONSE_CHARS в тестах):
 * больший ответ клиент уносит в файл или обрезает, и конец периода теряется уже у него.
 */
export const MAX_CARD_CHARS = 45_000;
/** Строк на страницу: наименования разрешаются запросами, незачем разрешать те, что не поместятся. */
export const MAX_CARD_PAGE = 100;
/** Группы карточки занимают не больше этого; иначе отказ — обрезанная сводка хуже её отсутствия. */
const MAX_GROUP_CHARS = 40_000;

export const CARD_GROUPS = ["month", "corrAccount", "counterparty", "contract"] as const;
export type CardGroup = (typeof CARD_GROUPS)[number];

/** Символов, которые значение занимает в ответе ok() как элемент массива верхнего уровня. */
const elementChars = (v: unknown): number => {
  const text = JSON.stringify(v, null, 2);
  return text.length + text.split("\n").length * 4 + 2;
};

/** Сколько первых строк поместится в остаток предела. Хотя бы одна — иначе offset не сдвинется. */
export function fitToChars<T>(items: readonly T[], usedChars: number, maxChars = MAX_CARD_CHARS): T[] {
  let used = usedChars;
  let n = 0;
  for (const item of items) {
    used += elementChars(item);
    if (used > maxChars && n > 0) break;
    n += 1;
  }
  return items.slice(0, n);
}

const isCounterparty = (d: Dim) => d.set === "Catalog_Контрагенты";
const isContract = (d: Dim) => !!d.set?.startsWith("Catalog_Договор");

/** Субконто, по которым идёт группировка: их наименования нужны и для строк вне страницы. */
export function groupDims(entries: readonly CardEntry[], groupBy: readonly CardGroup[]): Dim[] {
  return entries.flatMap((e) =>
    e.dims.filter(
      (d) =>
        (groupBy.includes("counterparty") && isCounterparty(d)) ||
        (groupBy.includes("contract") && isContract(d)),
    ),
  );
}

/**
 * Сводка по ВСЕМ движениям карточки, а не по странице: суммы групп сходятся с оборотами карточки.
 * Строка без нужного субконто попадает в группу без этого поля, а не теряется. Субконто группируются
 * по ссылке: одноимённые договоры разных контрагентов («Основной договор») — разные группы, ссылка
 * есть в ответе (counterpartyRef, contractRef) и годится для отбора.
 */
export function groupEntries(
  entries: readonly CardEntry[],
  groupBy: readonly CardGroup[],
  name: (d: Dim) => string,
  code: (key?: string) => string | undefined,
): Array<Record<string, string | number>> {
  const groups = new Map<
    string,
    { fields: Record<string, string>; sortKey: string; dr: number; cr: number; count: number }
  >();
  for (const e of entries) {
    const fields: Record<string, string> = {};
    const ids: string[] = [];
    for (const g of groupBy) {
      if (g === "month" || g === "corrAccount") {
        const value = g === "month" ? e.period.slice(0, 7) : code(e.corrAccount);
        if (value) fields[g] = value;
        ids.push(value ?? "");
        continue;
      }
      const d = e.dims.find(g === "counterparty" ? isCounterparty : isContract);
      if (d) {
        fields[g] = name(d);
        fields[`${g}Ref`] = d.ref;
      }
      ids.push(d?.ref ?? "");
    }
    const k = ids.join("\u0000");
    const group = groups.get(k) ?? {
      fields,
      sortKey: `${groupBy.map((g) => fields[g] ?? "").join("\u0000")}\u0000${k}`,
      dr: 0,
      cr: 0,
      count: 0,
    };
    if (e.side === "debit") group.dr += e.cents;
    else group.cr += e.cents;
    group.count += 1;
    groups.set(k, group);
  }
  return [...groups.values()]
    .sort((a, b) => a.sortKey.localeCompare(b.sortKey))
    .map((g) => ({ ...g.fields, debit: rub(g.dr), credit: rub(g.cr), count: g.count }));
}

async function accountCard(
  conn: Connection,
  o: {
    organization?: string | undefined;
    account: string;
    from: string;
    to: string;
    counterpartyRef?: string | undefined;
    dimensionRef?: string | undefined;
    limit: number;
    offset: number;
    groupBy: readonly CardGroup[];
    summaryOnly: boolean;
  },
) {
  const t0 = Date.now();
  const org = await orgOf(conn, o.organization);
  const accounts = await resolveAccounts(conn, [o.account]);
  if (!accounts.length) throw new InputError(`Счёт "${o.account}" не найден в плане счетов «Хозрасчётный».`);
  if (accounts.length >= conn.behavior.maxRows) {
    throw new InputError(`Под префикс "${o.account}" подходит слишком много счетов. Уточните код счёта.`);
  }
  const keys = accounts.map((a) => a.key);
  const refs = [o.counterpartyRef, o.dimensionRef].filter((r): r is string => !!r);

  const [openRows, closeRows, moves] = await Promise.all([
    balanceByAccounts(conn, keys, org?.ref, addDays(o.from, -1)),
    balanceByAccounts(conn, keys, org?.ref, o.to),
    fetchMovements(conn, {
      accountKeys: keys,
      from: o.from,
      to: o.to,
      ...(org ? { orgKey: org.ref } : {}),
      ...(o.counterpartyRef ? { counterpartyRef: o.counterpartyRef } : {}),
    }),
  ]);
  const entries = cardEntries(moves.rows, new Set(keys), refs);
  const opening = netBalance(openRows, refs);
  const expectedClosing = netBalance(closeRows, refs);

  let balance = opening;
  let dr = 0;
  let cr = 0;
  const withBalance = entries.map((e) => {
    if (e.side === "debit") {
      dr += e.cents;
      balance += e.cents;
    } else {
      cr += e.cents;
      balance -= e.cents;
    }
    return { e, balance };
  });
  const groupBy = [...new Set(o.groupBy)];
  const candidates = o.summaryOnly
    ? []
    : withBalance.slice(o.offset, o.offset + Math.min(o.limit, MAX_CARD_PAGE));
  const [names, code] = await Promise.all([
    describeRefs(
      conn,
      candidates.map((x) => x.e),
      groupDims(entries, groupBy),
    ),
    codesOf(conn, accounts, groupBy.includes("corrAccount") ? entries : candidates.map((x) => x.e)),
  ]);
  const rows = candidates.map(({ e, balance: b }) => {
    const doc = names.doc(e.recorderSet, e.recorder);
    return {
      date: e.period.slice(0, 10),
      ...(e.recorderSet
        ? { document: e.recorderSet.replace(/^Document_/, ""), documentTitle: documentTitle(e.recorderSet) }
        : {}),
      ...(doc.number ? { number: doc.number } : {}),
      ...(e.recorder ? { documentRef: e.recorder } : {}),
      ...(createdAt(e.recorder) ? { documentCreatedAt: createdAt(e.recorder) } : {}),
      ...(e.content ? { content: e.content } : {}),
      account: code(e.account),
      ...(e.dims.length ? { analytics: e.dims.map(names.name) } : {}),
      ...(e.side === "debit" ? { debit: rub(e.cents) } : { credit: rub(e.cents) }),
      ...(e.corrAccount ? { corrAccount: code(e.corrAccount) } : {}),
      ...(e.corrDims.length ? { corrAnalytics: e.corrDims.map(names.name) } : {}),
      balance: rub(b),
    };
  });
  const groups = groupBy.length ? groupEntries(entries, groupBy, names.name, code) : undefined;
  if (groups && elementChars(groups) > MAX_GROUP_CHARS) {
    throw new InputError(
      `Группировка ${groupBy.join(", ")} дала ${groups.length} групп — в ответ не поместится. ` +
        "Сократите groupBy или сузьте период, отбор — counterpartyRef или dimensionRef.",
    );
  }

  const head = {
    database: conn.cfg.name,
    ...(org ? { organization: org.name } : {}),
    account: o.account,
    accounts: accounts.map((a) => `${a.code} ${a.description}`),
    period: { from: o.from, to: o.to },
    ...(refs.length
      ? {
          filter: {
            ...(o.counterpartyRef ? { counterpartyRef: o.counterpartyRef } : {}),
            ...(o.dimensionRef ? { dimensionRef: o.dimensionRef } : {}),
          },
        }
      : {}),
    openingBalance: rub(opening),
    debitTurnover: rub(dr),
    creditTurnover: rub(cr),
    closingBalance: rub(balance),
    consistent: balance === expectedClosing,
    count: entries.length,
    ...(groups ? { groups } : {}),
  };
  const shown = fitToChars(rows, JSON.stringify(head, null, 2).length);
  const nextOffset = o.offset + shown.length;
  const truncated = !o.summaryOnly && nextOffset < entries.length;

  const notes: string[] = [];
  if (truncated) {
    notes.push(
      `Показаны движения ${o.offset + 1}–${nextOffset} из ${entries.length}, итоги посчитаны по всем. ` +
        `Следующая страница: offset=${nextOffset}. Сводка без строк: summaryOnly=true и groupBy.`,
    );
  }
  if (!o.summaryOnly && o.offset > 0 && o.offset >= entries.length) {
    notes.push(`offset=${o.offset} за пределами карточки: всего движений ${entries.length}.`);
  }
  if (balance !== expectedClosing) {
    notes.push(
      `Остаток по движениям (${rub(balance)}) не совпал с сальдо регистра на конец периода ` +
        `(${rub(expectedClosing)}) — возможно, движения изменились во время выборки; повторите запрос.`,
    );
  }
  return {
    ...head,
    offset: o.offset,
    returned: shown.length,
    truncated,
    ...(truncated ? { nextOffset } : {}),
    entries: shown,
    ...(notes.length ? { note: notes.join(" ") } : {}),
    scan: { rowsScanned: moves.meta.rowsScanned, windows: moves.meta.chunks, elapsedMs: Date.now() - t0 },
  };
}

// ─── Акт сверки ──────────────────────────────────────────────────────────────

/**
 * Счета взаиморасчётов для акта сверки: 60, 62 и 76 — кроме субсчетов НДС с авансов
 * (76.АВ, 76.ВА) и прочих буквенных: у них тоже есть субконто «Контрагенты», но это не расчёты.
 */
export function reconciliationAccounts(accounts: readonly Account[]): Account[] {
  return accounts.filter((a) => /^[0-9.]+$/.test(a.code));
}

const inFavorOf = (cents: number) => (cents > 0 ? "организации" : cents < 0 ? "контрагента" : undefined);

/** Строки акта: по документу, без внутренних проводок (зачёт аванса 62.02 → 62.01 того же контрагента). */
export function reconciliationLines(entries: readonly CardEntry[]): Array<{
  recorder?: string;
  recorderSet?: string;
  date: string;
  contract?: Dim;
  debit: number;
  credit: number;
}> {
  const sides = new Map<string, Set<string>>();
  const rowKey = (e: CardEntry) => `${e.recorder}|${e.recorderSet}|${e.lineNumber}|${e.period}`;
  for (const e of entries) sides.set(rowKey(e), (sides.get(rowKey(e)) ?? new Set()).add(e.side));
  const lines = new Map<
    string,
    { recorder?: string; recorderSet?: string; date: string; contract?: Dim; debit: number; credit: number }
  >();
  for (const e of entries) {
    if (sides.get(rowKey(e))!.size === 2) continue;
    const contract = e.dims.find((d) => d.set?.startsWith("Catalog_Договор"));
    const k = `${e.recorder}|${e.recorderSet}|${contract?.ref ?? ""}`;
    const l = lines.get(k) ?? {
      ...(e.recorder ? { recorder: e.recorder } : {}),
      ...(e.recorderSet ? { recorderSet: e.recorderSet } : {}),
      date: e.period.slice(0, 10),
      ...(contract ? { contract } : {}),
      debit: 0,
      credit: 0,
    };
    if (e.side === "debit") l.debit += e.cents;
    else l.credit += e.cents;
    lines.set(k, l);
  }
  return [...lines.values()].filter((l) => l.debit !== 0 || l.credit !== 0);
}

async function reconciliation(
  conn: Connection,
  o: {
    organization?: string | undefined;
    counterpartyRef: string;
    contractRef?: string | undefined;
    from: string;
    to: string;
  },
) {
  const t0 = Date.now();
  const org = await orgOf(conn, o.organization);
  const prefixes = [
    ...accountPrefixes(conn.cfg, "receivables"),
    ...accountPrefixes(conn.cfg, "payables"),
    "76",
  ];
  const accounts = reconciliationAccounts(await resolveAccounts(conn, [...new Set(prefixes)]));
  const keys = accounts.map((a) => a.key);
  const refs = [o.counterpartyRef, o.contractRef].filter((r): r is string => !!r);

  const [openRows, closeRows, moves] = await Promise.all([
    balanceByAccounts(conn, keys, org?.ref, addDays(o.from, -1)),
    balanceByAccounts(conn, keys, org?.ref, o.to),
    fetchMovements(conn, {
      accountKeys: keys,
      from: o.from,
      to: o.to,
      counterpartyRef: o.counterpartyRef,
      ...(org ? { orgKey: org.ref } : {}),
    }),
  ]);
  const entries = cardEntries(moves.rows, new Set(keys), refs);
  const lines = reconciliationLines(entries);
  const opening = netBalance(openRows, refs);
  const closingRegister = netBalance(closeRows, refs);
  const dr = lines.reduce((s, l) => s + l.debit, 0);
  const cr = lines.reduce((s, l) => s + l.credit, 0);
  const closing = opening + dr - cr;

  const names = await describeRefs(conn, entries);
  const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
  const cpName = cpSet
    ? (await resolveNames(conn, cpSet, [o.counterpartyRef])).get(o.counterpartyRef)
    : undefined;
  if (cpSet && cpName === undefined) {
    throw new InputError(`Контрагент ${o.counterpartyRef} не найден. Ref_Key берётся из find_counterparty.`);
  }

  // Разбивка по договорам: сальдо на начало и конец из регистра, обороты — из строк акта.
  const contractOf = (r: Record<string, unknown>) =>
    [1, 2, 3]
      .map((i) => ({ ref: r[`ExtDimension${i}`], t: String(r[`ExtDimension${i}_Type`] ?? "") }))
      .find((d) => d.t.includes("Catalog_Договор"))?.ref as string | undefined;
  const perContract = new Map<string, { opening: number; closing: number; debit: number; credit: number }>();
  const pc = (k: string) => perContract.get(k) ?? { opening: 0, closing: 0, debit: 0, credit: 0 };
  for (const [rows, field] of [
    [openRows, "opening"],
    [closeRows, "closing"],
  ] as const) {
    for (const r of rows) {
      if (netBalance([r], refs) === 0) continue;
      const k = contractOf(r) ?? "";
      const v = pc(k);
      v[field] += netBalance([r], refs);
      perContract.set(k, v);
    }
  }
  for (const l of lines) {
    const k = l.contract?.ref ?? "";
    const v = pc(k);
    v.debit += l.debit;
    v.credit += l.credit;
    perContract.set(k, v);
  }
  const contractNames = new Map<string, string>();
  const contractSet = resolveEntity(CATALOGS.contracts, await conn.available());
  if (contractSet) {
    for (const [k, v] of await resolveNames(conn, contractSet, [...perContract.keys()].filter(Boolean)))
      contractNames.set(k, v);
  }

  // Акт сверки составляется между двумя юрлицами: по какой организации данные — должно быть видно.
  const orgRefs = new Set<string>([
    ...entries.map((e) => e.org).filter((x): x is string => !!x),
    ...[...openRows, ...closeRows]
      .filter((r) => netBalance([r], refs) !== 0)
      .map((r) => refOf(r["Организация_Key"]))
      .filter((x): x is string => !!x),
  ]);
  const orgNames = org
    ? new Map([[org.ref, org.name]])
    : new Map((await listOrganizations(conn)).map((x) => [x.ref, x.name]));
  const orgsInData = [...orgRefs].map((r) => orgNames.get(r) ?? r);

  const notes = [
    "Это данные организации для акта сверки: дебет — отгрузки покупателю и оплаты поставщику, кредит — " +
      "оплаты покупателя и поступления от поставщика. Документ «Акт сверки» в 1С заполняется кнопкой " +
      "формы и печатается из 1С — через OData его не создать.",
  ];
  if (closing !== closingRegister) {
    notes.push(
      `Сальдо по строкам акта (${rub(closing)}) не совпало с регистром (${rub(closingRegister)}) — повторите запрос.`,
    );
  }
  if (!org && orgsInData.length > 1) {
    notes.push(
      `В данных несколько организаций (${orgsInData.join(", ")}): акт сверки составляется по одной — ` +
        "повторите запрос с organization.",
    );
  }
  if (!lines.length && opening === 0)
    notes.push("За период нет ни движений, ни остатков по расчётам с контрагентом.");

  return {
    database: conn.cfg.name,
    ...(org ? { organization: org.name } : orgsInData.length === 1 ? { organization: orgsInData[0] } : {}),
    ...(!org && orgsInData.length > 1 ? { organizations: orgsInData } : {}),
    counterparty: cpName ?? o.counterpartyRef,
    counterpartyRef: o.counterpartyRef,
    ...(o.contractRef ? { contractRef: o.contractRef } : {}),
    period: { from: o.from, to: o.to },
    accounts: accounts
      .filter(
        (a) =>
          entries.some((e) => e.account === a.key) ||
          [...openRows, ...closeRows].some((r) => r["Account_Key"] === a.key && netBalance([r], refs) !== 0),
      )
      .map((a) => a.code),
    opening: { amount: rub(Math.abs(opening)), inFavorOf: inFavorOf(opening) ?? null },
    debitTurnover: rub(dr),
    creditTurnover: rub(cr),
    closing: { amount: rub(Math.abs(closing)), inFavorOf: inFavorOf(closing) ?? null },
    consistent: closing === closingRegister,
    count: lines.length,
    lines: lines.map((l) => {
      const doc = names.doc(l.recorderSet, l.recorder);
      return {
        date: doc.date ?? l.date,
        ...(l.recorderSet
          ? { document: l.recorderSet.replace(/^Document_/, ""), documentTitle: documentTitle(l.recorderSet) }
          : {}),
        ...(doc.number ? { number: doc.number, printNumber: printNumber(doc.number) } : {}),
        ...(l.recorder ? { documentRef: l.recorder } : {}),
        ...(l.contract ? { contract: names.name(l.contract) } : {}),
        ...(l.debit ? { debit: rub(l.debit) } : {}),
        ...(l.credit ? { credit: rub(l.credit) } : {}),
      };
    }),
    byContract: [...perContract.entries()].map(([k, v]) => ({
      contract: k ? (contractNames.get(k) ?? k) : "(без договора)",
      ...(k ? { contractRef: k } : {}),
      opening: rub(v.opening),
      debit: rub(v.debit),
      credit: rub(v.credit),
      closing: rub(v.closing),
    })),
    note: notes.join(" "),
    scan: { rowsScanned: moves.meta.rowsScanned, windows: moves.meta.chunks, elapsedMs: Date.now() - t0 },
  };
}

// ─── Откуда сложилось сальдо ─────────────────────────────────────────────────

/**
 * Точный состав остатка по документам расчётов: строки сальдо, где среди субконто есть документ.
 * complete=false — часть остатка без документа (учёт не по документам), тогда — расчёт по движениям.
 */
export function originByDocuments(
  rows: readonly Record<string, unknown>[],
  refs: readonly string[],
): { items: Array<{ doc: Dim; others: Dim[]; cents: number }>; complete: boolean } {
  const byDoc = new Map<string, { doc: Dim; others: Dim[]; cents: number }>();
  let withoutDoc = 0;
  for (const r of rows) {
    const cents = netBalance([r as never], refs);
    if (cents === 0) continue;
    const dims: Dim[] = [];
    for (const i of [1, 2, 3]) {
      const ref = refOf(r[`ExtDimension${i}`]);
      const set = String(r[`ExtDimension${i}_Type`] ?? "");
      if (ref) dims.push({ ref, ...(set ? { set: set.slice(set.lastIndexOf(".") + 1) } : {}) });
    }
    const doc = dims.find((d) => d.set?.startsWith("Document_"));
    if (!doc) {
      withoutDoc += cents;
      continue;
    }
    const cur = byDoc.get(doc.ref) ?? { doc, others: dims.filter((d) => d !== doc), cents: 0 };
    cur.cents += cents;
    byDoc.set(doc.ref, cur);
  }
  return { items: [...byDoc.values()].filter((x) => x.cents !== 0), complete: withoutDoc === 0 };
}

/**
 * Какие движения образовали остаток B, который висит сейчас. entries — все движения по счёту
 * (по возрастанию времени). Идём назад от B, пока остаток не обнулится или не сменит знак — с
 * этого движения (crossing) остаток держится непрерывно. Дальше вперёд: доля движения в остатке —
 * насколько оно подняло «нижнюю границу» остатка до даты расчёта (минимум остатка от этого
 * движения до конца). Сумма долей = |B|; это точный расчёт по истории движений, без допущений.
 * found=false — момент появления раньше переданных движений: carried — часть остатка старше их.
 */
export function persistentOrigin(
  entries: readonly CardEntry[],
  balanceCents: number,
): { found: boolean; carried: number; items: Array<{ e: CardEntry; cents: number }> } {
  const sign = balanceCents > 0 ? 1 : -1;
  const delta = (e: CardEntry) => (e.side === "debit" ? e.cents : -e.cents);
  let bal = balanceCents;
  let start = 0;
  let found = false;
  for (let i = entries.length - 1; i >= 0; i--) {
    const before = bal - delta(entries[i]!);
    if (before === 0 || Math.sign(before) !== sign) {
      start = i;
      found = true;
      bal = before;
      break;
    }
    bal = before;
  }
  // bal — остаток перед первым учитываемым движением; в нашу сторону (sign) или 0.
  const v0 = Math.max(0, sign * bal);
  const tail = entries.slice(start);
  const after: number[] = [];
  let running = bal;
  for (const e of tail) {
    running += delta(e);
    after.push(sign * running);
  }
  const suffixMin: number[] = new Array<number>(after.length);
  let m = Infinity;
  for (let i = after.length - 1; i >= 0; i--) {
    m = Math.min(m, after[i]!);
    suffixMin[i] = m;
  }
  const items: Array<{ e: CardEntry; cents: number }> = [];
  let prev = found ? 0 : Math.min(v0, suffixMin[0] ?? v0);
  for (let i = 0; i < tail.length; i++) {
    const cur = Math.max(0, suffixMin[i]!);
    if (cur > prev) items.push({ e: tail[i]!, cents: cur - prev });
    prev = Math.max(prev, cur);
  }
  return { found, carried: found ? 0 : Math.min(v0, suffixMin[0] ?? v0), items };
}

async function balanceOrigin(
  conn: Connection,
  o: {
    organization?: string | undefined;
    account: string;
    asOf: string;
    counterpartyRef?: string | undefined;
    dimensionRef?: string | undefined;
    maxYears: number;
  },
) {
  const t0 = Date.now();
  const org = await orgOf(conn, o.organization);
  const accounts = await resolveAccounts(conn, [o.account]);
  if (!accounts.length) throw new InputError(`Счёт "${o.account}" не найден в плане счетов «Хозрасчётный».`);
  const keys = accounts.map((a) => a.key);
  const refs = [o.counterpartyRef, o.dimensionRef].filter((r): r is string => !!r);
  const balRows = await balanceByAccounts(conn, keys, org?.ref, o.asOf);
  const balance = netBalance(balRows, refs);
  const head = {
    database: conn.cfg.name,
    ...(org ? { organization: org.name } : {}),
    account: o.account,
    asOf: o.asOf,
    balance: rub(balance),
    side: balance >= 0 ? "debit" : "credit",
  };
  if (balance === 0) {
    return {
      ...head,
      method: "none",
      count: 0,
      items: [],
      note: "Остатка на дату нет.",
      scan: { elapsedMs: Date.now() - t0 },
    };
  }

  // 1. Точно — по документам расчётов в самом сальдо.
  const byDocs = originByDocuments(balRows, refs);
  if (byDocs.complete && byDocs.items.length) {
    const names = await describeRefs(
      conn,
      [],
      byDocs.items.flatMap((x) => [x.doc, ...x.others]),
    );
    const items = byDocs.items
      .map((x) => {
        const info = names.doc(x.doc.set, x.doc.ref);
        return {
          date: info.date ?? "",
          document: (x.doc.set ?? "").replace(/^Document_/, ""),
          documentTitle: documentTitle(x.doc.set ?? ""),
          ...(info.number ? { number: info.number } : {}),
          documentRef: x.doc.ref,
          ...(createdAt(x.doc.ref) ? { documentCreatedAt: createdAt(x.doc.ref) } : {}),
          ...(x.others.length ? { analytics: x.others.map(names.name) } : {}),
          amount: rub(x.cents),
        };
      })
      .sort((a, b) => a.date.localeCompare(b.date));
    return {
      ...head,
      method: "documents",
      since: items[0]?.date,
      count: items.length,
      items,
      note:
        "Точный состав остатка по документам расчётов из сальдо 1С. Отрицательная сумма — аванс или оплата, " +
        "ещё не закрытая встречным документом.",
      scan: { elapsedMs: Date.now() - t0 },
    };
  }

  // 2. По истории движений — назад окнами 3, 6, 12… месяцев, пока не найдём момент появления.
  const limitDate = addDays(o.asOf, -Math.round(o.maxYears * 365.25));
  let collected: CardEntry[] = [];
  let windowEnd = o.asOf;
  let days = 91;
  let scanned = 0;
  let origin = persistentOrigin([], balance);
  while (windowEnd > limitDate) {
    const from = addDays(windowEnd, -days) < limitDate ? limitDate : addDays(windowEnd, -days);
    const moves = await fetchMovements(conn, {
      accountKeys: keys,
      from,
      to: windowEnd,
      ...(org ? { orgKey: org.ref } : {}),
      ...(o.counterpartyRef ? { counterpartyRef: o.counterpartyRef } : {}),
    });
    scanned += moves.meta.rowsScanned;
    collected = [...cardEntries(moves.rows, new Set(keys), refs), ...collected];
    origin = persistentOrigin(collected, balance);
    if (origin.found) break;
    windowEnd = addDays(from, -1);
    days *= 2;
  }
  const taken = origin.items;
  const names = await describeRefs(
    conn,
    taken.map((t) => t.e),
  );
  const code = await codesOf(
    conn,
    accounts,
    taken.map((t) => t.e),
  );
  const items = taken
    .map(({ e, cents }) => {
      const info = names.doc(e.recorderSet, e.recorder);
      return {
        date: e.period.slice(0, 10),
        ...(e.recorderSet
          ? { document: e.recorderSet.replace(/^Document_/, ""), documentTitle: documentTitle(e.recorderSet) }
          : {}),
        ...(info.number ? { number: info.number } : {}),
        ...(e.recorder ? { documentRef: e.recorder } : {}),
        ...(createdAt(e.recorder) ? { documentCreatedAt: createdAt(e.recorder) } : {}),
        ...(e.content ? { content: e.content } : {}),
        account: code(e.account),
        ...(e.corrAccount ? { corrAccount: code(e.corrAccount) } : {}),
        ...(e.side === "debit" ? { debit: rub(e.cents) } : { credit: rub(e.cents) }),
        amount: rub(cents),
      };
    })
    .sort((a, b) => a.date.localeCompare(b.date));
  const notes = [
    "Расчёт по истории движений: найден момент, с которого остаток держится непрерывно (since), и доля " +
      "каждого документа, дожившая до даты расчёта (amount; debit/credit — полная сумма движения). Движения " +
      "одного дня делят остаток по порядку записи: для счетов, закрываемых в конце месяца (20–26, 44), какие " +
      "документы не попали в закрытие, точнее покажет documentCreatedAt позже closingCreatedAt.",
  ];
  if (!origin.found) {
    notes.push(
      `Остаток появился раньше, чем ${o.maxYears} г. назад: ${rub(origin.carried)} — с более ранних периодов, ` +
        "увеличьте maxYears.",
    );
  }
  return {
    ...head,
    method: "movements",
    since: items[0]?.date,
    ...(origin.found ? {} : { olderThanSearch: rub(origin.carried) }),
    count: items.length,
    items,
    note: notes.join(" "),
    scan: { rowsScanned: scanned, elapsedMs: Date.now() - t0 },
  };
}

export function registerCardTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.accounting.get_account_card",
    {
      title: "Карточка счёта",
      description:
        "Карточка счёта за период — каждое движение с корреспонденцией: дата, документ, содержание, " +
        "Дт или Кт, корреспондирующий счёт и его аналитика, остаток после движения; сальдо на начало, " +
        "обороты, сальдо на конец (сверено с регистром, consistent). account — код или префикс " +
        "(«51», «62.01», «26»). Отбор: counterpartyRef — контрагент (Ref_Key из find_counterparty), " +
        "dimensionRef — любое другое субконто (договор, номенклатура, банковский счёт, документ " +
        "расчётов). Сторно приходит отрицательной суммой в debit/credit. documentCreatedAt — когда документ " +
        "создан (по ссылке): помогает найти документы, внесённые после закрытия месяца. Если остаток сложился " +
        "до периода — возьмите период шире. Для разбора: откуда остаток на 26, кто дал красное сальдо, как сложился долг. " +
        "Ответ постраничный: count — всего движений, returned — сколько в этом ответе; при truncated=true " +
        "читайте дальше с offset=nextOffset, пока truncated не станет false. За один ответ не больше " +
        `${MAX_CARD_PAGE} строк и примерно ${MAX_CARD_CHARS} символов, что бы ни стояло в limit. Обороты и остатки всегда по всему периоду. ` +
        "Итоги по месяцам, корреспондирующим счетам, контрагентам или договорам: groupBy, сумма групп равна " +
        "оборотам карточки; с summaryOnly=true строк нет и ответ быстрее. Для этого не нужно читать все страницы. " +
        "Не отдаёт обороты по субсчетам (это get_account_turnover) и не разбирает текст назначений платежей.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        account: z
          .string()
          .trim()
          .min(1)
          .max(20)
          .regex(ACCOUNT_CODE_RE, "Код счёта: цифры/буквы через точку, напр. 51, 60.01")
          .describe("Код счёта или префикс субсчетов"),
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
        counterpartyRef: guidField("Ref_Key контрагента (из find_counterparty)").optional(),
        dimensionRef: guidField(
          "Ref_Key другого субконто: договор, номенклатура, банковский счёт",
        ).optional(),
        limit: z
          .number()
          .int()
          .positive()
          .max(2000)
          .default(200)
          .describe(`Сколько движений показать, в одном ответе не больше ${MAX_CARD_PAGE}`),
        offset: z
          .number()
          .int()
          .min(0)
          .default(0)
          .describe("С какого движения начать (0 — с начала); значение из nextOffset прошлого ответа"),
        groupBy: z
          .array(z.enum(CARD_GROUPS))
          .max(CARD_GROUPS.length)
          .optional()
          .describe(
            "Итоги по всем движениям периода: month — месяц, corrAccount — корреспондирующий счёт, " +
              "counterparty — контрагент, contract — договор (субконто счёта карточки, с их Ref_Key). Можно сочетать",
          ),
        summaryOnly: z
          .boolean()
          .default(false)
          .describe("Только сальдо, обороты и groups, без строк движений"),
      }),
      outputSchema: accountCardResultSchema,
    },
    ({
      database,
      organization,
      account,
      from,
      to,
      counterpartyRef,
      dimensionRef,
      limit,
      offset,
      groupBy,
      summaryOnly,
    }) =>
      guard("read.accounting.get_account_card", async () => {
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        return ok(
          await accountCard(ctx.db(database), {
            organization,
            account,
            from,
            to,
            counterpartyRef,
            dimensionRef,
            limit,
            offset,
            groupBy: groupBy ?? [],
            summaryOnly,
          }),
        );
      }),
  );

  server.registerTool(
    "read.accounting.get_balance_origin",
    {
      title: "Откуда сложилось сальдо",
      description:
        "Из каких документов состоит остаток по счёту на дату и с какого времени он висит — для старых " +
        "хвостов, которых не видно в карточке за месяц. Где учёт ведётся по документам расчётов (62, 60, 76 " +
        "в БП), состав — из сальдо 1С (method = documents). Иначе — по истории движений (method = movements): " +
        "момент, с которого остаток держится непрерывно, и доля каждого документа, дожившая до даты; поиск " +
        "назад до maxYears лет. since — дата самого " +
        "старого документа остатка, documentCreatedAt — когда документ создан. Отбор: counterpartyRef — " +
        "контрагент, dimensionRef — договор, номенклатура, банковский счёт.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        account: z
          .string()
          .trim()
          .min(1)
          .max(20)
          .regex(ACCOUNT_CODE_RE, "Код счёта: цифры/буквы через точку, напр. 62.01, 26")
          .describe("Код счёта или префикс субсчетов"),
        asOf: dateField("Остаток на конец этой даты (без параметра — сегодня)").optional(),
        counterpartyRef: guidField("Ref_Key контрагента (из find_counterparty)").optional(),
        dimensionRef: guidField(
          "Ref_Key другого субконто: договор, номенклатура, банковский счёт",
        ).optional(),
        maxYears: z.number().int().min(1).max(10).default(3).describe("Как далеко назад искать по движениям"),
      }),
      outputSchema: balanceOriginResultSchema,
    },
    ({ database, organization, account, asOf, counterpartyRef, dimensionRef, maxYears }) =>
      guard("read.accounting.get_balance_origin", async () =>
        ok(
          await balanceOrigin(ctx.db(database), {
            organization,
            account,
            asOf: asOf ?? localToday(),
            counterpartyRef,
            dimensionRef,
            maxYears,
          }),
        ),
      ),
  );

  server.registerTool(
    "read.counterparty.get_reconciliation",
    {
      title: "Данные для акта сверки",
      description:
        "Взаиморасчёты с контрагентом за период для акта сверки — по данным организации: сальдо на начало " +
        "и конец «в пользу организации / контрагента», строки по документам (дата, документ, номер, договор, " +
        "дебет, кредит; printNumber — номер как в печатной форме), обороты, разбивка по договорам, организация. " +
        "Проверяются счета 60, 62 и 76 (без 76.АВ/76.ВА), в accounts — те, где есть движения или остатки; зачёт аванса " +
        "внутри расчётов с тем же контрагентом в строки не попадает, как в акте сверки 1С. counterpartyRef — " +
        "Ref_Key из find_counterparty, contractRef — только один договор. Сам документ «Акт сверки» " +
        "печатается из 1С.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: guidField("Ref_Key контрагента (из find_counterparty)"),
        contractRef: guidField("Ref_Key договора — сверка по одному договору").optional(),
        from: dateField("Дата начала периода"),
        to: dateField("Дата конца периода"),
      }),
      outputSchema: reconciliationResultSchema,
    },
    ({ database, organization, counterpartyRef, contractRef, from, to }) =>
      guard("read.counterparty.get_reconciliation", async () => {
        if (from > to) return fail(`Период задан наоборот: from (${from}) позже to (${to}).`);
        return ok(
          await reconciliation(ctx.db(database), { organization, counterpartyRef, contractRef, from, to }),
        );
      }),
  );
}
