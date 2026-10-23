import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField, dateField } from "./_shared.js";
import { and, cmp, odataGuid } from "../odata/query.js";
import { accountPrefixes } from "../config/accounting.js";
import { CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import {
  balanceByAccounts,
  readByRefs,
  resolveAccounts,
  resolveNames,
  num,
  type Account,
} from "../odata/accounting.js";
import { collectDocuments, fetchAllForAggregation } from "../odata/aggregate.js";
import { requireEntity } from "../odata/publication.js";
import { resolveOrganization } from "../odata/orgs.js";
import {
  addDays,
  AGE_BUCKETS,
  ageCounterparty,
  contractTermDays,
  daysBetween,
  localToday,
  OVERDUE_BUCKETS,
  refOf,
  settlementRows,
  type AgeBucket,
  type AgedDocument,
  type DocInfo,
  type OverdueBucket,
  type SettlementRow,
  type SettlementSide,
} from "../odata/settlements-aging.js";
import {
  cashBalanceResultSchema,
  getCreditorsResultSchema,
  getDebtorsResultSchema,
  unpaidInvoicesResultSchema,
} from "../schemas/output.js";
import type { ODataEntity } from "../types/odata.js";
import { documentTitle, negativeCashHint } from "../odata/names.js";

// Деньги копим в целых копейках (float-сложение тысяч сумм даёт дрейф).
const toCents = (v: unknown): number => Math.round(num(v) * 100);
const rub = (c: number): number => Math.round(c) / 100;

/** Сколько документов показывать у одного контрагента в разбивке по срокам. */
const DOCS_PER_COUNTERPARTY = 10;

async function orgKeyOf(conn: Connection, organization?: string): Promise<{ key?: string; name?: string }> {
  if (!organization) return {};
  const org = await resolveOrganization(conn, organization);
  return { key: org.ref, name: org.name };
}

async function propertiesOf(conn: Connection, entitySet: string): Promise<Set<string>> {
  const em = (await conn.getMetadata()).entities.get(entitySet);
  return new Set(em?.properties.map((p) => p.name) ?? []);
}

const ymd = (v: unknown): string | undefined =>
  typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v) ? v.slice(0, 10) : undefined;

/** Дата и номер документов расчётов — пачками по виду документа. */
async function documentInfo(conn: Connection, rows: readonly SettlementRow[]): Promise<Map<string, DocInfo>> {
  const bySet = new Map<string, string[]>();
  for (const r of rows) {
    if (!r.document || !r.documentSet || r.cents <= 0) continue;
    bySet.set(r.documentSet, [...(bySet.get(r.documentSet) ?? []), r.document]);
  }
  const out = new Map<string, DocInfo>();
  for (const [set, refs] of bySet) {
    const found = await readByRefs(conn, set, refs, [DOC_FIELDS.ref, DOC_FIELDS.date, DOC_FIELDS.number]);
    for (const [ref, d] of found) {
      const date = ymd(d[DOC_FIELDS.date]);
      if (!date) continue;
      const number = typeof d[DOC_FIELDS.number] === "string" ? String(d[DOC_FIELDS.number]).trim() : "";
      out.set(ref, { date, ...(number ? { number } : {}) });
    }
  }
  return out;
}

/** Сроки оплаты и наименования договоров. termsPublished=false — полей срока в базе нет. */
async function contractInfo(
  conn: Connection,
  rows: readonly SettlementRow[],
): Promise<{ terms: Map<string, number | undefined>; names: Map<string, string>; termsPublished: boolean }> {
  const bySet = new Map<string, string[]>();
  for (const r of rows) {
    if (!r.contract || !r.contractSet || r.cents <= 0) continue;
    bySet.set(r.contractSet, [...(bySet.get(r.contractSet) ?? []), r.contract]);
  }
  const terms = new Map<string, number | undefined>();
  const names = new Map<string, string>();
  let termsPublished = bySet.size === 0;
  for (const [set, refs] of bySet) {
    const props = await propertiesOf(conn, set);
    const hasTerms = props.has("УстановленСрокОплаты") && props.has("СрокОплаты");
    termsPublished ||= hasTerms;
    const select = ["Ref_Key", "Description", ...(hasTerms ? ["УстановленСрокОплаты", "СрокОплаты"] : [])];
    for (const [ref, c] of await readByRefs(conn, set, refs, select)) {
      names.set(ref, String(c["Description"] ?? ""));
      terms.set(ref, hasTerms ? contractTermDays(c) : undefined);
    }
  }
  return { terms, names, termsPublished };
}

interface SettlementsOptions {
  organization?: string | undefined;
  asOf?: string | undefined;
  limit: number;
  aging: boolean;
  overdueOnly: boolean;
  defaultTermDays?: number | undefined;
}

const docOut = (d: AgedDocument, contractNames: ReadonlyMap<string, string>) => ({
  ...(d.entitySet
    ? { document: d.entitySet.replace(/^Document_/, ""), documentTitle: documentTitle(d.entitySet) }
    : {}),
  ...(d.number ? { number: d.number } : {}),
  ...(d.date ? { date: d.date } : {}),
  ...(d.ref ? { ref: d.ref } : {}),
  ...(d.contract ? { contract: contractNames.get(d.contract) || d.contract } : {}),
  amount: rub(d.cents),
  status: d.status,
  ...(d.termDays !== undefined ? { termDays: d.termDays, termSource: d.termSource } : {}),
  ...(d.dueDate ? { dueDate: d.dueDate } : {}),
  ...(d.overdueDays !== undefined ? { overdueDays: d.overdueDays } : {}),
  ...(d.ageDays !== undefined ? { ageDays: d.ageDays } : {}),
});

const bucketsOut = <K extends string>(keys: readonly K[], b: Record<K, number>) =>
  Object.fromEntries(keys.map((k) => [k, rub(b[k])])) as Record<K, number>;

/**
 * Дебиторка (62) или кредиторка (60) по контрагентам из сальдо регистра Хозрасчетный.
 * С aging — ещё и по срокам: просрочка по документам расчётов от срока оплаты в договоре.
 */
async function settlementsReport(conn: Connection, side: SettlementSide, o: SettlementsOptions) {
  const t0 = Date.now();
  const org = await orgKeyOf(conn, o.organization);
  const accounts = await resolveAccounts(
    conn,
    accountPrefixes(conn.cfg, side === "receivable" ? "receivables" : "payables"),
  );
  const raw = await balanceByAccounts(
    conn,
    accounts.map((a) => a.key),
    org.key,
    o.asOf,
  );
  const rows = settlementRows(raw, side);

  const byCp = new Map<string, SettlementRow[]>();
  for (const r of rows) byCp.set(r.counterparty, [...(byCp.get(r.counterparty) ?? []), r]);
  const net = new Map([...byCp].map(([cp, rs]) => [cp, rs.reduce((s, r) => s + r.cents, 0)]));
  const positive = [...net].filter(([, c]) => c > 0);
  const totalCents = positive.reduce((s, [, c]) => s + c, 0);

  const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
  const head = {
    database: conn.cfg.name,
    organization: org.name,
    ...(o.asOf ? { asOf: o.asOf } : {}),
    accounts: accounts.map((a) => `${a.code} ${a.description}`),
  };

  if (!o.aging) {
    const top = positive.sort((a, b) => b[1] - a[1]).slice(0, o.limit);
    const names = cpSet
      ? await resolveNames(
          conn,
          cpSet,
          top.map(([ref]) => ref),
        )
      : new Map<string, string>();
    return {
      head,
      totalCents,
      list: top.map(([ref, c]) => ({ counterparty: names.get(ref) ?? ref, ref, amount: rub(c) })),
      extra: {},
      scan: { rowsScanned: raw.length, elapsedMs: Date.now() - t0 },
    };
  }

  // Срок считаем только у тех, кто в итоге должен (сальдо > 0): у остальных аванс перекрывает долг.
  const asOf = o.asOf ?? localToday();
  const debtRows = positive.flatMap(([cp]) => byCp.get(cp)!);
  const [docs, contracts] = await Promise.all([documentInfo(conn, debtRows), contractInfo(conn, debtRows)]);
  const aged = positive.map(([cp]) => ({
    cp,
    a: ageCounterparty(byCp.get(cp)!, docs, contracts.terms, asOf, o.defaultTermDays),
  }));

  const sum = (pick: (x: (typeof aged)[number]["a"]) => number) => aged.reduce((s, x) => s + pick(x.a), 0);
  const totalBuckets = Object.fromEntries(
    OVERDUE_BUCKETS.map((k) => [k, sum((a) => a.buckets[k])]),
  ) as Record<OverdueBucket, number>;
  const totalAge = Object.fromEntries(AGE_BUCKETS.map((k) => [k, sum((a) => a.ageBuckets[k])])) as Record<
    AgeBucket,
    number
  >;
  const withoutDocument = sum((a) => a.noDocumentCents);
  const noTerm = sum((a) => a.noTermCents);

  // «Кто просрочил», а сроков в договорах нет: пустой список ничего не скажет — показываем долги
  // старше 60 дней от даты документа и помечаем это (fallback).
  const old = (a: (typeof aged)[number]["a"]) => a.ageBuckets["61-90"] + a.ageBuckets["90+"];
  const noneOverdue = aged.every((x) => x.a.overdueCents === 0);
  const fallback = o.overdueOnly && noneOverdue && aged.some((x) => x.a.noTermCents > 0 && old(x.a) > 0);
  const selected = (
    fallback
      ? aged.filter((x) => old(x.a) > 0).sort((x, y) => old(y.a) - old(x.a))
      : o.overdueOnly
        ? aged.filter((x) => x.a.overdueCents > 0).sort((x, y) => y.a.overdueCents - x.a.overdueCents)
        : aged.sort((x, y) => y.a.netCents - x.a.netCents)
  ).slice(0, o.limit);
  const names = cpSet
    ? await resolveNames(
        conn,
        cpSet,
        selected.map((x) => x.cp),
      )
    : new Map<string, string>();

  const notes: string[] = [];
  if (withoutDocument > 0 && withoutDocument === sum((a) => a.debtCents)) {
    notes.push(
      "Расчёты в базе ведутся без документов расчётов (субконто «Документы расчётов» пустое), поэтому срок " +
        "долга не определить: вся сумма — в withoutDocument.",
    );
  } else if (withoutDocument > 0) {
    notes.push(
      `Долг ${rub(withoutDocument)} без документа расчётов — его возраст не определить (withoutDocument).`,
    );
  }
  if (!contracts.termsPublished) {
    notes.push("Поля срока оплаты в справочнике договоров не опубликованы — просрочку не посчитать.");
  } else if (noTerm > 0) {
    notes.push(
      `В договорах на ${rub(noTerm)} срок оплаты не задан (нет флага «Установлен срок оплаты»): эти суммы — ` +
        "в noTerm, с возрастом от даты документа (ageDays). Если у компании общий срок, передайте defaultTermDays.",
    );
  }
  if (fallback) {
    notes.unshift(
      "Просрочку не посчитать: в договорах не задан срок оплаты. Вместо неё показаны контрагенты с долгом " +
        "старше 60 дней от даты документа (fallback = byAge), по убыванию такого долга.",
    );
  }
  const futureDated = sum((a) => a.futureDatedCents);
  const stale = sum((a) => a.staleCents);
  const staleCps = aged.filter((x) => x.a.staleCents > 0).length;
  if (stale > 0) {
    notes.push(
      `Долг старше года — ${rub(stale)} у ${staleCps} контрагентов (olderThanYear). Такие суммы часто не ` +
        "реальный долг, а незакрытые взаиморасчёты: не сделан зачёт аванса, корректировка или списание. " +
        "Проверьте актом сверки или карточкой счёта.",
    );
  }
  if (futureDated > 0) {
    notes.push(
      `Документы с датой позже ${asOf} на ${rub(futureDated)} (status future_dated) — вероятно, ошибка в дате ` +
        "документа; в возраст и просрочку они не входят.",
    );
  }
  notes.push(
    "overdue, notDue, noTerm и byAge считаются по документам долга ДО вычета авансов (debt); amount — чистый " +
      "долг после авансов, поэтому их сумма может быть больше amount.",
  );
  if (sum((a) => a.advanceCents) > 0) {
    notes.push(
      side === "receivable"
        ? "advances — авансы покупателя, ещё не закрытые отгрузкой; amount = долг − авансы."
        : "advances — наши авансы поставщику, ещё не закрытые поступлением; amount = долг − авансы.",
    );
  }

  return {
    head,
    totalCents,
    list: selected.map(({ cp, a }) => ({
      counterparty: names.get(cp) ?? cp,
      ref: cp,
      amount: rub(a.netCents),
      debt: rub(a.debtCents),
      advances: rub(a.advanceCents),
      overdue: rub(a.overdueCents),
      notDue: rub(a.notDueCents),
      noTerm: rub(a.noTermCents),
      withoutDocument: rub(a.noDocumentCents),
      ...(a.futureDatedCents ? { futureDated: rub(a.futureDatedCents) } : {}),
      ...(a.staleCents ? { olderThanYear: rub(a.staleCents) } : {}),
      ...(a.oldestDate ? { oldestDocumentDate: a.oldestDate } : {}),
      overdueByDays: bucketsOut(OVERDUE_BUCKETS, a.buckets),
      byAge: bucketsOut(AGE_BUCKETS, a.ageBuckets),
      ...(a.maxOverdueDays !== undefined ? { maxOverdueDays: a.maxOverdueDays } : {}),
      documents: a.documents.slice(0, DOCS_PER_COUNTERPARTY).map((d) => docOut(d, contracts.names)),
      ...(a.documents.length > DOCS_PER_COUNTERPARTY
        ? { documentsMore: a.documents.length - DOCS_PER_COUNTERPARTY }
        : {}),
    })),
    extra: {
      aging: {
        asOf,
        ...(o.defaultTermDays !== undefined ? { defaultTermDays: o.defaultTermDays } : {}),
        totalOverdue: rub(sum((a) => a.overdueCents)),
        overdueByDays: bucketsOut(OVERDUE_BUCKETS, totalBuckets),
        byAge: bucketsOut(AGE_BUCKETS, totalAge),
        totalNotDue: rub(sum((a) => a.notDueCents)),
        totalNoTerm: rub(noTerm),
        totalWithoutDocument: rub(withoutDocument),
        ...(futureDated ? { totalFutureDated: rub(futureDated) } : {}),
        ...(stale ? { totalOlderThanYear: rub(stale), counterpartiesOlderThanYear: staleCps } : {}),
        ...(fallback ? { fallback: "byAge" } : {}),
        counterpartiesOverdue: aged.filter((x) => x.a.overdueCents > 0).length,
      },
      ...(notes.length ? { note: notes.join(" ") } : {}),
    },
    scan: { rowsScanned: raw.length, elapsedMs: Date.now() - t0 },
  };
}

const agingFields = {
  aging: z
    .boolean()
    .default(false)
    .describe(
      "Разбивка по срокам: просрочено / в срок / срок не задан, просрочка по дням (1-30, 31-60, 61-90, 90+), " +
        "возраст долга от даты документа (byAge: 0-30 … 90+) и документы расчётов с датой оплаты по договору",
    ),
  overdueOnly: z
    .boolean()
    .default(false)
    .describe("Только контрагенты с просрочкой, по убыванию просрочки (включает aging)"),
  defaultTermDays: z
    .number()
    .int()
    .min(0)
    .max(3650)
    .optional()
    .describe(
      "Срок оплаты в днях для договоров, где он не задан. Только если пользователь его назвал — не угадывать",
    ),
};

function settlementsTool(
  server: McpServer,
  ctx: ServerContext,
  side: SettlementSide,
  def: { name: string; title: string; description: string; totalKey: string; listKey: string },
): void {
  server.registerTool(
    def.name,
    {
      title: def.title,
      description: def.description,
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        asOf: dateField("Дата сальдо — на конец этой даты (без параметра — текущее)").optional(),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько контрагентов вернуть"),
        ...agingFields,
      }),
      outputSchema: side === "receivable" ? getDebtorsResultSchema : getCreditorsResultSchema,
    },
    ({ database, organization, asOf, limit, aging, overdueOnly, defaultTermDays }) =>
      guard(def.name, async () => {
        const r = await settlementsReport(ctx.db(database), side, {
          organization,
          asOf,
          limit,
          aging: aging || overdueOnly,
          overdueOnly,
          defaultTermDays,
        });
        return ok({
          ...r.head,
          [def.totalKey]: rub(r.totalCents),
          count: r.list.length,
          [def.listKey]: r.list,
          ...r.extra,
          scan: r.scan,
        });
      }),
  );
}

// ─── Неоплаченные счета покупателям ──────────────────────────────────────────

const INVOICE_PAYMENTS = ["AccumulationRegister_ОплатаСчетов_RecordType"];
/** Оплата бывает раньше даты счёта (счёт выписали после предоплаты) — берём записи с запасом. */
const PAYMENT_LOOKBACK_DAYS = 90;
const INVOICE_STATUSES = ["open", "all"] as const;

type InvoiceStatus = "unpaid" | "partial" | "paid" | "overpaid";

export function invoiceStatus(amountCents: number, paidCents: number): InvoiceStatus {
  if (paidCents <= 0) return "unpaid";
  if (paidCents < amountCents) return "partial";
  return paidCents === amountCents ? "paid" : "overpaid";
}

/** Сальдо 62 по контрагентам (копейки; > 0 — покупатель должен). */
async function customerBalances(conn: Connection, orgKey?: string): Promise<Map<string, number>> {
  const accounts = await resolveAccounts(conn, accountPrefixes(conn.cfg, "receivables"));
  const raw = await balanceByAccounts(
    conn,
    accounts.map((a) => a.key),
    orgKey,
  );
  const out = new Map<string, number>();
  for (const r of settlementRows(raw, "receivable"))
    out.set(r.counterparty, (out.get(r.counterparty) ?? 0) + r.cents);
  return out;
}

async function unpaidInvoices(
  conn: Connection,
  o: {
    organization?: string | undefined;
    from: string;
    to: string;
    status: (typeof INVOICE_STATUSES)[number];
    limit: number;
  },
) {
  const t0 = Date.now();
  const org = await orgKeyOf(conn, o.organization);
  const invSet = await requireEntity(conn, DOCUMENTS.customerInvoice, "Документ «Счёт на оплату покупателю»");
  const paySet = await requireEntity(conn, INVOICE_PAYMENTS, "Регистр накопления «Оплата счетов»");
  const orgFilter = org.key ? cmp(DOC_FIELDS.organization, "eq", odataGuid(org.key)) : undefined;

  // Счёт в БП можно не проводить: отбираем все, кроме помеченных на удаление.
  const invoices = await collectDocuments(conn, invSet, {
    baseFilter: and(orgFilter, cmp(DOC_FIELDS.deletionMark, "eq", "false")) || undefined,
    dateField: DOC_FIELDS.date,
    from: o.from,
    to: o.to,
    select: [DOC_FIELDS.ref, DOC_FIELDS.number, DOC_FIELDS.date, DOC_FIELDS.counterparty, DOC_FIELDS.amount],
  });

  const today = localToday();
  const payments = await fetchAllForAggregation(
    conn,
    paySet,
    {
      filter:
        and(
          orgFilter,
          cmp("Period", "ge", `datetime'${addDays(o.from, -PAYMENT_LOOKBACK_DAYS)}T00:00:00'`),
        ) || undefined,
      select: ["СчетНаОплату", "Сумма", "Active"],
    },
    `оплаты счетов с ${addDays(o.from, -PAYMENT_LOOKBACK_DAYS)}`,
  );
  const paid = new Map<string, number>();
  for (const p of payments.rows) {
    const ref = refOf(p["СчетНаОплату"]);
    if (!ref || p["Active"] === false) continue;
    paid.set(ref, (paid.get(ref) ?? 0) + toCents(p["Сумма"]));
  }

  // Отгрузка по счёту — реализации со ссылкой на него. Ссылку ставят не всегда.
  const shipped = new Map<string, number>();
  let shipmentsLinked = 0;
  let shipmentsScanned = 0;
  const salesSet = resolveEntity(DOCUMENTS.sales, await conn.available());
  const invoiceLink = "СчетНаОплатуПокупателю_Key";
  if (salesSet && (await propertiesOf(conn, salesSet)).has(invoiceLink)) {
    const sales = await collectDocuments(conn, salesSet, {
      baseFilter: and(orgFilter, cmp(DOC_FIELDS.posted, "eq", "true")) || undefined,
      dateField: DOC_FIELDS.date,
      from: o.from,
      to: today > o.to ? today : o.to,
      select: [invoiceLink, DOC_FIELDS.amount],
    });
    shipmentsScanned = sales.meta.rowsScanned;
    for (const s of sales.rows) {
      const ref = refOf(s[invoiceLink]);
      if (!ref) continue;
      shipmentsLinked += 1;
      shipped.set(ref, (shipped.get(ref) ?? 0) + toCents(s[DOC_FIELDS.amount]));
    }
  }

  const all = invoices.rows.map((r: ODataEntity) => {
    const ref = String(r[DOC_FIELDS.ref]);
    const amount = toCents(r[DOC_FIELDS.amount]);
    const p = paid.get(ref) ?? 0;
    const s = shipped.get(ref) ?? 0;
    const date = ymd(r[DOC_FIELDS.date]) ?? "";
    return {
      ref,
      number: String(r[DOC_FIELDS.number] ?? "").trim(),
      date,
      counterpartyRef: refOf(r[DOC_FIELDS.counterparty]),
      amount,
      paid: p,
      shipped: s,
      status: invoiceStatus(amount, p),
      ageDays: date ? daysBetween(date, today) : undefined,
    };
  });
  const isOpen = (i: (typeof all)[number]) => i.status === "unpaid" || i.status === "partial";
  const picked = o.status === "open" ? all.filter(isOpen) : all;
  picked.sort((a, b) => a.date.localeCompare(b.date) || a.number.localeCompare(b.number));
  const shown = picked.slice(0, o.limit);

  const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
  const names = cpSet
    ? await resolveNames(
        conn,
        cpSet,
        shown.map((i) => i.counterpartyRef ?? ""),
      )
    : new Map<string, string>();

  // Сверка с 62: платёж без ссылки на счёт его не закрывает, но долг покупателя гасит.
  const balance = await customerBalances(conn, org.key);
  const openShown = shown.filter(isOpen);
  const noDebt = openShown.filter((i) => (balance.get(i.counterpartyRef ?? "") ?? 0) <= 0).length;

  const notes = [
    "Оплачено — по регистру «Оплата счетов»: учитываются платежи, в которых указан этот счёт. Платёж " +
      "без ссылки на счёт его не закрывает — так же считает статус оплаты в 1С.",
  ];
  if (noDebt > 0) {
    notes.push(
      `У ${noDebt} неоплаченных счетов покупатель по счёту 62 сейчас ничего не должен (customerBalance ≤ 0). ` +
        "Это либо счёт на предоплату без отгрузки, либо оплата пришла без ссылки на счёт — по данным " +
        "1С не различить, уточните у пользователя.",
    );
  }
  notes.push(
    shipmentsLinked === 0
      ? "Ни одна реализация за период не ссылается на счёт — отгрузку по счёту не определить (поля shipped нет)."
      : `Отгружено (shipped) — по реализациям со ссылкой на счёт; ссылка есть у ${shipmentsLinked} из ` +
          `${shipmentsScanned} реализаций за период. Нет поля shipped — отгрузка не определена, а не «не отгружено».`,
  );
  if (picked.length > shown.length)
    notes.push(`Показано ${shown.length} из ${picked.length} — увеличьте limit.`);

  const sum = (f: (i: (typeof picked)[number]) => number) => picked.reduce((s, i) => s + f(i), 0);
  return {
    database: conn.cfg.name,
    organization: org.name,
    period: { from: o.from, to: o.to },
    status: o.status,
    count: picked.length,
    totalAmount: rub(sum((i) => i.amount)),
    totalPaid: rub(sum((i) => i.paid)),
    totalUnpaid: rub(sum((i) => Math.max(0, i.amount - i.paid))),
    invoices: shown.map((i) => ({
      number: i.number,
      date: i.date,
      ref: i.ref,
      counterparty: (i.counterpartyRef && names.get(i.counterpartyRef)) || i.counterpartyRef || "",
      ...(i.counterpartyRef ? { counterpartyRef: i.counterpartyRef } : {}),
      amount: rub(i.amount),
      paid: rub(i.paid),
      unpaid: rub(Math.max(0, i.amount - i.paid)),
      // Отгрузку видно только по реализациям со ссылкой на счёт: нет ссылки — поля нет (не «0»).
      ...(i.shipped > 0 ? { shipped: rub(i.shipped) } : {}),
      status: i.status,
      ...(isOpen(i) ? { customerBalance: rub(balance.get(i.counterpartyRef ?? "") ?? 0) } : {}),
      ...(i.ageDays !== undefined ? { ageDays: i.ageDays } : {}),
    })),
    note: notes.join(" "),
    scan: {
      documentsScanned: invoices.meta.rowsScanned + shipmentsScanned,
      rowsScanned: payments.meta.rowsScanned,
      windows: invoices.meta.chunks,
      elapsedMs: Date.now() - t0,
    },
  };
}

// ─── Деньги на счетах и в кассе ──────────────────────────────────────────────

type CashKind = "cash" | "bank" | "currency" | "special" | "transit" | "other";

/** Вид денег по коду счёта плана БП: 50 касса, 51 р/с, 52 валютные, 55 спецсчета, 57 в пути. */
export function cashKind(code: string): CashKind {
  const top = code.split(".")[0];
  switch (top) {
    case "50":
      return "cash";
    case "51":
      return "bank";
    case "52":
      return "currency";
    case "55":
      return "special";
    case "57":
      return "transit";
    default:
      return "other";
  }
}

async function cashBalance(
  conn: Connection,
  o: { organization?: string | undefined; asOf?: string | undefined },
) {
  const t0 = Date.now();
  const org = await orgKeyOf(conn, o.organization);
  const accounts = await resolveAccounts(conn, accountPrefixes(conn.cfg, "cash"));
  const byKey = new Map<string, Account>(accounts.map((a) => [a.key, a]));
  const raw = await balanceByAccounts(
    conn,
    accounts.map((a) => a.key),
    org.key,
    o.asOf,
  );

  // Строка ответа = счёт × банковский счёт (или другое первое субконто) × валюта.
  type Line = {
    account: Account;
    holder?: string;
    holderSet?: string;
    currency?: string;
    cents: number;
    cur: number;
  };
  const lines = new Map<string, Line>();
  for (const r of raw) {
    const account = byKey.get(String(r["Account_Key"] ?? ""));
    if (!account) continue;
    const holder = refOf(r["ExtDimension1"]);
    const t = r["ExtDimension1_Type"];
    const holderSet = typeof t === "string" ? t.slice(t.lastIndexOf(".") + 1) : undefined;
    const currency = refOf(r["Валюта_Key"]);
    const k = `${account.key}|${holder ?? ""}|${currency ?? ""}`;
    const l = lines.get(k) ?? {
      account,
      ...(holder && holderSet?.startsWith("Catalog_") ? { holder, holderSet } : {}),
      ...(currency ? { currency } : {}),
      cents: 0,
      cur: 0,
    };
    l.cents += toCents(r["СуммаBalanceDr"]) - toCents(r["СуммаBalanceCr"]);
    l.cur += toCents(r["ВалютнаяСуммаBalanceDr"]) - toCents(r["ВалютнаяСуммаBalanceCr"]);
    lines.set(k, l);
  }

  const holderNames = new Map<string, string>();
  const bySet = new Map<string, string[]>();
  for (const l of lines.values()) {
    if (l.holder && l.holderSet) bySet.set(l.holderSet, [...(bySet.get(l.holderSet) ?? []), l.holder]);
  }
  for (const [set, refs] of bySet)
    for (const [k, v] of await resolveNames(conn, set, refs)) holderNames.set(k, v);
  const curSet = resolveEntity(CATALOGS.currencies, await conn.available());
  const currencies = [...lines.values()].map((l) => l.currency ?? "");
  const curNames = curSet ? await resolveNames(conn, curSet, currencies) : new Map<string, string>();

  const present = [...lines.values()].filter((l) => l.cents !== 0 || l.cur !== 0);
  present.sort((a, b) => a.account.code.localeCompare(b.account.code) || b.cents - a.cents);
  const total = (kinds: CashKind[]) =>
    present.filter((l) => kinds.includes(cashKind(l.account.code))).reduce((s, l) => s + l.cents, 0);

  const negative = present.filter((l) => l.cents < 0);
  const lineLabel = (l: (typeof present)[number]) =>
    `${l.account.code}${l.holder ? ` «${holderNames.get(l.holder) || l.holder}»` : " без банковского счёта"} ${rub(l.cents)}`;
  const notes: string[] = [
    "available — касса, расчётные и валютные счета (50, 51, 52); special — специальные счета и депозиты (55); " +
      "inTransit — переводы в пути и эквайринг (57), деньги ещё не на счёте.",
  ];
  if (negative.length) {
    notes.push(
      "Отрицательные остатки — проверить в учёте: " +
        negative.map((l) => `${lineLabel(l)} — ${negativeCashHint(l.account.code, !l.holder)}`).join("; ") +
        ". Итоги available / special / inTransit — чистые, с учётом этих строк.",
    );
  }

  return {
    database: conn.cfg.name,
    organization: org.name,
    ...(o.asOf ? { asOf: o.asOf } : {}),
    accounts: accounts.map((a) => `${a.code} ${a.description}`),
    available: rub(total(["cash", "bank", "currency"])),
    special: rub(total(["special"])),
    inTransit: rub(total(["transit"])),
    balances: present.map((l) => ({
      account: l.account.code,
      accountName: l.account.description,
      kind: cashKind(l.account.code),
      ...(l.holder ? { holder: holderNames.get(l.holder) || l.holder, holderRef: l.holder } : {}),
      amount: rub(l.cents),
      ...(l.cents < 0 ? { hint: negativeCashHint(l.account.code, !l.holder) } : {}),
      ...(l.currency && l.cur !== 0
        ? { currency: curNames.get(l.currency) || l.currency, currencyAmount: rub(l.cur) }
        : {}),
    })),
    note: notes.join(" "),
    scan: { rowsScanned: raw.length, elapsedMs: Date.now() - t0 },
  };
}

export function registerDebtTools(server: McpServer, ctx: ServerContext): void {
  settlementsTool(server, ctx, "receivable", {
    name: "read.analytics.get_debtors",
    title: "Дебиторская задолженность",
    totalKey: "totalReceivable",
    listKey: "debtors",
    description:
      "Кто и сколько должен компании: сальдо счёта 62 (расчёты с покупателями) из регистра " +
      "бухгалтерии Хозрасчетный, сгруппированное по контрагентам. Дебетовое сальдо = долг клиента, " +
      "кредитовое = полученные авансы (вычитается). Можно ограничить организацией. По умолчанию " +
      "берётся текущее сальдо; параметром asOf=YYYY-MM-DD можно получить дебиторку на конец " +
      "указанной даты (для аудита/исторических отчётов). Возвращает только долг > 0. " +
      "aging=true — по срокам: просрочено / в срок / срок не задан, просрочка по дням, возраст долга " +
      "от даты документа (byAge — работает и без срока в договоре) и документы расчётов; срок оплаты = дата документа + дни из договора (флаг «Установлен срок оплаты»). " +
      "overdueOnly=true — «кто просрочил оплату», по убыванию просрочки (без сроков в договорах — долги старше " +
      "60 дней, fallback). olderThanYear — долг старше года: часто незакрытые взаиморасчёты, а не реальный долг.",
  });

  settlementsTool(server, ctx, "payable", {
    name: "read.analytics.get_creditors",
    title: "Кредиторская задолженность",
    totalKey: "totalPayable",
    listKey: "creditors",
    description:
      "Кому и сколько должна компания: сальдо счёта 60 (расчёты с поставщиками) из регистра " +
      "бухгалтерии Хозрасчетный по контрагентам. Кредитовое сальдо = наш долг, дебетовое = наши " +
      "авансы поставщику (вычитаются). Возвращает только долг > 0. asOf=YYYY-MM-DD — на конец даты. " +
      "aging=true — по срокам: что уже пора платить (просрочено), что в срок, где срок не задан, с " +
      "документами поступления и датой оплаты по договору; overdueOnly=true — только просроченные. " +
      "olderThanYear — долг старше года: часто незакрытые взаиморасчёты, а не реальный долг. " +
      "Расчёты по счёту 76 (аренда, связь и прочие) сюда не входят — для них ОСВ по счёту 76.",
  });

  server.registerTool(
    "read.analytics.get_unpaid_invoices",
    {
      title: "Неоплаченные счета покупателям",
      description:
        "Счета на оплату покупателям за период и их оплата: сумма счёта, оплачено (по регистру «Оплата " +
        "счетов» — платежи со ссылкой на счёт), остаток, отгружено по счёту (реализации со ссылкой на " +
        "счёт), статус unpaid / partial / paid / overpaid, сколько дней счёту. status=open (по умолчанию) — " +
        "неоплаченные и частично оплаченные, с текущим сальдо покупателя по 62 (customerBalance) для " +
        "сверки; all — все. Счета старше 90 дней видны только с явным from. Период — по дате счёта; без from — последние 90 дней.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        from: dateField("Счета с даты (без параметра — 90 дней назад)").optional(),
        to: dateField("Счета по дату (без параметра — сегодня)").optional(),
        status: z.enum(INVOICE_STATUSES).default("open").describe("open | all"),
        limit: z.number().int().positive().max(1000).default(100).describe("Сколько счетов вернуть"),
      }),
      outputSchema: unpaidInvoicesResultSchema,
    },
    ({ database, organization, from, to, status, limit }) =>
      guard("read.analytics.get_unpaid_invoices", async () => {
        const end = to ?? localToday();
        const start = from ?? addDays(end, -90);
        if (start > end) return fail(`Период задан наоборот: from (${start}) позже to (${end}).`);
        return ok(
          await unpaidInvoices(ctx.db(database), { organization, from: start, to: end, status, limit }),
        );
      }),
  );

  server.registerTool(
    "read.analytics.get_cash_balance",
    {
      title: "Деньги на счетах и в кассе",
      description:
        "Сколько денег сейчас: остатки счетов 50 (касса), 51 (расчётные счета), 52 (валютные), 55 " +
        "(спецсчета, депозиты), 57 (переводы в пути, эквайринг) из регистра Хозрасчетный — по каждому " +
        "банковскому счёту и кассе, итоги available (50+51+52), special (55), inTransit (57). Для валютных " +
        "счетов — сумма в валюте. asOf=YYYY-MM-DD — остатки на конец даты; без него — текущие.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        asOf: dateField("Дата остатков — на конец этой даты (без параметра — текущие)").optional(),
      }),
      outputSchema: cashBalanceResultSchema,
    },
    ({ database, organization, asOf }) =>
      guard("read.analytics.get_cash_balance", async () =>
        ok(await cashBalance(ctx.db(database), { organization, asOf })),
      ),
  );
}
