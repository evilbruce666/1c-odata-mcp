/**
 * Живая проверка серверного отбора по полиморфному «Контрагент» (cast) — часть live-read, только чтение.
 *
 * Эталон — прежний способ: все банковские и кассовые документы периода без отбора по контрагенту, отбор здесь, по
 * guid без учёта типа ссылки (так работала 0.15.0). Ссылки берутся из самих документов: самые частые контрагенты,
 * физлицо или сотрудник, контрагент без документов в периоде, неизвестный guid. По каждой: набор документов
 * search_documents и итоги get_payments_breakdown совпадают с эталоном; видно, приняла ли 1С cast и сколько
 * запросов ушло.
 */
import type { Connection, ServerContext } from "../context.js";
import { CATALOGS, DOC_FIELDS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { collectDocuments } from "../odata/aggregate.js";
import { buildQuery } from "../odata/query.js";

type Json = Record<string, unknown>;
type Run = (name: string, args: Json) => Promise<{ err: boolean; sc: Json; text: string }>;
type Check = (name: string, ok: boolean, detail?: string) => void;

const POLY = "Контрагент";
const CP_TYPE = "StandardODATA.Catalog_Контрагенты";
/** Виды search_documents, где лежат банковские и кассовые документы. */
const SEARCH_KINDS = ["bank", "cash"] as const;
const SEARCH_KEYS = ["bankIn", "bankOut", "paymentOrder", "cashIn", "cashOut", "cardPayment"] as const;
/** Движение денег get_payments_breakdown: поступления и списания по счёту, ПКО и РКО. */
const PAYMENT_KEYS: Record<string, boolean> = { bankIn: true, cashIn: true, bankOut: false, cashOut: false };

const norm = (v: unknown) =>
  String(v ?? "")
    .replace(/[{}]/g, "")
    .trim()
    .toLowerCase();
const cents = (v: unknown) => Math.round(Number(v ?? 0) * 100);
const iso = (d: Date) => d.toISOString().slice(0, 10);

interface PolyDoc {
  set: string;
  ref: string;
  holder: string;
  holderType: string;
  posted: boolean;
  amountCents: number;
}

async function propsOf(conn: Connection, set: string): Promise<Set<string>> {
  return new Set(((await conn.getMetadata()).entities.get(set)?.properties ?? []).map((p) => p.name));
}

/** Виды документов с полиморфным «Контрагент» (рядом Контрагент_Type, без Контрагент_Key). */
async function polySets(conn: Connection, keys: readonly string[]): Promise<string[]> {
  const available = await conn.available();
  const out: string[] = [];
  for (const key of keys) {
    const set = resolveEntity(DOCUMENTS[key as keyof typeof DOCUMENTS], available);
    if (!set || out.includes(set)) continue;
    const props = await propsOf(conn, set);
    if (props.has(POLY) && props.has(`${POLY}_Type`) && !props.has(DOC_FIELDS.counterparty)) out.push(set);
  }
  return out;
}

export async function checkPolyCounterparty(
  ctx: ServerContext,
  db: string,
  run: Run,
  check: Check,
  line: (s: string) => void,
  requests: () => number,
): Promise<void> {
  const conn = ctx.db(db);
  const searchSets = await polySets(conn, SEARCH_KEYS);
  if (searchSets.length === 0) {
    line(`  ${db}: полиморфного «Контрагент» в банке и кассе нет, проверка cast пропущена`);
    return;
  }
  const to = iso(new Date());
  const from = iso(new Date(Date.now() - 365 * 86_400_000));

  // Эталон: прежний способ — всё за период, отбор здесь.
  const before = requests();
  const docs: PolyDoc[] = [];
  for (const set of searchSets) {
    const { rows } = await collectDocuments(conn, set, {
      dateField: DOC_FIELDS.date,
      from,
      to,
      select: [DOC_FIELDS.ref, DOC_FIELDS.posted, DOC_FIELDS.amount, POLY, `${POLY}_Type`],
    });
    for (const r of rows)
      docs.push({
        set,
        ref: norm(r[DOC_FIELDS.ref]),
        holder: norm(r[POLY]),
        holderType: String(r[`${POLY}_Type`] ?? ""),
        posted: r[DOC_FIELDS.posted] === true,
        amountCents: cents(r[DOC_FIELDS.amount]),
      });
  }
  const oracleRequests = requests() - before;

  // Ссылки: частые контрагенты, физлицо или сотрудник, контрагент без документов, неизвестный guid.
  const freq = new Map<string, { n: number; type: string }>();
  for (const d of docs)
    if (d.holder && !/^0{8}-/.test(d.holder))
      freq.set(d.holder, { n: (freq.get(d.holder)?.n ?? 0) + 1, type: d.holderType });
  const byFreq = [...freq].sort((a, b) => b[1].n - a[1].n);
  const refs: Array<[string, string]> = byFreq
    .filter(([, v]) => v.type === CP_TYPE)
    .slice(0, 3)
    .map(([ref]) => ["контрагент", ref]);
  const other = byFreq.find(([, v]) => v.type && v.type !== CP_TYPE);
  if (other) refs.push([other[1].type.replace("StandardODATA.Catalog_", "").toLowerCase(), other[0]]);
  const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
  if (cpSet) {
    const page = await conn.client.getCollection(`${cpSet}${buildQuery({ select: ["Ref_Key"], top: 200 })}`);
    const idle = (page.value ?? []).map((r) => norm(r["Ref_Key"])).find((r) => r && !freq.has(r));
    if (idle) refs.push(["контрагент без документов", idle]);
  }
  refs.push(["неизвестный guid", "d1d1d1d1-0000-4000-8000-000000000001"]);
  if (!other) line(`  ${db}: в документах периода нет физлиц и сотрудников — откат для них не проверен`);

  const paymentSets = new Map<string, boolean>();
  const available = await conn.available();
  for (const [key, isIn] of Object.entries(PAYMENT_KEYS)) {
    const set = resolveEntity(DOCUMENTS[key as keyof typeof DOCUMENTS], available);
    if (set) paymentSets.set(set, isIn);
  }
  const paymentsComparable = [...paymentSets.keys()].every((s) => searchSets.includes(s));

  let toolRequests = 0;
  let mismatches = 0;
  const casts = new Set<string>();
  for (const [label, ref] of refs) {
    const want = docs.filter((d) => d.holder === ref);
    // search_documents: все страницы, только полиморфные виды.
    const got: string[] = [];
    const notes: string[] = [];
    const start = requests();
    for (const kind of SEARCH_KINDS) {
      let offset = 0;
      for (let guard = 0; guard < 200; guard++) {
        const r = await run("read.document.search_documents", {
          database: db,
          kind,
          counterpartyRef: ref,
          from,
          to,
          limit: 500,
          offset,
        });
        if (r.err) {
          notes.push(`ошибка: ${r.text.slice(0, 120)}`);
          break;
        }
        for (const row of (r.sc["rows"] as Json[] | undefined) ?? [])
          if (searchSets.includes(String(row["entitySet"]))) got.push(norm(row["ref"]));
        if (typeof r.sc["counterpartyFilter"] === "string") notes.push(r.sc["counterpartyFilter"]);
        if (r.sc["nextOffset"] === undefined) break;
        offset = r.sc["nextOffset"] as number;
      }
    }
    const wantRefs = want.map((d) => d.ref).sort();
    const sameDocs = JSON.stringify([...new Set(got)].sort()) === JSON.stringify(wantRefs);

    // get_payments_breakdown: проведённые, итоги в копейках.
    let samePayments = true;
    let paymentsDetail = "";
    if (paymentsComparable) {
      const r = await run("read.analytics.get_payments_breakdown", {
        database: db,
        from,
        to,
        counterpartyRef: ref,
        groupBy: "total",
      });
      const posted = want.filter((d) => d.posted && paymentSets.has(d.set));
      const inflow = posted.filter((d) => paymentSets.get(d.set)).reduce((s, d) => s + d.amountCents, 0);
      const outflow = posted.filter((d) => !paymentSets.get(d.set)).reduce((s, d) => s + d.amountCents, 0);
      samePayments =
        !r.err &&
        r.sc["documents"] === posted.length &&
        cents(r.sc["inflow"]) === inflow &&
        cents(r.sc["outflow"]) === outflow;
      if (typeof r.sc["note"] === "string") notes.push(r.sc["note"]);
      paymentsDetail = r.err
        ? `платежи: ошибка ${r.text.slice(0, 80)}`
        : `платежи ${String(r.sc["documents"])} док. (эталон ${posted.length})`;
    }
    toolRequests += requests() - start;
    const text = notes.join(" ");
    casts.add(text.includes("не приняла") ? "отвергнут" : "принят");
    if (!sameDocs || !samePayments) mismatches++;
    check(
      `${db}: cast «Контрагент» — ${label}`,
      sameDocs && samePayments,
      `документов ${new Set(got).size} (эталон ${wantRefs.length})${paymentsDetail ? `, ${paymentsDetail}` : ""}` +
        `${text ? ` | ${text.slice(0, 140)}` : ""}`,
    );
  }
  line(
    `  ${db}: cast ${[...casts].join("/")}; эталон прежним способом — ${docs.length} строк за ${oracleRequests} ` +
      `запросов на все ссылки сразу, новый способ — ${toolRequests} запросов на ${refs.length} ссылок ` +
      `(${SEARCH_KINDS.length} поиска${paymentsComparable ? " и отчёт по платежам" : ""} на ссылку); ` +
      `расхождений: ${mismatches}`,
  );
}
