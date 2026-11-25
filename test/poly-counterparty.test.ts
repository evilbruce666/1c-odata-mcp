import { writeFileSync } from "node:fs";
import { afterAll, afterEach, describe, expect, it, vi } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { matchesIncomingNumber, matchesOwnNumber, parseNumberQuery } from "../src/odata/doc-number.js";
import { fake1C } from "./helpers/fake-1c.js";
import { checkPolyCounterparty } from "../src/scripts/live-poly.js";

/**
 * Отбор по полиморфному «Контрагент» банковских и кассовых документов. Корпус: документы, где в «Контрагент» лежат
 * контрагенты, физлица, сотрудники и пустая ссылка; ссылки — контрагент с документами и без, физлицо, сотрудник,
 * неизвестный guid. Метрика: найденные документы и итоги совпадают с эталоном, посчитанным по данным корпуса без
 * учёта типа ссылки (как отбор на стороне коннектора), — и когда 1С принимает cast, и когда отвергает его.
 * POLY_METRICS_OUT=<файл> пишет по каждому случаю запросы, строки и байты ответов 1С — для сравнения до и после.
 */

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const CP = (i: number) => uuid("c", i);
const PERSON = (i: number) => uuid("f", i);
const EMPLOYEE = uuid("e", 1);
const UNKNOWN = uuid("d", 99);
const EMPTY = "00000000-0000-0000-0000-000000000000";
const ORG = (i: number) => uuid("a", i);

const POLY_SETS = {
  Document_ПоступлениеНаРасчетныйСчет: { kind: "bank", isIn: true },
  Document_СписаниеСРасчетногоСчета: { kind: "bank", isIn: false },
  Document_ПлатежноеПоручение: { kind: "bank", isIn: false },
  Document_ПриходныйКассовыйОрдер: { kind: "cash", isIn: true },
  Document_РасходныйКассовыйОрдер: { kind: "cash", isIn: false },
  Document_ОплатаПлатежнойКартой: { kind: "cash", isIn: true },
} as const;
type PolySet = keyof typeof POLY_SETS;
const TYPED_SET = "Document_РеализацияТоваровУслуг";
const PAYMENT_SETS: PolySet[] = [
  "Document_ПоступлениеНаРасчетныйСчет",
  "Document_СписаниеСРасчетногоСчета",
  "Document_ПриходныйКассовыйОрдер",
  "Document_РасходныйКассовыйОрдер",
];

/** Кто лежит в «Контрагент»: ссылка и её тип (пустая ссылка — без типа). CP(5) — контрагент без документов. */
const HOLDERS: Array<{ ref: string; type: string }> = [
  ...[1, 2, 3, 4].map((i) => ({ ref: CP(i), type: "StandardODATA.Catalog_Контрагенты" })),
  { ref: PERSON(1), type: "StandardODATA.Catalog_ФизическиеЛица" },
  { ref: PERSON(2), type: "StandardODATA.Catalog_ФизическиеЛица" },
  { ref: EMPLOYEE, type: "StandardODATA.Catalog_Сотрудники" },
  { ref: EMPTY, type: "" },
];
const COUNTERPARTIES = [1, 2, 3, 4, 5].map((i) => ({ Ref_Key: CP(i), Description: `Контрагент ${i}` }));
const OPERATIONS = ["ОплатаПокупателя", "ОплатаПоставщику", "ПрочееПоступление", "ПрочийРасход"];

type Doc = Record<string, unknown> & { set: string };
function corpus(): Doc[] {
  const docs: Doc[] = [];
  let n = 0;
  for (const [s, set] of Object.keys(POLY_SETS).entries()) {
    for (let i = 0; i < 40; i++) {
      n++;
      const h = HOLDERS[(i * 3 + s) % HOLDERS.length]!;
      const month = 1 + (i % 3);
      const day = 1 + ((i * 7) % 28);
      // Регистр букв guid в ответе 1С — как придётся: отбор у себя сравнивает без регистра.
      const ref = i % 11 === 0 ? h.ref.toUpperCase() : h.ref;
      docs.push({
        set,
        Ref_Key: uuid("b", n),
        Number: `0000-${String(n).padStart(6, "0")}`,
        Date: `2026-0${month}-${String(day).padStart(2, "0")}T${i % 5 === 0 ? "00:00:00" : "12:30:00"}`,
        Posted: i % 6 !== 0,
        DeletionMark: false,
        СуммаДокумента: 100 * ((i % 9) + 1) + i / 100,
        Организация_Key: ORG(1 + (i % 2)),
        Контрагент: ref,
        Контрагент_Type: h.type,
        ВидОперации: OPERATIONS[i % OPERATIONS.length],
        НомерВходящегоДокумента: i % 4 === 0 ? `ВХ-${n}` : "",
        НазначениеПлатежа: i % 3 === 0 ? "оплата по договору" : "прочее",
        СтатьяДвиженияДенежныхСредств_Key: EMPTY,
      });
    }
  }
  for (let i = 0; i < 12; i++) {
    n++;
    docs.push({
      set: TYPED_SET,
      Ref_Key: uuid("b", n),
      Number: `0000-${String(n).padStart(6, "0")}`,
      Date: `2026-0${1 + (i % 3)}-15T10:00:00`,
      Posted: true,
      DeletionMark: false,
      СуммаДокумента: 1000 + i,
      Организация_Key: ORG(1 + (i % 2)),
      Контрагент_Key: CP(1 + (i % 3)),
    });
  }
  return docs;
}

/** УТ: те же виды движения денег с типизированным Контрагент_Key — отбор на сервере голым guid, как раньше. */
const UT_SETS: Record<string, boolean> = {
  Document_ПоступлениеБезналичныхДенежныхСредств: true,
  Document_СписаниеБезналичныхДенежныхСредств: false,
  Document_ПриходныйКассовыйОрдер: true,
  Document_РасходныйКассовыйОрдер: false,
};
function utCorpus(): Doc[] {
  const docs: Doc[] = [];
  let n = 500;
  for (const [s, set] of Object.keys(UT_SETS).entries())
    for (let i = 0; i < 20; i++) {
      n++;
      docs.push({
        set,
        Ref_Key: uuid("b", n),
        Number: `УТ-${n}`,
        Date: `2026-0${1 + (i % 3)}-10T09:00:00`,
        Posted: true,
        DeletionMark: false,
        СуммаДокумента: 50 + i + s,
        Организация_Key: ORG(1),
        Контрагент_Key: CP(1 + ((i + s) % 4)),
        ХозяйственнаяОперация: "ОплатаПоставщику",
      });
    }
  return docs;
}

// ---------- Отбор $filter, как его выполняет 1С: and/or, сравнения, cast с типом, endswith/substringof. ----------

const literal = (raw: string): unknown => {
  const s = raw.trim();
  let m: RegExpExecArray | null;
  if ((m = /^cast\(guid'([0-9a-f-]{36})',\s*'([^']+)'\)$/i.exec(s)))
    return { guid: m[1]!.toLowerCase(), type: `StandardODATA.${m[2]}` };
  if ((m = /^(?:guid|datetime)'([^']*)'$/i.exec(s))) return m[1]!;
  if ((m = /^'((?:[^']|'')*)'$/.exec(s))) return m[1]!.replace(/''/g, "'");
  if (s === "true" || s === "false") return s === "true";
  return Number(s);
};
const unlike = (v: string) => v.replace(/\[(.)\]/g, "$1").replace(/\\\\/g, "\\");

/** Разбивка по разделителю верхнего уровня (вне скобок и кавычек). */
function splitTop(expr: string, sep: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let quoted = false;
  let start = 0;
  for (let i = 0; i < expr.length; i++) {
    const c = expr[i]!;
    if (c === "'") quoted = !quoted;
    else if (!quoted && c === "(") depth++;
    else if (!quoted && c === ")") depth--;
    else if (!quoted && depth === 0 && expr.startsWith(sep, i)) {
      out.push(expr.slice(start, i));
      start = i + sep.length;
      i += sep.length - 1;
    }
  }
  out.push(expr.slice(start));
  return out;
}

function evaluate(expr: string, row: Record<string, unknown>): boolean {
  const e = expr.trim();
  const ors = splitTop(e, " or ");
  if (ors.length > 1) return ors.some((x) => evaluate(x, row));
  const ands = splitTop(e, " and ");
  if (ands.length > 1) return ands.every((x) => evaluate(x, row));
  if (e.startsWith("(") && e.endsWith(")")) return evaluate(e.slice(1, -1), row);
  let m: RegExpExecArray | null;
  if ((m = /^endswith\(([^,]+),\s*'((?:[^']|'')*)'\)$/.exec(e)))
    return String(row[m[1]!.trim()] ?? "")
      .toLowerCase()
      .endsWith(m[2]!.replace(/''/g, "'").toLowerCase());
  if ((m = /^substringof\('((?:[^']|'')*)',\s*([^)]+)\)$/.exec(e)))
    return String(row[m[2]!.trim()] ?? "")
      .toLowerCase()
      .includes(unlike(m[1]!.replace(/''/g, "'")).toLowerCase());
  m = /^([\wА-Яа-яЁё]+)\s+(eq|ne|ge|le|gt|lt)\s+(.+)$/.exec(e);
  if (!m) throw new Error(`Отбор не поддержан корпусом: ${e}`);
  const [, field, op, raw] = m;
  const want = literal(raw!);
  const have = row[field!];
  if (want && typeof want === "object") {
    const w = want as { guid: string; type: string };
    const eq = String(have ?? "").toLowerCase() === w.guid && row[`${field}_Type`] === w.type;
    return op === "eq" ? eq : !eq;
  }
  const a = typeof have === "string" && typeof want === "string" ? have.toLowerCase() : have;
  const b = typeof want === "string" ? want.toLowerCase() : want;
  switch (op) {
    case "eq":
      return a === b;
    case "ne":
      return a !== b;
    case "ge":
      return (a as number) >= (b as number);
    case "le":
      return (a as number) <= (b as number);
    case "gt":
      return (a as number) > (b as number);
    default:
      return (a as number) < (b as number);
  }
}

// ---------- Поддельная 1С на корпусе: счёт запросов, строк и байтов; отказ на cast по желанию. ----------

interface Counters {
  requests: number;
  rows: number;
  bytes: number;
}

async function setup(opts: { rejectCast?: boolean; rejectIn?: string[]; ut?: boolean } = {}) {
  const docs = opts.ut ? utCorpus() : corpus();
  const fake = fake1C(
    { READ_ONLY: "true" },
    {
      ...(opts.ut ? { schema: "ut11" as const } : {}),
      data: (q) => {
        const filter = q.params.get("$filter") ?? "";
        if (q.base.startsWith("Document_")) {
          const rows = docs.filter((d) => d.set === q.base);
          return rows
            .filter((r) => !filter || evaluate(filter, r))
            .sort((x, y) => String(y["Date"]).localeCompare(String(x["Date"])))
            .map(({ set: _set, ...r }) => r);
        }
        if (q.base === "Catalog_Организации")
          return [
            { Ref_Key: ORG(1), Description: "Организация 1", ИНН: "7700000011" },
            { Ref_Key: ORG(2), Description: "Организация 2", ИНН: "7700000022" },
          ];
        if (q.base === "Catalog_Контрагенты")
          return COUNTERPARTIES.filter((c) => !filter || evaluate(filter, c));
        return undefined;
      },
    },
  );
  const counters: Counters = { requests: 0, rows: 0, bytes: 0 };
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    const decoded = decodeURIComponent(url);
    if (decoded.includes("$metadata")) return inner(input, init);
    counters.requests++;
    const rejected =
      /\$filter=[^&]*Контрагент eq cast\(/.test(decoded) &&
      (opts.rejectCast || (opts.rejectIn ?? []).some((set) => decoded.includes(`/${set}?`)));
    if (rejected)
      return new Response(
        JSON.stringify({
          "odata.error": {
            code: "-1",
            message: {
              lang: "ru",
              value:
                "Неверные параметры в операции сравнения. Нельзя сравнивать поля неограниченной длины и поля несовместимых типов.",
            },
          },
        }),
        { status: 500, headers: { "Content-Type": "application/json" } },
      );
    // Ссылка по ключу: контрагента нет — 404, как в 1С.
    const keyed = /Catalog_Контрагенты\(guid'([0-9a-f-]{36})'\)/i.exec(decoded);
    if (keyed && !COUNTERPARTIES.some((c) => c.Ref_Key === keyed[1]!.toLowerCase())) {
      return new Response(
        JSON.stringify({ "odata.error": { code: "-1", message: { lang: "ru", value: "Не найден" } } }),
        { status: 404, headers: { "Content-Type": "application/json" } },
      );
    }
    const res = await inner(input, init);
    const text = await res.clone().text();
    counters.bytes += text.length;
    try {
      const body = JSON.parse(text) as { value?: unknown[] };
      counters.rows += Array.isArray(body.value) ? body.value.length : 1;
    } catch {
      // не JSON — не строка данных
    }
    return res;
  });
  type Tools = Record<
    string,
    {
      inputSchema?: { parse: (v: unknown) => unknown };
      handler: (
        a: unknown,
        e: unknown,
      ) => Promise<{
        isError?: boolean;
        structuredContent?: Record<string, unknown>;
        content?: { text?: string }[];
      }>;
    }
  >;
  const tools = (createServer(fake.ctx) as unknown as { _registeredTools: Tools })._registeredTools;
  const call = async (name: string, args: Record<string, unknown>) => {
    const tool = tools[name]!;
    const r = await tool.handler(tool.inputSchema!.parse(args), {});
    const text = (r.content ?? []).map((c) => c.text ?? "").join("");
    if (r.isError) throw new Error(`${name}: ${text}`);
    return r.structuredContent!;
  };
  /** Вызов как в live-read: ошибка — не исключение, а err с текстом. */
  const run = async (name: string, args: Record<string, unknown>) => {
    const tool = tools[name]!;
    const r = await tool.handler(tool.inputSchema!.parse(args), {});
    return {
      err: !!r.isError,
      sc: r.structuredContent ?? {},
      text: (r.content ?? []).map((c) => c.text ?? "").join(""),
    };
  };
  return { docs, call, run, counters, fake };
}

// ---------- Эталон: отбор по данным корпуса, ссылка сравнивается без типа. ----------

const KIND_SETS: Record<string, string[]> = {
  bank: Object.keys(POLY_SETS).filter((s) => POLY_SETS[s as PolySet].kind === "bank"),
  cash: Object.keys(POLY_SETS).filter((s) => POLY_SETS[s as PolySet].kind === "cash"),
  all: [...Object.keys(POLY_SETS), TYPED_SET],
};

interface SearchCase {
  ref: string;
  kind?: "bank" | "cash";
  entitySet?: string;
  from?: string;
  to?: string;
  organization?: string;
  postedOnly?: boolean;
  minAmount?: number;
  number?: string;
}

const norm = (v: unknown) =>
  String(v ?? "")
    .replace(/[{}]/g, "")
    .toLowerCase();

function expectedSearch(docs: Doc[], c: SearchCase): string[] {
  const sets = c.entitySet ? [c.entitySet] : KIND_SETS[c.kind ?? "all"]!;
  const q = c.number ? parseNumberQuery(c.number) : undefined;
  return docs
    .filter((d) => sets.includes(d.set))
    .filter((d) => norm(d["Контрагент"] ?? d["Контрагент_Key"]) === c.ref)
    .filter((d) => !c.from || String(d["Date"]) >= `${c.from}T00:00:00`)
    .filter((d) => !c.to || String(d["Date"]) <= `${c.to}T23:59:59`)
    .filter((d) => !c.organization || d["Организация_Key"] === c.organization)
    .filter((d) => !c.postedOnly || d["Posted"] === true)
    .filter((d) => c.minAmount === undefined || (d["СуммаДокумента"] as number) >= c.minAmount)
    .filter(
      (d) => !q || matchesOwnNumber(d["Number"], q) || matchesIncomingNumber(d["НомерВходящегоДокумента"], q),
    )
    .map((d) => String(d["Ref_Key"]))
    .sort();
}

const REFS: Array<[string, string]> = [
  ["контрагент с документами", CP(1)],
  ["контрагент без документов", CP(5)],
  ["физлицо", PERSON(1)],
  ["сотрудник", EMPLOYEE],
  ["неизвестный guid", UNKNOWN],
];

const SEARCH_VARIANTS: Array<[string, Omit<SearchCase, "ref">]> = [
  ["банк", { kind: "bank" }],
  ["касса", { kind: "cash" }],
  ["все виды", {}],
  ["один вид", { entitySet: "Document_ПоступлениеНаРасчетныйСчет" }],
  ["период", { kind: "bank", from: "2026-02-01", to: "2026-02-28" }],
  ["организация и проведённые", { kind: "cash", organization: ORG(1), postedOnly: true }],
  ["сумма от", { kind: "bank", minAmount: 500 }],
  ["номер", { kind: "bank", number: "1" }],
];

/**
 * Пометка в ответе: откат на отбор у себя виден всегда. 1С отвергла cast — «не приняла»; ссылка не контрагент — «не
 * найдена среди контрагентов». Контрагент при работающем cast — без пометок.
 */
function expectNotes(text: string, rejectCast: boolean, ref: string) {
  const counterparty = COUNTERPARTIES.some((c) => c.Ref_Key === ref);
  if (rejectCast) expect(text).toContain("не приняла серверный отбор");
  else if (!counterparty) expect(text).toContain("не найдена среди контрагентов");
  else expect(text).toBe("");
}

const metrics: Array<Record<string, unknown>> = [];
afterAll(() => {
  if (process.env.POLY_METRICS_OUT)
    writeFileSync(process.env.POLY_METRICS_OUT, JSON.stringify(metrics, null, 1));
});

/** Все документы по страницам: limit 5, пока есть nextOffset. */
async function searchAll(
  call: (n: string, a: Record<string, unknown>) => Promise<Record<string, unknown>>,
  c: SearchCase,
) {
  const refs: string[] = [];
  const notes = new Set<string>();
  let offset = 0;
  for (let page = 0; page < 60; page++) {
    const { ref, ...rest } = c;
    const s = await call("read.document.search_documents", {
      ...rest,
      counterpartyRef: ref,
      limit: 5,
      offset,
    });
    for (const r of s["rows"] as { ref: string }[]) refs.push(r.ref);
    if (typeof s["counterpartyFilter"] === "string") notes.add(s["counterpartyFilter"]);
    if (s["nextOffset"] === undefined) return { refs, notes, capped: s["cappedSets"] };
    offset = s["nextOffset"] as number;
  }
  throw new Error("страницы не кончились");
}

describe("корпус", () => {
  it("у контрагента, физлица и сотрудника есть документы в каждом виде, у CP(5) и неизвестного — нет", () => {
    const docs = corpus();
    for (const set of Object.keys(POLY_SETS))
      for (const ref of [CP(1), PERSON(1), EMPLOYEE]) {
        const n = docs.filter((d) => d.set === set && norm(d["Контрагент"]) === ref).length;
        expect(n, `${set} · ${ref}`).toBeGreaterThan(2);
      }
    for (const ref of [CP(5), UNKNOWN]) expect(docs.filter((d) => norm(d["Контрагент"]) === ref)).toEqual([]);
    // Документы с заглавными буквами в guid есть: отбор у себя сравнивает без регистра.
    expect(docs.some((d) => /[A-F]/.test(String(d["Контрагент"])))).toBe(true);
  });
});

describe("search_documents: полиморфный «Контрагент» — те же документы, что эталон", () => {
  for (const rejectCast of [false, true])
    for (const [refName, ref] of REFS)
      for (const [variantName, variant] of SEARCH_VARIANTS) {
        const label = `${rejectCast ? "1С отвергает cast" : "1С принимает cast"} · ${refName} · ${variantName}`;
        it(label, async () => {
          const { docs, call, counters } = await setup({ rejectCast });
          const c: SearchCase = { ref, ...variant };
          const got = await searchAll(call, c);
          expect([...got.refs].sort()).toEqual(expectedSearch(docs, c));
          expect(new Set(got.refs).size, "документ повторился на страницах").toBe(got.refs.length);
          expect(got.capped, "выборка упёрлась в предел").toBeUndefined();
          expectNotes([...got.notes].join(" "), rejectCast, ref);
          metrics.push({ tool: "search", case: label, ...counters });
        });
      }
});

// ---------- get_payments_breakdown ----------

interface PayCase {
  ref: string;
  direction: "in" | "out" | "both";
  groupBy: "operation" | "month" | "counterparty" | "total";
  organization?: string;
}

function expectedPayments(docs: Doc[], c: PayCase) {
  // Движение денег — поступления и списания по счёту, ПКО и РКО; платёжное поручение и оплата картой сюда не входят.
  const sets = PAYMENT_SETS.filter((s) => {
    const isIn = POLY_SETS[s].isIn;
    return c.direction === "both" || (c.direction === "in") === isIn;
  });
  const rows = docs
    .filter((d) => sets.includes(d.set))
    .filter((d) => d["Posted"] === true)
    .filter((d) => norm(d["Контрагент"]) === c.ref)
    .filter((d) => !c.organization || d["Организация_Key"] === c.organization)
    .filter((d) => String(d["Date"]) >= "2026-01-01T00:00:00" && String(d["Date"]) <= "2026-03-31T23:59:59");
  const cents = (d: Doc) => Math.round((d["СуммаДокумента"] as number) * 100);
  const inflow = rows.filter((d) => POLY_SETS[d.set as PolySet].isIn).reduce((s, d) => s + cents(d), 0);
  const outflow = rows.filter((d) => !POLY_SETS[d.set as PolySet].isIn).reduce((s, d) => s + cents(d), 0);
  return { documents: rows.length, inflow: inflow / 100, outflow: outflow / 100 };
}

const PAY_VARIANTS: Array<[string, Omit<PayCase, "ref">]> = [
  ["оба направления, по операциям", { direction: "both", groupBy: "operation" }],
  ["приход, по месяцам", { direction: "in", groupBy: "month" }],
  ["расход, по контрагентам", { direction: "out", groupBy: "counterparty" }],
  ["итог по организации", { direction: "both", groupBy: "total", organization: ORG(2) }],
];

describe("get_payments_breakdown: полиморфный «Контрагент» — те же итоги, что эталон", () => {
  for (const rejectCast of [false, true])
    for (const [refName, ref] of REFS)
      for (const [variantName, variant] of PAY_VARIANTS) {
        const label = `${rejectCast ? "1С отвергает cast" : "1С принимает cast"} · ${refName} · ${variantName}`;
        it(label, async () => {
          const { docs, call, counters } = await setup({ rejectCast });
          const c: PayCase = { ref, ...variant };
          const s = await call("read.analytics.get_payments_breakdown", {
            from: "2026-01-01",
            to: "2026-03-31",
            counterpartyRef: ref,
            direction: c.direction,
            groupBy: c.groupBy,
            ...(c.organization ? { organization: c.organization } : {}),
          });
          const want = expectedPayments(docs, c);
          expect({ documents: s["documents"], inflow: s["inflow"], outflow: s["outflow"] }).toEqual(want);
          const groups = s["groups"] as { count: number; sum: number }[];
          expect(groups.reduce((n, g) => n + g.count, 0)).toBe(want.documents);
          expectNotes(String(s["note"] ?? ""), rejectCast, ref);
          metrics.push({ tool: "payments", case: label, groups: groups.length, ...counters });
        });
      }
});

describe("УТ: типизированный Контрагент_Key — отбор по контрагенту не теряется", () => {
  it.each([CP(1), CP(5)])("get_payments_breakdown, %s", async (ref) => {
    const { docs, call } = await setup({ ut: true });
    const s = await call("read.analytics.get_payments_breakdown", {
      from: "2026-01-01",
      to: "2026-03-31",
      counterpartyRef: ref,
      groupBy: "total",
    });
    const rows = docs.filter((d) => d["Контрагент_Key"] === ref);
    const sum = (isIn: boolean) =>
      rows
        .filter((d) => UT_SETS[d.set] === isIn)
        .reduce((n, d) => n + Math.round((d["СуммаДокумента"] as number) * 100), 0) / 100;
    expect({ documents: s["documents"], inflow: s["inflow"], outflow: s["outflow"] }).toEqual({
      documents: rows.length,
      inflow: sum(true),
      outflow: sum(false),
    });
    expect(s["note"]).toBeUndefined();
  });
});

describe("живая проверка live-read на корпусе (src/scripts/live-poly.ts)", () => {
  const live = async (opts: {
    rejectCast?: boolean;
    dropFirstRow?: boolean;
    failEmptySearch?: boolean;
    duplicateRow?: boolean;
    endlessPages?: boolean;
  }) => {
    const { run, counters, fake } = await setup(opts);
    const checks: Array<[string, boolean, string]> = [];
    const lines: string[] = [];
    // Испорченный инструмент: теряет первый документ в каждом ответе поиска.
    const tool = async (name: string, args: Record<string, unknown>) => {
      // Листание не кончается: проверка упирается в предел страниц и не должна считаться совпавшей (ревью).
      if (opts.endlessPages && name === "read.document.search_documents")
        return { err: false, sc: { rows: [], nextOffset: Number(args["offset"] ?? 0) + 1 }, text: "" };
      const r = await run(name, args);
      if (name !== "read.document.search_documents" || !Array.isArray(r.sc["rows"])) return r;
      const rows = r.sc["rows"] as unknown[];
      if (opts.dropFirstRow) r.sc["rows"] = rows.slice(1);
      // Поиск упал у ссылки без документов: пустой ответ совпал бы с пустым эталоном (ревью).
      if (opts.failEmptySearch && rows.length === 0)
        return { err: true, sc: {}, text: "Превышен таймаут запроса" };
      if (opts.duplicateRow && rows.length) r.sc["rows"] = [...rows, rows[0]];
      return r;
    };
    await checkPolyCounterparty(
      fake.ctx,
      fake.conn.cfg.name,
      tool,
      (name, ok, detail = "") => checks.push([name, ok, detail]),
      (s) => lines.push(s),
      () => counters.requests,
    );
    return { checks, lines };
  };

  it.each([false, true])("1С отвергает cast: %s — все ссылки совпали с эталоном", async (rejectCast) => {
    const { checks, lines } = await live({ rejectCast });
    // Три частых контрагента, физлицо или сотрудник, контрагент без документов, неизвестный guid.
    expect(checks.length).toBe(6);
    expect(checks.filter(([, ok]) => !ok)).toEqual([]);
    expect(lines.join("\n")).toContain(rejectCast ? "cast отвергнут" : "cast принят");
    expect(lines.join("\n")).toContain("расхождений: 0");
  });

  it.each([
    ["потерял документ", { dropFirstRow: true }],
    ["упал на ссылке без документов", { failEmptySearch: true }],
    ["повторил документ", { duplicateRow: true }],
  ] as const)("ловит испорченный поиск: %s — проверка не проходит", async (_name, opts) => {
    const { checks } = await live(opts);
    expect(checks.filter(([, ok]) => !ok).length).toBeGreaterThan(0);
  });

  it("листание не кончается — не проходит ни одна проверка, и у ссылок без документов тоже", async () => {
    // Пустой список у ссылки без документов совпадает с пустым эталоном: без отметки «не дочитан» был бы ✓ (ревью).
    const { checks } = await live({ endlessPages: true });
    expect(checks.length).toBe(6);
    expect(checks.filter(([, ok]) => ok)).toEqual([]);
    expect(checks.find(([name]) => name.includes("неизвестный guid"))![2]).toContain("не дочитан");
  });
});

describe("крупные страницы при серверном отборе (живой замер: 139 запросов вместо 45)", () => {
  it("контрагент с 300 документами в одном виде, limit 500 — один запрос к 1С, как при отборе у себя", async () => {
    const { call } = await setup();
    // Документы корпуса — по 40 на вид; добавляем крупного контрагента через тот же отбор.
    const set = "Document_ПоступлениеНаРасчетныйСчет";
    const extra = Array.from({ length: 300 }, (_, i) => ({
      Ref_Key: uuid("9", i + 1),
      Number: `БОЛ-${i}`,
      Date: "2026-02-10T10:00:00",
      Posted: true,
      DeletionMark: false,
      СуммаДокумента: 10,
      Организация_Key: ORG(1),
      Контрагент: CP(4),
      Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
    }));
    const inner = globalThis.fetch;
    let documentRequests = 0;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
      const decoded = decodeURIComponent(url);
      if (decoded.includes(`${set}?`)) documentRequests++;
      if (!decoded.includes(`${set}?`) || !decoded.includes("cast(")) return inner(input, init);
      const res = await inner(input, init);
      const body = (await res.json()) as { value: Record<string, unknown>[] };
      const skip = Number(/\$skip=(\d+)/.exec(decoded)?.[1] ?? 0);
      const top = Number(/\$top=(\d+)/.exec(decoded)?.[1] ?? 10_000);
      const rows = [...body.value.filter((r) => !String(r["Number"]).startsWith("БОЛ-")), ...extra]
        .filter((r) => r["Контрагент"] === CP(4))
        .slice(skip, skip + top);
      return new Response(JSON.stringify({ value: rows }), {
        headers: { "Content-Type": "application/json" },
      });
    });
    const s = await call("read.document.search_documents", {
      entitySet: set,
      counterpartyRef: CP(4),
      limit: 500,
    });
    // Страница короче limit — предел размера ответа; читается одна выборка из 1С.
    expect((s["rows"] as unknown[]).length).toBeGreaterThan(100);
    // К документам — один запрос (названия контрагента и организации читаются отдельно, из справочников).
    expect(documentRequests).toBe(1);
  });
});

describe("1С отвергает cast только в одном виде (ревью: документы физлица в остальных видах терялись молча)", () => {
  const rejectIn = ["Document_СписаниеСРасчетногоСчета"];
  for (const [refName, ref] of REFS) {
    it(`search_documents, все виды · ${refName}`, async () => {
      const { docs, call } = await setup({ rejectIn });
      const c: SearchCase = { ref };
      const got = await searchAll(call, c);
      expect([...got.refs].sort()).toEqual(expectedSearch(docs, c));
      expect([...got.notes].join(" ")).toContain("не приняла серверный отбор");
    });
    it(`get_payments_breakdown, расход · ${refName}`, async () => {
      const { docs, call } = await setup({ rejectIn });
      const c: PayCase = { ref, direction: "out", groupBy: "operation" };
      const s = await call("read.analytics.get_payments_breakdown", {
        from: "2026-01-01",
        to: "2026-03-31",
        counterpartyRef: ref,
        direction: "out",
        groupBy: "operation",
      });
      const want = expectedPayments(docs, c);
      expect({ documents: s["documents"], inflow: s["inflow"], outflow: s["outflow"] }).toEqual(want);
    });
  }
});

describe("найденное типизированным Контрагент_Key — тоже доказательство, что ссылка — контрагент", () => {
  it("cast в банке и кассе ничего не нашёл, реализации нашлись — без проверки справочника и без отката", async () => {
    // Ветка typedCounterparty в recheckPoly: без неё ушёл бы лишний GET Catalog_Контрагенты(guid) (ревью).
    const { call } = await setup();
    const inner = globalThis.fetch;
    let keyedChecks = 0;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
      const decoded = decodeURIComponent(url);
      if (/Catalog_Контрагенты\(guid'/.test(decoded)) keyedChecks++;
      // Серверный отбор по полиморфному полю ничего не находит: документов контрагента в банке и кассе нет.
      if (decoded.includes("cast(guid'"))
        return new Response(JSON.stringify({ value: [] }), {
          headers: { "Content-Type": "application/json" },
        });
      return inner(input, init);
    });
    const s = await call("read.document.search_documents", { counterpartyRef: CP(1), limit: 500 });
    const rows = s["rows"] as Array<{ entitySet: string }>;
    expect(rows.length).toBeGreaterThan(0);
    expect(rows.every((r) => r.entitySet === TYPED_SET)).toBe(true);
    expect(keyedChecks).toBe(0);
    expect(String(s["counterpartyFilter"] ?? "")).not.toContain("не найдена среди контрагентов");
  });
});
