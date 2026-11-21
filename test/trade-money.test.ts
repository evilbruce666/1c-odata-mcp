import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Денежные документы Управления торговлей: ПКО, РКО, поступление и списание безналичных, заявка на расходование. Состав
 * — как у документов, которые приняла и провела копия рабочей УТ 11.5 (21.11.2026) по каждой поддержанной операции.
 * Проведение сверяется по регистру денег (касса или счёт документа, сумма, направление) и расчётам с партнёром.
 */

const g = (p: string, n: number) => `${p.repeat(8)}-0000-0000-0000-${String(n).padStart(12, "0")}`;
const ORG = g("b", 1);
const CASH = g("c", 1);
const CASH2 = g("c", 2);
const BANK = g("d", 1);
const BANK2 = g("d", 2);
const CP = g("e", 1);
const PARTNER = g("e", 2);
const CONTRACT = g("e", 3);
const PERSON = g("e", 4);
const CF = g("f", 1);
const INCOME = g("f", 2);
const EXPENSE = g("f", 3);
const ASSET = g("f", 4);
const RUB = g("a", 9);

const TABLES: Record<string, Record<string, unknown>[]> = {
  Catalog_Организации: [{ Ref_Key: ORG, Description: "ТЕСТ", Code: "1" }],
  Catalog_Кассы: [
    { Ref_Key: CASH, Description: "Касса ТЕСТ", Code: "1", Owner_Key: ORG, DeletionMark: false },
    { Ref_Key: CASH2, Description: "Касса 2", Code: "2", Owner_Key: ORG, DeletionMark: false },
  ],
  Catalog_БанковскиеСчетаОрганизаций: [
    { Ref_Key: BANK, Description: "Основной", Code: "1", Owner_Key: ORG, DeletionMark: false },
    { Ref_Key: BANK2, Description: "Второй", Code: "2", Owner_Key: ORG, DeletionMark: false },
  ],
  Catalog_Контрагенты: [{ Ref_Key: CP, Description: "ZZ-TEST", Code: "1", Партнер_Key: PARTNER }],
  Catalog_ФизическиеЛица: [{ Ref_Key: PERSON, Description: "Подотчётник", Code: "1" }],
  Catalog_СтатьиДвиженияДенежныхСредств: [{ Ref_Key: CF, Description: "Оплата", Code: "1" }],
  ChartOfCharacteristicTypes_СтатьиДоходов: [{ Ref_Key: INCOME, Description: "Прочие доходы", Code: "1" }],
  ChartOfCharacteristicTypes_СтатьиРасходов: [{ Ref_Key: EXPENSE, Description: "Прочие расходы", Code: "1" }],
  ChartOfCharacteristicTypes_СтатьиАктивовПассивов: [{ Ref_Key: ASSET, Description: "Займ", Code: "1" }],
  Catalog_Валюты: [{ Ref_Key: RUB, Description: "RUB", Code: "643" }],
};

function pick(base: string, rows: Record<string, unknown>[], filter: string): Record<string, unknown>[] {
  const owner = /Owner_Key eq guid'([0-9a-f-]{36})'/i.exec(filter)?.[1];
  if (owner) return rows.filter((r) => r["Owner_Key"] === owner);
  const refs = [...filter.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());
  if (refs.length) return rows.filter((r) => refs.includes(String(r["Ref_Key"]).toLowerCase()));
  const part = /substringof\('((?:[^']|'')*)',\s*Description\)/i.exec(filter);
  if (part) return rows.filter((r) => String(r["Description"]).includes(part[1]!.replace(/''/g, "'")));
  const s = /(Code|Description) eq '((?:[^']|'')*)'/.exec(filter);
  if (s) return rows.filter((r) => String(r[s[1]!]) === s[2]!.replace(/''/g, "'"));
  void base;
  return rows;
}

const MONEY = {
  Document_ПриходныйКассовыйОрдер: { reg: "ДенежныеСредстваНаличные", acc: "Касса_Key", dir: "Receipt" },
  Document_РасходныйКассовыйОрдер: { reg: "ДенежныеСредстваНаличные", acc: "Касса_Key", dir: "Expense" },
  Document_ПоступлениеБезналичныхДенежныхСредств: {
    reg: "ДенежныеСредстваБезналичные",
    acc: "БанковскийСчет_Key",
    dir: "Receipt",
  },
  Document_СписаниеБезналичныхДенежныхСредств: {
    reg: "ДенежныеСредстваБезналичные",
    acc: "БанковскийСчет_Key",
    dir: "Expense",
  },
  Document_ЗаявкаНаРасходованиеДенежныхСредств: { reg: "ДенежныеСредстваКВыплате", acc: "", dir: "Receipt" },
} as const;
type MoneySet = keyof typeof MONEY;

/** Как проводит 1С: деньги по кассе или счёту и расчёты с партнёром (по данным копии рабочей базы). */
function movesOf(set: MoneySet, doc: Record<string, unknown>): Record<string, Record<string, unknown>[]> {
  const m = MONEY[set];
  const sum = Number(doc["СуммаДокумента"]);
  const out: Record<string, Record<string, unknown>[]> = {
    [m.reg]: [{ RecordType: m.dir, Сумма: sum, ...(m.acc ? { [m.acc]: doc[m.acc] } : {}), Active: true }],
  };
  const op = doc["ХозяйственнаяОперация"];
  if (op === "ПоступлениеОплатыОтКлиента")
    out["РасчетыСКлиентами"] = [
      { RecordType: "Expense", Сумма: 0, КОплате: sum },
      { RecordType: "Expense", Сумма: sum },
    ];
  if (op === "ОплатаПоставщику" && set !== "Document_ЗаявкаНаРасходованиеДенежныхСредств")
    out["РасчетыСПоставщиками"] = [{ RecordType: "Receipt", Сумма: sum }];
  if (op === "ВозвратОплатыКлиенту" && set !== "Document_ЗаявкаНаРасходованиеДенежныхСредств")
    out["РасчетыСКлиентами"] = [{ RecordType: "Receipt", Сумма: sum }];
  return out;
}

function setup(
  o: {
    schema?: FakeSchema;
    spoil?: (reg: string, rows: Record<string, unknown>[]) => Record<string, unknown>[];
    keepOnUnpost?: boolean;
  } = {},
) {
  const posts: { set: string; body: Record<string, unknown> }[] = [];
  const patches: Record<string, unknown>[] = [];
  const docs = new Map<
    string,
    {
      set: MoneySet;
      body: Record<string, unknown>;
      posted: boolean;
      moves: Record<string, Record<string, unknown>[]>;
      v: number;
    }
  >();
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-tm-")),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) => {
        const rows = TABLES[q.base];
        return rows && !q.fn ? pick(q.base, rows, q.params.get("$filter") ?? "") : undefined;
      },
    },
  );
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    const path = decodeURIComponent(url.replace(/^https?:\/\/[^/]+\/[^?]*?standard\.odata\//, ""));
    const method = (init?.method ?? "GET").toUpperCase();
    const json = (status: number, body: unknown) =>
      new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
    const set = /^[^(?/]+/.exec(path)?.[0] ?? "";
    const ref = /\(guid'([0-9a-f-]{36})'\)/i.exec(path)?.[1]?.toLowerCase();
    if (set in MONEY) {
      if (method === "POST" && !ref) {
        const res = await inner(input, init);
        if (!res.ok) return res;
        const body = JSON.parse(String(init?.body)) as Record<string, unknown>;
        posts.push({ set, body });
        docs.set(String(body["Ref_Key"]).toLowerCase(), {
          set: set as MoneySet,
          body,
          posted: false,
          moves: {},
          v: 1,
        });
        return json(201, { ...body, Number: "0000-000001", DataVersion: "v1" });
      }
      const d = ref ? docs.get(ref) : undefined;
      if (d && /\/(Post|Unpost)\?/.test(path)) {
        d.posted = path.includes("/Post?");
        if (d.posted) {
          const m = movesOf(d.set, d.body);
          d.moves = Object.fromEntries(
            Object.entries(m).map(([r, rows]) => [r, o.spoil ? o.spoil(r, rows) : rows]),
          );
        } else if (!o.keepOnUnpost) d.moves = {};
        d.v++;
        return json(200, {});
      }
      if (d && method === "GET")
        return json(200, { ...d.body, Ref_Key: ref, Posted: d.posted, DataVersion: `v${d.v}` });
      if (d && method === "PATCH") {
        const patch = JSON.parse(String(init?.body)) as Record<string, unknown>;
        patches.push(patch);
        Object.assign(d.body, patch);
        return json(200, { ...d.body, Ref_Key: ref });
      }
    }
    const reg = /^AccumulationRegister_([^(]+)\(Recorder='([0-9a-f-]{36})'/i.exec(path);
    if (reg) {
      const rows = docs.get(reg[2]!.toLowerCase())?.moves[reg[1]!] ?? [];
      if (!rows.length)
        return json(404, {
          "odata.error": { code: "9", message: { lang: "ru", value: "Экземпляр сущности не найден" } },
        });
      return json(200, { RecordSet: rows });
    }
    return inner(input, init);
  });
  const call = async (name: string, args: Record<string, unknown>) => {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, sc: r.structuredContent ?? {}, text: r.content.map((c) => c.text).join("") };
  };
  const create = async (name: string, args: Record<string, unknown>) => {
    const p = await call(name, args);
    expect(p.isError, p.text.slice(0, 400)).toBeFalsy();
    const d = await call(name, { ...args, confirm: true, operationId: p.sc["operationId"] });
    expect(d.isError, d.text.slice(0, 400)).toBeFalsy();
    return { preview: p, ref: String(posts.at(-1)!.body["Ref_Key"]) };
  };
  const post = async (entitySet: string, ref: string, p = true) => {
    await call("write.document.post_document", { entitySet, ref, post: p });
    return call("write.document.post_document", { entitySet, ref, post: p, confirm: true });
  };
  return { call, create, post, posts, patches, docs, fake };
}

const IN = "write.money.create_cash_receipt";
const OUT = "write.money.create_cash_payment";
const BIN = "write.money.create_payment";
const BOUT = "write.money.create_bank_writeoff";
const REQ = "write.money.create_payment_request";
const base = { organization: "ТЕСТ", amount: 10, cashflowItem: "Оплата" };
const payloadOf = (r: { sc: Record<string, unknown> }) => r.sc["payload"] as Record<string, unknown>;
const rowOf = (r: { sc: Record<string, unknown> }) =>
  (payloadOf(r)["РасшифровкаПлатежа"] as Record<string, unknown>[] | undefined)?.[0];

describe("денежные документы УТ: состав по хозяйственной операции", () => {
  it.each([
    [
      IN,
      {
        operationKind: "ПоступлениеОплатыОтКлиента",
        cash: "Касса ТЕСТ",
        counterpartyRef: CP,
        contractRef: CONTRACT,
      },
    ],
    [OUT, { operationKind: "ОплатаПоставщику", cash: CASH, counterpartyRef: CP }],
    [BIN, { operationKind: "ПоступлениеОплатыОтКлиента", orgBankAccount: "Основной", counterpartyRef: CP }],
    [BOUT, { operationKind: "ОплатаПоставщику", orgBankAccount: BANK, counterpartyRef: CP }],
    [BOUT, { operationKind: "ВозвратОплатыКлиенту", orgBankAccount: BANK, counterpartyRef: CP }],
  ] as const)(
    "%s %o: партнёр из контрагента, расшифровка с партнёром и взаиморасчётами в рублях",
    async (tool, args) => {
      const s = setup();
      const r = await s.call(tool, { ...base, ...args });
      expect(r.isError, r.text).toBeFalsy();
      expect(payloadOf(r)).toMatchObject({
        Организация_Key: ORG,
        Контрагент_Key: CP,
        Партнер_Key: PARTNER,
        СуммаДокумента: 10,
        Валюта_Key: RUB,
      });
      expect(rowOf(r)).toMatchObject({
        Партнер_Key: PARTNER,
        СтатьяДвиженияДенежныхСредств_Key: CF,
        Сумма: 10,
        ВалютаВзаиморасчетов_Key: RUB,
        СуммаВзаиморасчетов: 10,
        Организация_Key: ORG,
      });
      expect(s.fake.violations).toEqual([]);
    },
  );

  it("ПКО от клиента: договор — составное поле с типом, «Принято от» — контрагент", async () => {
    const s = setup();
    const r = await s.call(IN, {
      ...base,
      operationKind: "ПоступлениеОплатыОтКлиента",
      cash: CASH,
      counterpartyRef: CP,
      contractRef: CONTRACT,
    });
    expect(payloadOf(r)).toMatchObject({
      Касса_Key: CASH,
      Договор: CONTRACT,
      Договор_Type: "StandardODATA.Catalog_ДоговорыКонтрагентов",
      ПринятоОт: "ZZ-TEST",
    });
  });

  it("безналичные — «проведено банком» датой документа; оплата картой — статья в шапке, без расшифровки", async () => {
    const s = setup();
    const r = await s.call(BIN, {
      ...base,
      operationKind: "ПоступлениеОплатыПоПлатежнойКарте",
      orgBankAccount: BANK,
      counterpartyRef: CP,
      date: "2026-11-20",
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(payloadOf(r)).toMatchObject({
      ПроведеноБанком: true,
      ДатаПроведенияБанком: "2026-11-20T00:00:00",
      СтатьяДвиженияДенежныхСредств_Key: CF,
      Контрагент_Key: CP,
    });
    expect(payloadOf(r)).not.toHaveProperty("РасшифровкаПлатежа");
  });

  it("прочее: статья доходов из плана статей доходов, расходов — из статей активов и пассивов, с типом", async () => {
    const s = setup();
    const inc = await s.call(IN, {
      ...base,
      operationKind: "ПрочееПоступлениеДенежныхСредств",
      cash: CASH,
      incomeItem: "Прочие доходы",
    });
    expect(rowOf(inc)).toMatchObject({
      СтатьяДоходов: INCOME,
      СтатьяДоходов_Type: "StandardODATA.ChartOfCharacteristicTypes_СтатьиДоходов",
    });
    const exp = await s.call(BOUT, {
      ...base,
      operationKind: "ПрочаяВыдачаДенежныхСредств",
      orgBankAccount: BANK,
      expenseItem: "Займ",
      department: undefined,
    });
    expect(rowOf(exp)).toMatchObject({
      СтатьяРасходов: ASSET,
      СтатьяРасходов_Type: "StandardODATA.ChartOfCharacteristicTypes_СтатьиАктивовПассивов",
    });
  });

  it("перемещения денег: из другой кассы, инкассация, с другого счёта, из кассы на счёт — без расшифровки", async () => {
    const s = setup();
    const a = await s.call(IN, {
      organization: "ТЕСТ",
      amount: 10,
      operationKind: "ПоступлениеДенежныхСредствИзДругойКассы",
      cash: CASH2,
      fromCash: CASH,
    });
    expect(payloadOf(a)).toMatchObject({
      Касса_Key: CASH2,
      КассаОтправитель: CASH,
      КассаОтправитель_Type: "StandardODATA.Catalog_Кассы",
    });
    const b = await s.call(OUT, {
      organization: "ТЕСТ",
      amount: 10,
      operationKind: "ИнкассацияДенежныхСредствВБанк",
      cash: CASH,
      toBankAccount: "Основной",
    });
    expect(payloadOf(b)).toMatchObject({ Касса_Key: CASH, БанковскийСчет_Key: BANK });
    const c = await s.call(BIN, {
      organization: "ТЕСТ",
      amount: 10,
      operationKind: "ПоступлениеДенежныхСредствСДругогоСчета",
      orgBankAccount: BANK,
      fromBankAccount: BANK2,
    });
    expect(payloadOf(c)).toMatchObject({ БанковскийСчет_Key: BANK, БанковскийСчетОтправитель_Key: BANK2 });
    const d = await s.call(BIN, {
      organization: "ТЕСТ",
      amount: 10,
      operationKind: "ПоступлениеДенежныхСредствИзКассыНаРасчетныйСчет",
      orgBankAccount: BANK,
      fromCash: CASH,
    });
    expect(payloadOf(d)).toMatchObject({ КассаОтправитель_Key: CASH });
    for (const r of [a, b, c, d]) {
      expect(r.isError, r.text).toBeFalsy();
      expect(payloadOf(r)).not.toHaveProperty("РасшифровкаПлатежа");
    }
    expect(s.fake.violations).toEqual([]);
  });

  it("подотчётнику — физлицо, «Выдать» — его имя; в бюджет — КБК, ОКТМО, статус составителя", async () => {
    const s = setup();
    const p = await s.call(OUT, {
      ...base,
      operationKind: "ВыдачаДенежныхСредствПодотчетнику",
      cash: CASH,
      person: "Подотчётник",
    });
    expect(payloadOf(p)).toMatchObject({ ПодотчетноеЛицо_Key: PERSON, Выдать: "Подотчётник" });
    const b = await s.call(BOUT, {
      ...base,
      operationKind: "ПеречислениеВБюджет",
      orgBankAccount: BANK,
      counterpartyRef: CP,
      budget: { kbk: "18201061201010000510", oktmo: "70000000", payerStatus: "01" },
    });
    expect(payloadOf(b)).toMatchObject({
      ПеречислениеВБюджет: true,
      ВидПеречисленияВБюджет: "НалоговыйПлатеж",
      КодБК: "18201061201010000510",
      КодОКАТО: "70000000",
      СтатусСоставителя: "01",
    });
    expect(s.fake.violations).toEqual([]);
  });

  it("заявка: «Не согласована» по умолчанию, форма оплаты — одна, касса или счёт по ней", async () => {
    const s = setup();
    const cash = await s.call(REQ, {
      ...base,
      operationKind: "ПрочаяВыдачаДенежныхСредств",
      payForm: "cash",
      cash: CASH,
      expenseItem: "Прочие расходы",
    });
    expect(payloadOf(cash)).toMatchObject({
      Статус: "НеСогласована",
      ФормаОплатыНаличная: true,
      ФормаОплатыБезналичная: false,
      ФормаОплатыЗаявки: "Наличная",
      Касса_Key: CASH,
    });
    expect(JSON.stringify(cash.sc["notes"])).toContain("Не согласована");
    const bank = await s.call(REQ, {
      ...base,
      operationKind: "ОплатаПоставщику",
      payForm: "bank",
      orgBankAccount: BANK,
      counterpartyRef: CP,
      status: "КОплате",
      paymentDate: "2026-11-25",
    });
    expect(payloadOf(bank)).toMatchObject({
      Статус: "КОплате",
      ФормаОплатыБезналичная: true,
      ФормаОплатыНаличная: false,
      БанковскийСчет_Key: BANK,
      ДатаПлатежа: "2026-11-25T00:00:00",
    });
    expect(s.fake.violations).toEqual([]);
  });
});

describe("денежные документы УТ: отказы без записи", () => {
  it.each([
    [IN, { ...base, operationKind: "ПоступлениеДенежныхСредствИзКассыККМ", cash: CASH }, "розница"],
    [OUT, { ...base, operationKind: "ВыдачаДенежныхСредствВКассуККМ", cash: CASH }, "розница"],
    [
      BOUT,
      { ...base, operationKind: "ВыплатаЗарплатыПоЗарплатномуПроекту", orgBankAccount: BANK },
      "зарплатной",
    ],
    [
      IN,
      {
        organization: "ТЕСТ",
        amount: 10,
        operationKind: "ПоступлениеОплатыОтКлиента",
        cash: CASH,
        counterpartyRef: CP,
      },
      "cashflowItem",
    ],
    [OUT, { ...base, operationKind: "ОплатаПоставщику", cash: CASH }, "counterpartyRef"],
    [IN, { ...base, operationKind: "ПрочееПоступлениеДенежныхСредств", cash: CASH }, "incomeItem"],
    [
      IN,
      {
        organization: "ТЕСТ",
        amount: 10,
        operationKind: "ПоступлениеДенежныхСредствИзДругойКассы",
        cash: CASH,
        fromCash: CASH,
      },
      "совпадает",
    ],
    [
      BOUT,
      {
        ...base,
        operationKind: "ПеречислениеВБюджет",
        orgBankAccount: BANK,
        counterpartyRef: CP,
        budget: { kbk: "1", oktmo: "", payerStatus: "01" },
      },
      "oktmo",
    ],
    [REQ, { ...base, operationKind: "ОплатаПоставщику", counterpartyRef: CP }, "payForm"],
    [IN, { ...base, operationKind: "ПоступлениеОплатыОтКлиента", counterpartyRef: CP }, "несколько"],
    [OUT, { ...base, cash: CASH }, "operationKind"],
  ] as const)("%s %o → «%s»", async (tool, args, text) => {
    const s = setup();
    const r = await s.call(tool, args as Record<string, unknown>);
    expect(r.isError, r.text).toBe(true);
    expect(r.text).toContain(text);
    expect(s.posts).toEqual([]);
  });

  it("платёжное поручение в УТ — отказ с указанием заявки и списания безналичных", async () => {
    const s = setup();
    const r = await s.call("write.money.create_payout_order", {
      counterpartyRef: CP,
      contractRef: CONTRACT,
      amount: 10,
      number: "1",
      purposeText: "ZZ",
      recipientText: "ZZ",
    });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("create_payment_request");
  });
});

describe("денежные документы УТ: проведение со сверкой движений", () => {
  const cases = [
    [
      IN,
      "Document_ПриходныйКассовыйОрдер",
      { ...base, operationKind: "ПоступлениеОплатыОтКлиента", cash: CASH, counterpartyRef: CP },
    ],
    [
      OUT,
      "Document_РасходныйКассовыйОрдер",
      { ...base, operationKind: "ПрочаяВыдачаДенежныхСредств", cash: CASH, expenseItem: "Прочие расходы" },
    ],
    [
      BIN,
      "Document_ПоступлениеБезналичныхДенежныхСредств",
      {
        ...base,
        operationKind: "ПрочееПоступлениеДенежныхСредств",
        orgBankAccount: BANK,
        incomeItem: "Прочие доходы",
      },
    ],
    [
      BOUT,
      "Document_СписаниеБезналичныхДенежныхСредств",
      { ...base, operationKind: "ОплатаПоставщику", orgBankAccount: BANK, counterpartyRef: CP },
    ],
    [
      BOUT,
      "Document_СписаниеБезналичныхДенежныхСредств",
      { ...base, operationKind: "ВозвратОплатыКлиенту", orgBankAccount: BANK, counterpartyRef: CP },
    ],
    [
      REQ,
      "Document_ЗаявкаНаРасходованиеДенежныхСредств",
      {
        ...base,
        operationKind: "ОплатаПоставщику",
        payForm: "bank",
        orgBankAccount: BANK,
        counterpartyRef: CP,
      },
    ],
  ] as const;

  it.each(cases)(
    "%s / %s: движения совпали — verified; отмена — пусто, verified",
    async (tool, set, args) => {
      const s = setup();
      const { ref } = await s.create(tool, args);
      const p = await s.post(set, ref);
      expect(p.isError, p.text).toBeFalsy();
      expect(p.sc).toMatchObject({ done: true, verified: true, movementsVerified: true });
      const u = await s.post(set, ref, false);
      expect(u.sc).toMatchObject({ done: true, verified: true, movementsVerified: true });
      expect(s.fake.requests.some((q) => q.resource.startsWith("AccountingRegister_"))).toBe(false);
    },
  );

  it("другая сумма в регистре денег, движение по чужой кассе, нет расчётов с партнёром, остаток после отмены — verified=false", async () => {
    const run = async (
      o: Parameters<typeof setup>[0],
      args: Record<string, unknown>,
      set: string,
      unpost = false,
    ) => {
      const s = setup(o);
      const { ref } = await s.create(OUT, args);
      const p = await s.post(set, ref);
      return unpost ? s.post(set, ref, false) : p;
    };
    const SET = "Document_РасходныйКассовыйОрдер";
    const pay = { ...base, operationKind: "ОплатаПоставщику", cash: CASH, counterpartyRef: CP };
    const amount = await run(
      {
        spoil: (r, rows) => (r === "ДенежныеСредстваНаличные" ? rows.map((x) => ({ ...x, Сумма: 9 })) : rows),
      },
      pay,
      SET,
    );
    expect(amount.sc).toMatchObject({ verified: false, movementsVerified: false });
    expect(JSON.stringify(amount.sc["movements"])).toContain("расход 9, в документе 10");
    const other = await run(
      {
        spoil: (r, rows) =>
          r === "ДенежныеСредстваНаличные" ? rows.map((x) => ({ ...x, Касса_Key: CASH2 })) : rows,
      },
      pay,
      SET,
    );
    expect(other.sc).toMatchObject({ verified: false });
    const noParty = await run({ spoil: (r, rows) => (r === "РасчетыСПоставщиками" ? [] : rows) }, pay, SET);
    expect(JSON.stringify(noParty.sc["movements"])).toContain("РасчетыСПоставщиками");
    const kept = await run({ keepOnUnpost: true }, pay, SET, true);
    expect(kept.sc).toMatchObject({ verified: false });
    expect(String(kept.sc["note"])).toContain("остались");
  });
});

describe("денежные документы УТ: копия и поля", () => {
  const SET = "Document_СписаниеБезналичныхДенежныхСредств";
  const pay = { ...base, operationKind: "ОплатаПоставщику", orgBankAccount: BANK, counterpartyRef: CP };
  const twice = async (s: ReturnType<typeof setup>, name: string, args: Record<string, unknown>) => {
    const p = await s.call(name, args);
    if (p.isError) return p;
    return s.call(name, {
      ...args,
      confirm: true,
      ...(p.sc["operationId"] ? { operationId: p.sc["operationId"] } : {}),
    });
  };

  it("сумма документа меняется вместе с единственной строкой расшифровки; поле вне списка и проведённый — отказ", async () => {
    const s = setup();
    const { ref } = await s.create(BOUT, pay);
    s.docs.get(ref)!.body["РасшифровкаПлатежа"] = [
      { LineNumber: 1, Сумма: 10, СуммаВзаиморасчетов: 10, ОбъектРасчетов_Key: g("9", 1) },
    ];
    const r = await twice(s, "write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { СуммаДокумента: 25, НазначениеПлатежа: "ZZ" },
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(s.patches.at(-1)).toMatchObject({
      СуммаДокумента: 25,
      НазначениеПлатежа: "ZZ",
      РасшифровкаПлатежа: [{ Сумма: 25, СуммаВзаиморасчетов: 25, ОбъектРасчетов_Key: g("9", 1) }],
    });
    expect(r.sc).toMatchObject({ updated: true, verified: true });
    const bad = await s.call("write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { Контрагент_Key: CP },
    });
    expect(bad.text).toContain("Контрагент_Key");
    await s.post(SET, ref);
    const posted = await s.call("write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { Комментарий: "x" },
    });
    expect(posted.text).toContain("проведён");
    expect(s.patches).toHaveLength(1);
  });

  it("сумма при нескольких строках расшифровки — отказ: разнести нельзя", async () => {
    const s = setup();
    const { ref } = await s.create(BOUT, pay);
    s.docs.get(ref)!.body["РасшифровкаПлатежа"] = [{ Сумма: 5 }, { Сумма: 5 }];
    const r = await s.call("write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { СуммаДокумента: 12 },
    });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("несколько строк");
  });

  it("копия: без объекта расчётов и связей, «проведено банком» новой датой, новая сумма; строки и чужие поля — отказ", async () => {
    const s = setup();
    const { ref } = await s.create(BOUT, pay);
    Object.assign(s.docs.get(ref)!.body, {
      ДокументОснование: g("9", 2),
      НомерВходящегоДокумента: "77",
      РасшифровкаПлатежа: [
        {
          LineNumber: 1,
          Сумма: 10,
          СуммаВзаиморасчетов: 10,
          ОбъектРасчетов_Key: g("9", 1),
          Партнер_Key: PARTNER,
        },
      ],
    });
    const c = await s.create("write.document.copy_document", {
      entitySet: SET,
      ref,
      date: "2026-11-25",
      fields: { СуммаДокумента: 30 },
    });
    const body = s.posts.at(-1)!.body;
    expect(body).toMatchObject({
      Posted: false,
      ПроведеноБанком: true,
      ДатаПроведенияБанком: "2026-11-25T00:00:00",
      СуммаДокумента: 30,
    });
    expect(body).not.toHaveProperty("ДокументОснование");
    expect(body).not.toHaveProperty("НомерВходящегоДокумента");
    expect(body["РасшифровкаПлатежа"]).toEqual([
      expect.objectContaining({ Сумма: 30, СуммаВзаиморасчетов: 30, Партнер_Key: PARTNER }),
    ]);
    expect(JSON.stringify(body["РасшифровкаПлатежа"])).not.toContain("ОбъектРасчетов");
    expect(JSON.stringify(c.preview.sc["notes"])).toContain("без объекта расчётов");
    const lines = await s.call("write.document.copy_document", {
      entitySet: SET,
      ref,
      lines: [{ lineNumber: 1, fields: { Сумма: 1 } }],
    });
    expect(lines.text).toContain("строк товаров нет");
    const bad = await s.call("write.document.copy_document", {
      entitySet: SET,
      ref,
      fields: { Контрагент_Key: CP },
    });
    expect(bad.text).toContain("Контрагент_Key");
  });

  it("копия заявки — «Не согласована» и даты платежа по новой дате", async () => {
    const s = setup();
    const { ref } = await s.create(REQ, {
      ...base,
      operationKind: "ОплатаПоставщику",
      payForm: "bank",
      orgBankAccount: BANK,
      counterpartyRef: CP,
      status: "КОплате",
    });
    await s.create("write.document.copy_document", {
      entitySet: "Document_ЗаявкаНаРасходованиеДенежныхСредств",
      ref,
      date: "2026-11-26",
    });
    expect(s.posts.at(-1)!.body).toMatchObject({
      Статус: "НеСогласована",
      ДатаПлатежа: "2026-11-26T00:00:00",
    });
  });
});

describe("денежные инструменты в Бухгалтерии — как раньше", () => {
  it("параметры УТ — отказ; оплата от покупателя без контрагента или договора — отказ; заявки нет", async () => {
    const s = setup({ schema: "bp30" });
    const a = await s.call(IN, { operationKind: "ОплатаПокупателя", amount: 10, cashflowItem: "Оплата" });
    expect(a.text).toContain("только для Управления торговлей");
    const b = await s.call(BIN, { amount: 10, counterpartyRef: CP });
    expect(b.text).toContain("counterpartyRef и contractRef");
    const c = await s.call(REQ, { operationKind: "ОплатаПоставщику", amount: 10, payForm: "bank" });
    expect(c.text).toContain("create_payout_order");
    for (const r of [a, b, c]) expect(r.isError).toBe(true);
    expect(s.posts).toEqual([]);
  });
});
