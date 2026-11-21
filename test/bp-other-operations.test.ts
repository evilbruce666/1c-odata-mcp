import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { emptySettlementWarning } from "../src/tools/write.js";
import { fake1C } from "./helpers/fake-1c.js";

/**
 * Бухгалтерия: прочий приход и расход (ПКО, РКО) и прочее списание с расчётного счёта. Найдено вживую 21.11.2026:
 * без корреспондирующего счёта 1С такие документы не проводит («Не заполнен счёт кредита/дебета»), а у списания 1С
 * при записи добавляла строку расшифровки с нулевой суммой — проведение останавливала проверка суммы. На рабочих
 * базах счёт у таких документов разный (76.03, 91.02, 66.01, 55.03…), поэтому он — обязательный параметр.
 */
const g = (c: string, n: number) => `${c.repeat(8)}-0000-0000-0000-${String(n).padStart(12, "0")}`;
const A9102 = g("a", 1);
const A91 = g("a", 2);
const A7609 = g("a", 3);
const K_OTHER = g("b", 1);
const K_CP = g("b", 2);
const K_CTR = g("b", 3);
const K_DOC = g("b", 4);
const CP = g("c", 1);
const CTR = g("c", 2);
const ITEM91 = g("c", 3);

const ACCOUNTS = [
  {
    Ref_Key: A91,
    Code: "91",
    Description: "Прочие",
    ЗапретитьИспользоватьВПроводках: true,
    ExtDimensionTypes: [],
  },
  {
    Ref_Key: A9102,
    Code: "91.02",
    Description: "Прочие расходы",
    ЗапретитьИспользоватьВПроводках: false,
    ExtDimensionTypes: [{ LineNumber: "1", ExtDimensionType_Key: K_OTHER }],
  },
  {
    Ref_Key: A7609,
    Code: "76.09",
    Description: "Прочие расчёты",
    ЗапретитьИспользоватьВПроводках: false,
    ExtDimensionTypes: [
      { LineNumber: "2", ExtDimensionType_Key: K_CTR },
      { LineNumber: "1", ExtDimensionType_Key: K_CP },
      { LineNumber: "3", ExtDimensionType_Key: K_DOC },
    ],
  },
];
ACCOUNTS.push({
  Ref_Key: g("a", 6),
  Code: "91.01",
  Description: "Прочие доходы",
  ЗапретитьИспользоватьВПроводках: false,
  ExtDimensionTypes: [{ LineNumber: "1", ExtDimensionType_Key: K_OTHER }],
});
ACCOUNTS.push({
  Ref_Key: g("a", 4),
  Code: "68.90",
  Description: "Прочие налоги",
  ЗапретитьИспользоватьВПроводках: false,
  ExtDimensionTypes: [],
});
const KINDS = [
  { Ref_Key: K_OTHER, PredefinedDataName: "ПрочиеДоходыИРасходы", Description: "Прочие доходы и расходы" },
  { Ref_Key: K_CP, PredefinedDataName: "Контрагенты", Description: "Контрагенты" },
  { Ref_Key: K_CTR, PredefinedDataName: "Договоры", Description: "Договоры" },
  {
    Ref_Key: K_DOC,
    PredefinedDataName: "ДокументыРасчетовСКонтрагентами",
    Description: "Документы расчётов",
  },
];

function setup(o: { kinds500?: boolean; plainChart?: boolean; env?: Record<string, string> } = {}) {
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-other-")),
      ...o.env,
    },
    {
      data: (q) =>
        q.fn || o.plainChart
          ? undefined
          : q.base === "ChartOfAccounts_Хозрасчетный"
            ? ACCOUNTS
            : q.base === "ChartOfCharacteristicTypes_ВидыСубконтоХозрасчетные"
              ? o.kinds500
                ? []
                : KINDS
              : undefined,
    },
  );
  const tools = (
    createServer(fake.ctx) as unknown as {
      _registeredTools: Record<
        string,
        {
          inputSchema: { parse: (v: unknown) => Record<string, unknown> };
          handler: (
            a: unknown,
            e: unknown,
          ) => Promise<{
            isError?: boolean;
            structuredContent?: Record<string, unknown>;
            content?: { text: string }[];
          }>;
        }
      >;
    }
  )._registeredTools;
  const call = async (name: string, args: Record<string, unknown>) => {
    const t = tools[name]!;
    const r = await t.handler(t.inputSchema.parse(args), {});
    return {
      err: !!r.isError,
      sc: r.structuredContent ?? {},
      text: r.content?.map((c) => c.text).join("") ?? "",
    };
  };
  return { call, fake };
}
const payloadOf = (r: { sc: Record<string, unknown> }) => r.sc["payload"] as Record<string, unknown>;

describe("Бухгалтерия: прочие операции денежных документов — корреспондирующий счёт и субконто", () => {
  it("РКО «Прочий расход» на 76.09: счёт и субконто по видам счёта, контрагента в шапке нет, третье — заметка", async () => {
    const s = setup();
    const r = await s.call("write.money.create_cash_payment", {
      operationKind: "ПрочийРасход",
      amount: 10,
      corrAccount: "76.09",
      counterpartyRef: CP,
      contractRef: CTR,
    });
    expect(r.err, r.text).toBe(false);
    const p = payloadOf(r);
    expect(p).toMatchObject({
      ВидОперации: "ПрочийРасход",
      СчетУчетаРасчетовСКонтрагентом_Key: A7609,
      СубконтоДт1: CP,
      СубконтоДт1_Type: "StandardODATA.Catalog_Контрагенты",
      СубконтоДт2: CTR,
      СубконтоДт2_Type: "StandardODATA.Catalog_ДоговорыКонтрагентов",
    });
    expect(p).not.toHaveProperty("Контрагент");
    expect(p).not.toHaveProperty("ДоговорКонтрагента_Key");
    expect(p).not.toHaveProperty("СубконтоДт3");
    expect(JSON.stringify(r.sc["notes"])).toContain("Документы расчётов");
    expect(s.fake.violations).toEqual([]);
  });

  it("ПКО «Прочий приход» — субконто кредита; списание «Прочее списание» на 91.02 — строка расшифровки с суммой", async () => {
    const s = setup();
    const pko = await s.call("write.money.create_cash_receipt", {
      operationKind: "ПрочийПриход",
      amount: 5,
      corrAccount: "91.02",
      otherItemRef: ITEM91,
    });
    expect(payloadOf(pko)).toMatchObject({
      СчетУчетаРасчетовСКонтрагентом_Key: A9102,
      СубконтоКт1: ITEM91,
      СубконтоКт1_Type: "StandardODATA.Catalog_ПрочиеДоходыИРасходы",
    });
    const w = await s.call("write.money.create_bank_writeoff", {
      operationKind: "ПрочееСписание",
      amount: 7,
      corrAccount: "91.02",
      otherItemRef: ITEM91,
    });
    expect(w.err, w.text).toBe(false);
    expect(payloadOf(w)).toMatchObject({
      СчетУчетаРасчетовСКонтрагентом_Key: A9102,
      СубконтоДт1: ITEM91,
      РасшифровкаПлатежа: [{ LineNumber: 1, СуммаПлатежа: 7 }],
    });
    expect(s.fake.violations).toEqual([]);
  });

  it.each([
    [{ operationKind: "ПрочийПриход", amount: 1 }, "укажите корреспондирующий счёт corrAccount"],
    [{ operationKind: "ПрочийПриход", amount: 1, corrAccount: "91.02" }, "статья прочих доходов и расходов"],
    [{ operationKind: "ПрочийПриход", amount: 1, corrAccount: "91", otherItemRef: ITEM91 }, "группа"],
    [{ operationKind: "ПрочийПриход", amount: 1, corrAccount: "99.99" }, "не найден"],
    [{ operationKind: "ОплатаПокупателя", amount: 1, corrAccount: "91.02" }, "только для операций"],
    [{ operationKind: "ОплатаПокупателя", amount: 1, otherItemRef: ITEM91 }, "только для операций"],
  ])("отказ до записи: %j", async (args, text) => {
    const s = setup();
    const r = await s.call("write.money.create_cash_receipt", { ...args, confirm: false });
    expect(r.err).toBe(true);
    expect(r.text).toContain(text);
    expect(s.fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });

  it("ревью выпуска 5: вид субконто не прочитался — отказ, а не тихая потеря; значение без субконто — заметка", async () => {
    const bad = setup({ kinds500: true });
    const r = await bad.call("write.money.create_cash_payment", {
      operationKind: "ПрочийРасход",
      amount: 1,
      corrAccount: "76.09",
      counterpartyRef: CP,
      contractRef: CTR,
    });
    expect(r.err).toBe(true);
    expect(r.text).toContain("Не удалось прочитать вид субконто");
    expect(bad.fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
    // 91.02 без субконто «контрагент»: в РКО прочего расхода контрагента нет и в шапке — заметка.
    const s = setup();
    const n = await s.call("write.money.create_cash_payment", {
      operationKind: "ПрочийРасход",
      amount: 1,
      corrAccount: "91.02",
      otherItemRef: ITEM91,
      counterpartyRef: CP,
    });
    expect(n.err, n.text).toBe(false);
    expect(JSON.stringify(n.sc["notes"])).toContain("нет субконто «контрагент");
    // У списания контрагент остаётся в шапке — заметки о потере нет.
    const w = await s.call("write.money.create_bank_writeoff", {
      operationKind: "ПрочееСписание",
      amount: 1,
      corrAccount: "91.02",
      otherItemRef: ITEM91,
      counterpartyRef: CP,
    });
    expect(w.err, w.text).toBe(false);
    expect(JSON.stringify(w.sc["notes"] ?? [])).not.toContain("нет субконто");
  });

  it("прочее списание по договору: после записи нет ложного «пуст счёт расчётов» — счёт в шапке", async () => {
    const s = setup();
    const args = {
      operationKind: "ПрочееСписание",
      amount: 1,
      corrAccount: "76.09",
      counterpartyRef: CP,
      contractRef: CTR,
    };
    const p = await s.call("write.money.create_bank_writeoff", args);
    const c = await s.call("write.money.create_bank_writeoff", {
      ...args,
      confirm: true,
      operationId: p.sc["operationId"],
    });
    expect(c.err, c.text).toBe(false);
    expect(JSON.stringify(c.sc)).not.toContain("пуст счёт расчётов");
    // Как ответила живая БП 3.0: строка расшифровки с пустым счётом расчётов, счёт — в шапке.
    const created = {
      СчетУчетаРасчетовСКонтрагентом_Key: A7609,
      РасшифровкаПлатежа: [
        { СчетУчетаРасчетовСКонтрагентом_Key: "00000000-0000-0000-0000-000000000000", СуммаПлатежа: 1 },
      ],
    };
    const withHeader = {
      ВидОперации: "ПрочееСписание",
      ДоговорКонтрагента_Key: CTR,
      СчетУчетаРасчетовСКонтрагентом_Key: A7609,
    };
    expect(emptySettlementWarning(created, withHeader, true)).toBeUndefined();
    // Тот же счёт шапки без corrAccount (строгие умолчания, депозит по договору) — предупреждение, как в 0.17.0.
    expect(emptySettlementWarning(created, { ...withHeader, ВидОперации: "Депозит" })).toContain(
      "пуст счёт расчётов",
    );
    // Обычная оплата по договору без счёта — предупреждение, как раньше.
    expect(emptySettlementWarning(created, { ДоговорКонтрагента_Key: CTR })).toContain("пуст счёт расчётов");
  });

  it("списание: налог — счёт обязателен; зарплата, снятие наличных, перевод — отказ; комиссия и депозит — счёт по желанию; строка с суммой у любого вида без договора", async () => {
    const s = setup();
    const noAcc = await s.call("write.money.create_bank_writeoff", {
      operationKind: "ПеречислениеНалога",
      amount: 1,
    });
    expect(noAcc.err).toBe(true);
    expect(noAcc.text).toContain("укажите корреспондирующий счёт");
    // Зарплата, снятие наличных, перевод на другой счёт: 1С проводила их без проводок или Дт 51 Кт 51 — отказ,
    // и со счётом тоже.
    for (const [op, text] of [
      ["ПеречислениеЗП", "по ведомости"],
      ["СнятиеНаличных", "без проводок"],
      ["ПереводНаДругойСчет", "счёта-получателя"],
    ] as const) {
      const r = await s.call("write.money.create_bank_writeoff", {
        operationKind: op,
        amount: 1,
        corrAccount: "68.90",
      });
      expect(r.err, op).toBe(true);
      expect(r.text).toContain(text);
    }
    const tax = await s.call("write.money.create_bank_writeoff", {
      operationKind: "ПеречислениеНалога",
      amount: 5,
      corrAccount: "68.90",
    });
    expect(tax.err, tax.text).toBe(false);
    expect(payloadOf(tax)).toMatchObject({
      СчетУчетаРасчетовСКонтрагентом_Key: g("a", 4),
      РасшифровкаПлатежа: [{ LineNumber: 1, СуммаПлатежа: 5 }],
    });
    // Комиссия банка без счёта — как раньше (1С ставит 91.02 сама), но строка уже с суммой.
    const fee = await s.call("write.money.create_bank_writeoff", {
      operationKind: "КомиссияБанка",
      amount: 3,
    });
    expect(fee.err, fee.text).toBe(false);
    expect(payloadOf(fee)).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
    expect(payloadOf(fee)).toMatchObject({ РасшифровкаПлатежа: [{ LineNumber: 1, СуммаПлатежа: 3 }] });
    const feeItem = await s.call("write.money.create_bank_writeoff", {
      operationKind: "КомиссияБанка",
      amount: 3,
      corrAccount: "91.02",
      otherItemRef: ITEM91,
    });
    expect(payloadOf(feeItem)).toMatchObject({
      СчетУчетаРасчетовСКонтрагентом_Key: A9102,
      СубконтоДт1: ITEM91,
    });
    const lonely = await s.call("write.money.create_bank_writeoff", {
      operationKind: "КомиссияБанка",
      amount: 3,
      otherItemRef: ITEM91,
    });
    expect(lonely.err).toBe(true);
    expect(lonely.text).toContain("вместе с corrAccount");
    expect(s.fake.violations).toEqual([]);
  });

  it("статья ДДС как субконто счёта не обязательна — она и в шапке (57.01 в документах рабочих баз без субконто)", async () => {
    ACCOUNTS.push({
      Ref_Key: g("a", 5),
      Code: "57.01",
      Description: "Переводы в пути",
      ЗапретитьИспользоватьВПроводках: false,
      ExtDimensionTypes: [{ LineNumber: "1", ExtDimensionType_Key: g("b", 5) }],
    });
    KINDS.push({
      Ref_Key: g("b", 5),
      PredefinedDataName: "СтатьиДвиженияДенежныхСредств",
      Description: "Статьи ДДС",
    });
    const s = setup();
    const r = await s.call("write.money.create_bank_writeoff", {
      operationKind: "Депозит",
      amount: 1,
      corrAccount: "57.01",
    });
    expect(r.err, r.text).toBe(false);
    expect(payloadOf(r)).toMatchObject({ СчетУчетаРасчетовСКонтрагентом_Key: g("a", 5) });
    expect(payloadOf(r)).not.toHaveProperty("СубконтоДт1");
    const cf = g("c", 9);
    const w = await s.call("write.money.create_bank_writeoff", {
      operationKind: "Депозит",
      amount: 1,
      corrAccount: "57.01",
      cashflowItemRef: cf,
    });
    expect(payloadOf(w)).toMatchObject({ СубконтоДт1: cf, СтатьяДвиженияДенежныхСредств_Key: cf });
  });

  it("поступление на счёт «Прочее поступление»: Кт 91.01 со статьёй, строка с суммой, без договора; corrAccount у оплаты — отказ", async () => {
    const s = setup();
    const cf = g("c", 8);
    const r = await s.call("write.money.create_payment", {
      operationKind: "ПрочееПоступление",
      amount: 12,
      corrAccount: "91.01",
      otherItemRef: ITEM91,
      cashflowItemRef: cf,
      comment: "ZZ",
    });
    expect(r.err, r.text).toBe(false);
    const p = payloadOf(r);
    expect(p).toMatchObject({
      ВидОперации: "ПрочееПоступление",
      СчетУчетаРасчетовСКонтрагентом_Key: g("a", 6),
      СубконтоКт1: ITEM91,
      СубконтоКт1_Type: "StandardODATA.Catalog_ПрочиеДоходыИРасходы",
      СтатьяДвиженияДенежныхСредств_Key: cf,
      Комментарий: "ZZ",
      РасшифровкаПлатежа: [{ LineNumber: 1, СуммаПлатежа: 12 }],
    });
    expect(p).not.toHaveProperty("ДоговорКонтрагента_Key");
    expect(s.fake.violations).toEqual([]);
    for (const [args, text] of [
      [{ operationKind: "ПрочееПоступление", amount: 1 }, "укажите корреспондирующий счёт"],
      [
        { amount: 1, counterpartyRef: CP, contractRef: CTR, corrAccount: "91.01" },
        "для operationKind «ПрочееПоступление»",
      ],
      [{ amount: 1, counterpartyRef: CP, contractRef: CTR, comment: "x" }, "comment"],
      [{ amount: 1 }, "укажите counterpartyRef и contractRef"],
    ] as const) {
      const bad = await s.call("write.money.create_payment", args);
      expect(bad.err, JSON.stringify(args)).toBe(true);
      expect(bad.text).toContain(text);
    }
  });

  it("ревью выпуска 5, раунд 3: строка с суммой и молчание о счёте — только белому списку; расчётные виды как раньше", async () => {
    const s = setup({ plainChart: true });
    for (const op of [
      "ПрочиеРасчетыСКонтрагентами",
      "РасчетыПоКредитамИЗаймам",
      "ПеречислениеЗаработнойПлатыРаботнику",
    ]) {
      const r = await s.call("write.money.create_bank_writeoff", {
        operationKind: op,
        amount: 100,
        counterpartyRef: CP,
        contractRef: CTR,
      });
      expect(r.err, `${op}: ${r.text}`).toBe(false);
      expect(payloadOf(r), op).not.toHaveProperty("РасшифровкаПлатежа");
    }
    // Счёт в шапке из настройки договора у расчётного вида — предупреждение о пустом счёте строки остаётся.
    const created = {
      СчетУчетаРасчетовСКонтрагентом_Key: A7609,
      РасшифровкаПлатежа: [
        { СчетУчетаРасчетовСКонтрагентом_Key: "00000000-0000-0000-0000-000000000000", СуммаПлатежа: 1 },
      ],
    };
    const payload = {
      ВидОперации: "ПрочиеРасчетыСКонтрагентами",
      ДоговорКонтрагента_Key: CTR,
      СчетУчетаРасчетовСКонтрагентом_Key: A7609,
    };
    expect(emptySettlementWarning(created, payload)).toContain("пуст счёт расчётов");
    expect(
      emptySettlementWarning(created, { ...payload, ВидОперации: "ПрочееСписание" }, true),
    ).toBeUndefined();
  });

  it("хвосты: оплата поставщику без договора — отказ; запрещённый вид — отказ и в копии, и в update_entity", async () => {
    const s = setup({ plainChart: true });
    for (const op of ["ОплатаПоставщику", "ВозвратПокупателю"]) {
      const r = await s.call("write.money.create_bank_writeoff", {
        operationKind: op,
        amount: 1,
        counterpartyRef: CP,
      });
      expect(r.err, op).toBe(true);
      expect(r.text).toContain("укажите contractRef");
    }
    const SET = "Document_СписаниеСРасчетногоСчета";
    const ref = g("d", 1);
    const copy = await s.call("write.document.copy_document", {
      entitySet: SET,
      ref,
      fields: { ВидОперации: "ПеречислениеЗП" },
    });
    expect(copy.err).toBe(true);
    expect(copy.text).toContain("по ведомости");
    const upd = await s.call("write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { ВидОперации: "СнятиеНаличных" },
    });
    expect(upd.err).toBe(true);
    expect(upd.text).toContain("без проводок");
    expect(s.fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });

  it("хвосты: «ПрочееПоступление» из настройки базы (operationKinds.customerPayment) — путь прочего поступления", async () => {
    const s = setup({
      env: {
        ODATA_ACCOUNTING_CONFIG: JSON.stringify({ operationKinds: { customerPayment: "ПрочееПоступление" } }),
      },
    });
    const r = await s.call("write.money.create_payment", {
      amount: 1,
      counterpartyRef: CP,
      contractRef: CTR,
    });
    expect(r.err).toBe(true);
    expect(r.text).toContain("укажите корреспондирующий счёт");
    const ok = await s.call("write.money.create_payment", {
      amount: 1,
      corrAccount: "91.01",
      otherItemRef: ITEM91,
    });
    expect(ok.err, ok.text).toBe(false);
    expect(payloadOf(ok)).toMatchObject({ ВидОперации: "ПрочееПоступление", СубконтоКт1: ITEM91 });
  });

  it("ревью 1.0: копия «Прочего списания» по договору — без ложного «пуст счёт расчётов»", async () => {
    const E = "00000000-0000-0000-0000-000000000000";
    const SET = "Document_СписаниеСРасчетногоСчета";
    const SRC = g("d", 7);
    const make = (kind: string) =>
      fake1C(
        {
          READ_ONLY: "false",
          ODATA_WRITABLE: "true",
          ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-copywarn-")),
        },
        {
          data: (q) =>
            q.base === SET && !q.fn
              ? [
                  {
                    Ref_Key: SRC,
                    Posted: true,
                    ВидОперации: kind,
                    Контрагент: CP,
                    Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
                    ДоговорКонтрагента_Key: CTR,
                    СчетУчетаРасчетовСКонтрагентом_Key: A7609,
                    СуммаДокумента: 1,
                    РасшифровкаПлатежа: [
                      {
                        LineNumber: "1",
                        СуммаПлатежа: 1,
                        ДоговорКонтрагента_Key: kind === "ПрочееСписание" ? E : CTR,
                        СчетУчетаРасчетовСКонтрагентом_Key: kind === "ПрочееСписание" ? E : A7609,
                        СуммаВзаиморасчетов: 1,
                      },
                    ],
                  },
                ]
              : undefined,
        },
      );
    for (const [kind, warn] of [
      ["ПрочееСписание", false],
      ["ОплатаПоставщику", false],
    ] as const) {
      const fake = make(kind);
      const t = (
        createServer(fake.ctx) as unknown as {
          _registeredTools: Record<
            string,
            {
              inputSchema: { parse: (v: unknown) => Record<string, unknown> };
              handler: (a: unknown, e: unknown) => Promise<{ structuredContent?: Record<string, unknown> }>;
            }
          >;
        }
      )._registeredTools["write.document.copy_document"]!;
      const p = await t.handler(t.inputSchema.parse({ entitySet: SET, ref: SRC }), {});
      const c = await t.handler(
        t.inputSchema.parse({
          entitySet: SET,
          ref: SRC,
          confirm: true,
          operationId: p.structuredContent?.["operationId"],
        }),
        {},
      );
      expect(/пуст счёт расчётов/.test(JSON.stringify(c.structuredContent)), kind).toBe(warn);
    }
  });
});
