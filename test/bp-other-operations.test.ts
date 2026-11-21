import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { createServer } from "../src/mcp/server.js";
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

function setup() {
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-other-")),
    },
    {
      data: (q) =>
        q.fn
          ? undefined
          : q.base === "ChartOfAccounts_Хозрасчетный"
            ? ACCOUNTS
            : q.base === "ChartOfCharacteristicTypes_ВидыСубконтоХозрасчетные"
              ? KINDS
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
    [{ operationKind: "ОплатаПокупателя", amount: 1, corrAccount: "91.02" }, "только для прочих операций"],
  ])("отказ до записи: %j", async (args, text) => {
    const s = setup();
    const r = await s.call("write.money.create_cash_receipt", { ...args, confirm: false });
    expect(r.err).toBe(true);
    expect(r.text).toContain(text);
    expect(s.fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });
});
