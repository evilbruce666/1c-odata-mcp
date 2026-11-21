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

function setup(o: { kinds500?: boolean } = {}) {
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
    [{ operationKind: "ОплатаПокупателя", amount: 1, corrAccount: "91.02" }, "только для прочих операций"],
    [{ operationKind: "ОплатаПокупателя", amount: 1, otherItemRef: ITEM91 }, "только для прочих операций"],
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
    const withHeader = { ДоговорКонтрагента_Key: CTR, СчетУчетаРасчетовСКонтрагентом_Key: A7609 };
    expect(emptySettlementWarning(created, withHeader)).toBeUndefined();
    // Обычная оплата по договору без счёта — предупреждение, как раньше.
    expect(emptySettlementWarning(created, { ДоговорКонтрагента_Key: CTR })).toContain("пуст счёт расчётов");
  });
});
