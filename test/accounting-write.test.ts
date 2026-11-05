import { describe, expect, it, vi } from "vitest";
import { parseConfig } from "../src/config/env.js";
import { createServer } from "../src/mcp/server.js";
import { Connection } from "../src/context.js";
const id = "11111111-1111-4111-8111-111111111111";
const special = "22222222-2222-4222-8222-222222222222";
const line = { nomenclatureRef: id, quantity: 1, price: 1, vatRate: "БезНДС" };
function fixture(config: object, register = false, settlement?: "both" | "settle" | "advance") {
  const parsed = parseConfig({
    ODATA_BASE_URL: "https://example.invalid/odata/",
    ODATA_USERNAME: "u",
    ODATA_PASSWORD: "p",
    ODATA_ACCOUNTING_CONFIG: JSON.stringify(config),
  });
  const conn = new Connection(parsed.connections[0]!, parsed.behavior);
  const sets = [
    "ChartOfAccounts_Test",
    "Document_РеализацияТоваровУслуг",
    "Document_ПоступлениеНаРасчетныйСчет",
    "Document_СписаниеСРасчетногоСчета",
    "Document_ПриходныйКассовыйОрдер",
    "Document_РасходныйКассовыйОрдер",
    "Document_ПоступлениеТоваровУслуг",
    "Catalog_Организации",
    "Catalog_ДоговорыКонтрагентов",
    "Catalog_Контрагенты",
    "Catalog_Валюты",
    ...(settlement ? ["InformationRegister_СчетаУчетаРасчетовСКонтрагентами"] : []),
    ...(register ? ["InformationRegister_СчетаУчетаНоменклатуры"] : []),
  ];
  vi.spyOn(conn, "available").mockResolvedValue(new Set(sets));
  vi.spyOn(conn, "getMetadata").mockResolvedValue({
    entities: new Map(
      sets.flatMap((entitySet) => [
        [
          entitySet,
          {
            properties:
              entitySet === "Catalog_Контрагенты"
                ? ["Ref_Key", "Parent_Key", "IsFolder", "DeletionMark"].map((name) => ({
                    name,
                    type: name.endsWith("_Key") ? "Edm.Guid" : "Edm.Boolean",
                  }))
                : [{ name: "СчетУчетаРасчетовСКонтрагентом_Key" }],
          },
        ],
        [
          `${entitySet}_РасшифровкаПлатежа`,
          {
            properties: ["СуммаВзаиморасчетов", "КурсВзаиморасчетов", "КратностьВзаиморасчетов"].map(
              (name) => ({ name }),
            ),
          },
        ],
      ]),
    ),
  } as never);
  vi.spyOn(conn.client, "prepareCreate").mockResolvedValue(undefined as never);
  vi.spyOn(conn.client, "getEntity").mockResolvedValue({
    Ref_Key: id,
    Description: "Synthetic",
    ВидДоговора: "СПокупателем",
    ВалютаВзаиморасчетов_Key: id,
    Code: "643",
    Owner_Key: id,
    Parent_Key: "00000000-0000-0000-0000-000000000000",
    IsFolder: false,
    DeletionMark: false,
  });
  const reads = vi.spyOn(conn.client, "getCollection").mockImplementation(async (path) => {
    if (path.startsWith("Catalog_Организации")) return { value: [{ Ref_Key: id, Description: "Synthetic" }] };
    if (path.startsWith("InformationRegister_СчетаУчетаРасчетовСКонтрагентами"))
      return {
        value: [
          {
            Договор_Key: id,
            Контрагент_Key: id,
            ...(settlement !== "advance" ? { СчетУчетаРасчетовСПокупателем_Key: special } : {}),
            ...(settlement !== "settle" ? { СчетУчетаАвансовПолученных_Key: special } : {}),
          },
        ],
      };
    if (path.startsWith("InformationRegister_"))
      return { value: [{ СчетДоходовОтРеализации_Key: id, СчетРасходовОтРеализации_Key: id }] };
    return {
      value: [
        { Ref_Key: id, Code: "62.1" },
        { Ref_Key: id, Code: "62.2" },
        { Ref_Key: special, Code: "76.09" },
        ...["62.01", "62.02", "60.01", "60.02"].map((Code) => ({ Ref_Key: id, Code })),
      ],
    };
  });
  const tools = (
    createServer({ db: () => conn, databases: () => [] } as never) as unknown as {
      _registeredTools: Record<
        string,
        {
          handler: (
            args: object,
            extra: object,
          ) => Promise<{ isError?: boolean; structuredContent?: { payload?: Record<string, unknown> } }>;
        }
      >;
    }
  )._registeredTools;
  const act = (args: object = {}) =>
    tools["write.sales.create_act"]!.handler(
      {
        organization: id,
        counterpartyRef: id,
        contractRef: id,
        lines: [line],
        sumIncludesVat: true,
        confirm: false,
        ...args,
      },
      {},
    );
  return { act, tools, reads };
}
const strict = { chart: "ChartOfAccounts_Test", strictDefaults: true, defaults: { receivable: ["62.1"] } };
describe("accounting creation flows", () => {
  it("does not require goods/VAT/overridden income or expense defaults for a service", async () => {
    const { act, reads } = fixture({ ...strict, operationKinds: { salesAct: "ПродажаКомиссия" } });
    const r = await act({ incomeAccount: id, expenseAccount: id });
    expect(r.isError, JSON.stringify(r)).not.toBe(true);
    expect(r.structuredContent?.payload).toMatchObject({
      ВидОперации: "ПродажаКомиссия",
      СчетУчетаРасчетовСКонтрагентом_Key: id,
    });
    expect(r.structuredContent?.payload?.Услуги).toMatchObject([
      { СчетДоходов_Key: id, СчетРасходов_Key: id },
    ]);
    expect(r.structuredContent?.payload?.Услуги).not.toMatchObject([{ СчетУчета_Key: id }]);
    expect(
      decodeURIComponent(reads.mock.calls.find(([path]) => path.startsWith("ChartOfAccounts_"))![0]),
    ).not.toContain("41");
  });
  it("uses register accounts before demanding fallback mappings", async () => {
    const { act } = fixture(strict, true);
    expect((await act()).isError).not.toBe(true);
  });
  it("rejects a missing actually needed mapping before preparing a write", async () => {
    const { act } = fixture(strict);
    expect((await act()).isError).toBe(true);
  });
  it("requires purchase goods but no incoming VAT or sales-account mappings for a no-VAT purchase", async () => {
    const { tools } = fixture({ ...strict, defaults: { goods: ["62.1"], payable: ["62.1"] } });
    const r = await tools["write.purchase.create_purchase"]!.handler(
      {
        organization: id,
        counterpartyRef: id,
        contractRef: id,
        lines: [{ ...line, incomeAccount: "unused" }],
        sumIncludesVat: true,
        confirm: false,
      },
      {},
    );
    expect(r.isError, JSON.stringify(r)).not.toBe(true);
    expect(r.structuredContent?.payload?.Товары).toMatchObject([{ СчетУчета_Key: id }]);
  });
  it("explicit operation kind wins over the base setting", async () => {
    const { act } = fixture({ ...strict, operationKinds: { salesAct: "Configured" } });
    expect(
      (await act({ operationKind: "Explicit", incomeAccount: id, expenseAccount: id })).structuredContent
        ?.payload?.ВидОперации,
    ).toBe("Explicit");
  });
  it("uses a published settlement choice before demanding strict settlement defaults in create_act", async () => {
    const { act } = fixture({ ...strict, defaults: {} }, false, "both");
    const r = await act({ incomeAccount: id, expenseAccount: id });
    expect(r.isError, JSON.stringify(r)).not.toBe(true);
    expect(r.structuredContent?.payload?.СчетУчетаРасчетовСКонтрагентом_Key).toBe(special);
  });
  it.each(["settle", "advance"] as const)(
    "requires only the missing %s counterpart default after a partial settlement choice",
    async (part) => {
      const defaults = part === "settle" ? { receivableAdvance: ["62.2"] } : { receivable: ["62.1"] };
      const { tools } = fixture({ ...strict, defaults }, false, part);
      const r = await tools["write.money.create_payment"]!.handler(
        {
          organization: id,
          counterpartyRef: id,
          contractRef: id,
          amount: 1,
          orgBankAccount: id,
          confirm: false,
        },
        {},
      );
      expect(r.isError, JSON.stringify(r)).not.toBe(true);
      expect(r.structuredContent?.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
      expect(r.structuredContent?.payload?.РасшифровкаПлатежа).toMatchObject([
        {
          СчетУчетаРасчетовСКонтрагентом_Key: part === "settle" ? special : id,
          СчетУчетаРасчетовПоАвансам_Key: part === "advance" ? special : id,
        },
      ]);
    },
  );
  it("refuses a partial settlement choice when its missing advance has no configured fallback", async () => {
    const { tools } = fixture({ ...strict, defaults: {} }, false, "settle");
    const r = await tools["write.money.create_payment"]!.handler(
      {
        organization: id,
        counterpartyRef: id,
        contractRef: id,
        amount: 1,
        orgBankAccount: id,
        confirm: false,
      },
      {},
    );
    expect(r.isError).toBe(true);
  });
  it.each([false, true])(
    "custom payment operation with register=%s keeps settlement in rows",
    async (published) => {
      const { tools } = fixture(
        {
          ...strict,
          defaults: published ? {} : { receivable: ["62.1"], receivableAdvance: ["62.2"] },
          operationKinds: { customerPayment: "ConfiguredPayment" },
        },
        false,
        published ? "both" : undefined,
      );
      for (const operationKind of [undefined, "ExplicitPayment"]) {
        const r = await tools["write.money.create_payment"]!.handler(
          {
            organization: id,
            counterpartyRef: id,
            contractRef: id,
            amount: 1,
            orgBankAccount: id,
            operationKind,
            confirm: false,
          },
          {},
        );
        expect(r.isError, JSON.stringify(r)).not.toBe(true);
        expect(r.structuredContent?.payload?.ВидОперации).toBe(operationKind ?? "ConfiguredPayment");
        expect(r.structuredContent?.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
        expect(r.structuredContent?.payload?.РасшифровкаПлатежа).toMatchObject([
          {
            СчетУчетаРасчетовСКонтрагентом_Key: published ? special : id,
            СчетУчетаРасчетовПоАвансам_Key: published ? special : id,
          },
        ]);
      }
    },
  );
});

describe("счета оплат по банку и кассе — только в расшифровке", () => {
  it.each([
    ["create_payment", "ОплатаПокупателя"],
    ["create_bank_writeoff", "ОплатаПоставщику"],
    ["create_cash_receipt", "ОплатаПокупателя"],
    ["create_cash_payment", "ОплатаПоставщику"],
  ])("%s: конструктор не дублирует счёт в шапку без регистра", async (tool, operationKind) => {
    const { tools } = fixture({ chart: "ChartOfAccounts_Test" });
    const r = await tools[`write.money.${tool}`]!.handler(
      {
        organization: id,
        counterpartyRef: id,
        contractRef: id,
        amount: 125,
        orgBankAccount: id,
        operationKind,
        confirm: false,
      },
      {},
    );
    expect(r.isError, JSON.stringify(r)).not.toBe(true);
    expect(r.structuredContent?.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
    expect(r.structuredContent?.payload?.РасшифровкаПлатежа).toMatchObject([
      {
        СчетУчетаРасчетовСКонтрагентом_Key: id,
        СчетУчетаРасчетовПоАвансам_Key: id,
        СуммаПлатежа: 125,
        СуммаВзаиморасчетов: 125,
        КурсВзаиморасчетов: 1,
        КратностьВзаиморасчетов: 1,
      },
    ]);
  });
});
