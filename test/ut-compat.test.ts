import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { contractField, operationField, TAX_PAYMENT_OPERATION } from "../src/odata/doc-fields.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/**
 * Инструменты Бухгалтерии на схеме Управления торговлей 11: те же факты, другие поля. На живой УТ без этих правок
 * find_counterparty и get_sales_breakdown падали с 400, а get_taxes_paid молча давал 0. Данные синтетические.
 */

afterEach(() => {
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const CP = (i: number) => uuid("c", i);

const COUNTERPARTIES = [
  {
    Ref_Key: CP(1),
    Description: "ООО Тест",
    ИНН: "7700000001",
    ЮрФизЛицо: "ЮрЛицо",
    НаименованиеПолное: "ООО «Тест»",
  },
  {
    Ref_Key: CP(2),
    Description: "ИП Тестов",
    ИНН: "770000000002",
    ЮрФизЛицо: "ИндивидуальныйПредприниматель",
  },
  { Ref_Key: CP(3), Description: "Test Ltd", ЮрФизЛицо: "ЮрЛицоНеРезидент" },
];
const CONTRACTS = [
  { Контрагент_Key: CP(1), ТипДоговора: "СПокупателем" },
  { Контрагент_Key: CP(1), ТипДоговора: "СПоставщиком" },
  { Контрагент_Key: CP(2), ТипДоговора: "СКомитентом" },
];

const guidsIn = (text: string) =>
  [...text.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());

async function setup(data: (q: FakeQuery) => Record<string, unknown>[] | undefined) {
  const seen: FakeQuery[] = [];
  const fake = fake1C(
    { READ_ONLY: "true" },
    {
      schema: "ut11",
      data: (q) => {
        seen.push(q);
        return data(q);
      },
    },
  );
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (name: string, args: object = {}) => {
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
      _meta?: Record<string, string>;
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake, seen };
}

describe("find_counterparty в УТ: без кода и групп, роли — из ТипДоговора", () => {
  const data = (q: FakeQuery) => {
    const filter = q.params.get("$filter") ?? "";
    if (q.base === "Catalog_Контрагенты") return COUNTERPARTIES;
    if (q.base === "Catalog_ДоговорыКонтрагентов") {
      const owners = new Set(guidsIn(filter));
      return CONTRACTS.filter((c) => owners.has(c.Контрагент_Key));
    }
    return undefined;
  };

  it("находит, не запрашивая Code и IsFolder, роли из договоров по Контрагент_Key", async () => {
    const { call, fake, seen } = await setup(data);
    const r = await call("read.counterparty.find_counterparty", { query: "Тест" });
    expect(r.isError, r.text).toBeFalsy();
    expect(fake.violations).toEqual([]);
    const select = seen.find((q) => q.base === "Catalog_Контрагенты")!.params.get("$select")!;
    expect(select).not.toMatch(/\bCode\b|IsFolder|Parent_Key/);
    const rows = r.structuredContent["rows"] as { ref: string; roles: string[]; contracts: number }[];
    expect(rows.find((x) => x.ref === CP(1))).toMatchObject({
      roles: ["покупатель", "поставщик"],
      contracts: 2,
    });
    expect(rows.find((x) => x.ref === CP(2))).toMatchObject({ roles: ["комитент"], contracts: 1 });
  });

  it("поиск групп — понятный отказ: в УТ группы у партнёров", async () => {
    const { call } = await setup(data);
    const r = await call("read.counterparty.find_counterparty", { query: "Тест", subjectKind: "group" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("без групп");
  });
});

describe("категория контрагента в УТ — по ЮрФизЛицо", () => {
  const sales = [
    { Date: "2026-10-05T10:00:00", СуммаДокумента: 100, Контрагент_Key: CP(1), Договор_Key: uuid("d", 1) },
    { Date: "2026-10-06T10:00:00", СуммаДокумента: 40, Контрагент_Key: CP(2), Договор_Key: uuid("d", 2) },
    { Date: "2026-10-07T10:00:00", СуммаДокумента: 7, Контрагент_Key: CP(3), Договор_Key: uuid("d", 3) },
  ];
  const data = (q: FakeQuery) => {
    const filter = q.params.get("$filter") ?? "";
    if (q.base === "Catalog_Контрагенты") {
      const kinds = [...filter.matchAll(/ЮрФизЛицо eq '([^']+)'/g)].map((m) => m[1]);
      return COUNTERPARTIES.filter((c) => kinds.includes(c.ЮрФизЛицо)).map((c) => ({ Ref_Key: c.Ref_Key }));
    }
    if (q.base === "Document_РеализацияТоваровУслуг") {
      const contract = guidsIn(filter);
      return sales.filter((d) => !contract.length || contract.includes(d.Договор_Key));
    }
    return undefined;
  };

  it("ИП, ЮрЛицо (с нерезидентами) и ФизЛицо дают разные суммы; Госорган — отказ", async () => {
    const { call, fake } = await setup(data);
    const total = async (counterpartyKind: string) =>
      (
        await call("read.analytics.get_sales_breakdown", {
          from: "2026-10-01",
          to: "2026-10-31",
          counterpartyKind,
        })
      ).structuredContent["total"];
    expect(await total("ИП")).toBe(40);
    expect(await total("ЮрЛицо")).toBe(107);
    expect(await total("Нерезидент")).toBe(7);
    expect(await total("ФизЛицо")).toBe(0);
    const gov = await call("read.analytics.get_sales_breakdown", {
      from: "2026-10-01",
      to: "2026-10-31",
      counterpartyKind: "Госорган",
    });
    expect(gov.isError).toBe(true);
    expect(gov.text).toContain("госоргана");
    expect(fake.violations).toEqual([]);
  });

  it("договор — Договор_Key; разбивка по договору и отбор работают, примечание про розницу", async () => {
    const { call, seen } = await setup(data);
    const r = await call("read.analytics.get_sales_breakdown", {
      from: "2026-10-01",
      to: "2026-10-31",
      contractRef: uuid("d", 2),
    });
    expect(r.structuredContent["total"]).toBe(40);
    expect(seen.find((q) => q.base === "Document_РеализацияТоваровУслуг")!.params.get("$filter")).toContain(
      "Договор_Key eq guid",
    );
    expect(r.structuredContent["note"]).toContain("розница");
  });
});

describe("деньги в УТ: ХозяйственнаяОперация вместо ВидОперации", () => {
  it("налоги — ПеречислениеВБюджет, а не тихий ноль", async () => {
    const { call, seen } = await setup((q) =>
      q.base === "Document_СписаниеБезналичныхДенежныхСредств"
        ? [{ Date: "2026-10-10T10:00:00", СуммаДокумента: 500, Контрагент: CP(1), Контрагент_Type: "x" }]
        : undefined,
    );
    const r = await call("read.analytics.get_taxes_paid", { from: "2026-10-01", to: "2026-10-31" });
    expect(r.isError, r.text).toBeFalsy();
    expect(r.structuredContent["total"]).toBe(500);
    const filter = seen
      .find((q) => q.base === "Document_СписаниеБезналичныхДенежныхСредств")!
      .params.get("$filter");
    expect(filter).toContain("ХозяйственнаяОперация eq 'ПеречислениеВБюджет'");
  });

  it("поля по схеме: вид операции, маркер налога, договор (в т.ч. составной через cast)", () => {
    expect(operationField(new Set(["ВидОперации"]))).toBe("ВидОперации");
    expect(operationField(new Set(["ХозяйственнаяОперация"]))).toBe("ХозяйственнаяОперация");
    expect(TAX_PAYMENT_OPERATION["ХозяйственнаяОперация"]).toBe("ПеречислениеВБюджет");
    const ref = uuid("d", 9);
    expect(
      contractField(new Set(["ДоговорКонтрагента_Key"]), "Catalog_ДоговорыКонтрагентов")!.filter(ref),
    ).toBe(`ДоговорКонтрагента_Key eq guid'${ref}'`);
    expect(
      contractField(new Set(["Договор", "Договор_Type"]), "Catalog_ДоговорыКонтрагентов")!.filter(ref),
    ).toBe(`Договор eq cast(guid'${ref}', 'Catalog_ДоговорыКонтрагентов')`);
    expect(contractField(new Set(["Прочее"]), "Catalog_ДоговорыКонтрагентов")).toBeUndefined();
  });
});

describe("инструменты бухучёта в УТ — отказ без совета публиковать Хозрасчётный", () => {
  it.each(["read.analytics.get_debtors", "read.analytics.get_cash_balance"])("%s", async (name) => {
    const { call } = await setup(() => undefined);
    const r = await call(name, {});
    expect(r.isError).toBe(true);
    expect(r._meta?.["odata-1c/errorKind"]).toBe("not_published");
    expect(r.text).toContain("Управление торговлей");
    expect(r.text).toContain("read.analytics.get_warehouse_stock");
    expect(r.text).not.toContain("Добавьте их в 1С");
  });
});
