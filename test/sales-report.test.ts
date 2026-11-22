import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { aggregate, buckets, minusYear, previousPeriod } from "../src/odata/sales-report.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * 1.1: продажи в разрезах (get_sales_report). Итоги — с ручным расчётом рядом с данными; УТ — обороты регистра продаж
 * со свёрткой Dimensions, Бухгалтерия — обороты счетов 90. Данные выдуманы, схемы — снимки УТ 11.5 и БП 3.0.
 */
const g = (k: string, n: number) => `${k.repeat(8)}-0000-4000-8000-${String(n).padStart(12, "0")}`;
const D1 = g("d", 1);
const D2 = g("d", 2);
const M1 = g("e", 1);
const W1 = g("f", 1);
const I1 = g("a", 1);
const I2 = g("a", 2);
const P1 = g("b", 1);

describe("части периода и сравнение", () => {
  it("месяцы, недели с понедельника, кварталы; края обрезаны по периоду", () => {
    expect(buckets("2026-09-15", "2026-11-02", "month").map((b) => [b.label, b.from, b.to])).toEqual([
      ["2026-09", "2026-09-15", "2026-09-30"],
      ["2026-10", "2026-10-01", "2026-10-31"],
      ["2026-11", "2026-11-01", "2026-11-02"],
    ]);
    expect(buckets("2026-11-04", "2026-11-12", "week").map((b) => [b.from, b.to])).toEqual([
      ["2026-11-04", "2026-11-08"],
      ["2026-11-09", "2026-11-12"],
    ]);
    expect(buckets("2026-01-01", "2026-12-31", "quarter").map((b) => b.label)).toEqual([
      "2026-Q1",
      "2026-Q2",
      "2026-Q3",
      "2026-Q4",
    ]);
    expect(() => buckets("2026-01-01", "2026-12-31", "day")).toThrow(/Слишком много частей/);
    expect(() => buckets("2026-02-01", "2026-01-01", "month")).toThrow(/позже конца/);
  });

  it("прошлый период: целые месяцы — столько же месяцев назад, иначе столько же дней; год назад — 29.02 → 28.02", () => {
    expect(previousPeriod("2026-10-01", "2026-10-31")).toEqual({ from: "2026-09-01", to: "2026-09-30" });
    expect(previousPeriod("2026-10-01", "2026-12-31")).toEqual({ from: "2026-07-01", to: "2026-09-30" });
    expect(previousPeriod("2026-03-01", "2026-03-31")).toEqual({ from: "2026-02-01", to: "2026-02-28" });
    expect(previousPeriod("2026-11-10", "2026-11-16")).toEqual({ from: "2026-11-03", to: "2026-11-09" });
    expect(minusYear("2024-02-29")).toBe("2023-02-28");
    expect(minusYear("2026-10-31")).toBe("2025-10-31");
  });

  it("свёртка: метрика, которую нельзя разделить (undefined), не превращается в ноль", () => {
    const rows = aggregate(
      [
        {
          dims: { department: D1 },
          quantity: 1,
          revenueCents: 100,
          revenueExVatCents: 80,
          costCents: 50,
          costExVatCents: 40,
        },
        {
          dims: { department: D1 },
          quantity: 2,
          revenueCents: 200,
          revenueExVatCents: undefined,
          costCents: 60,
          costExVatCents: 50,
        },
      ],
      ["department"],
    );
    expect(rows).toEqual([
      expect.objectContaining({
        quantity: 3,
        revenueCents: 300,
        revenueExVatCents: undefined,
        costCents: 110,
      }),
    ]);
  });
});

// ─── Управление торговлей ───

const sale = (o: {
  dep: string;
  mgr?: string;
  item: string;
  q: number;
  rev: number;
  cost: number;
  partner?: string;
  wh?: string;
}) => ({
  Подразделение_Key: o.dep,
  Менеджер_Key: o.mgr ?? M1,
  Склад: o.wh ?? W1,
  Склад_Type: "StandardODATA.Catalog_Склады",
  ХозяйственнаяОперация: "РеализацияКлиенту",
  АналитикаУчетаНоменклатуры: { Номенклатура_Key: o.item },
  АналитикаУчетаПоПартнерам: { Партнер_Key: o.partner ?? P1, Организация_Key: g("c", 1) },
  КоличествоTurnover: o.q,
  СуммаВыручкиTurnover: o.rev,
  СуммаВыручкиБезНДСTurnover: Math.round((o.rev / 1.2) * 100) / 100,
  СтоимостьTurnover: o.cost,
  СтоимостьБезНДСTurnover: Math.round((o.cost / 1.2) * 100) / 100,
  ДопРасходыTurnover: 0,
  ДопРасходыБезНДСTurnover: 0,
});
// Октябрь: D1 — 2 строки (выручка 300), D2 — 1 строка (выручка 50). Сентябрь: D1 — 100, D2 — 100.
const OCT = [
  sale({ dep: D1, item: I1, q: 2, rev: 200, cost: 120 }),
  sale({ dep: D1, item: I2, q: 1, rev: 100, cost: 70 }),
  sale({ dep: D2, item: I1, q: 1, rev: 50, cost: 40 }),
];
const SEP = [
  sale({ dep: D1, item: I1, q: 1, rev: 100, cost: 60 }),
  sale({ dep: D2, item: I2, q: 1, rev: 100, cost: 60 }),
];

async function tradeSetup(schema: FakeSchema = "ut11", byStart?: (start: string) => object[]) {
  const seen: FakeQuery[] = [];
  const fake = fake1C(
    { READ_ONLY: "true" },
    {
      schema,
      data: (q) => {
        seen.push(q);
        if (q.base === "AccumulationRegister_ВыручкаИСебестоимостьПродаж" && q.fn === "Turnovers" && byStart)
          return byStart(String(q.args["StartPeriod"]));
        if (q.base === "AccumulationRegister_ВыручкаИСебестоимостьПродаж" && q.fn === "Turnovers")
          return String(q.args["StartPeriod"]).includes("2026-10")
            ? OCT
            : String(q.args["StartPeriod"]).includes("2026-09")
              ? SEP
              : [];
        return undefined;
      },
    },
  );
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (args: object) => {
    const r = (await client.callTool({
      name: "read.analytics.get_sales_report",
      arguments: { from: "2026-10-01", to: "2026-10-31", ...args },
    })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown> & { items: Array<Record<string, unknown>> };
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake, seen };
}

describe("get_sales_report в Управлении торговлей", () => {
  it("по подразделениям: выручка, себестоимость и маржа совпадают с ручным расчётом; свёртка Dimensions на стороне 1С", async () => {
    const { call, fake, seen } = await tradeSetup();
    const r = await call({ groupBy: ["department"] });
    expect(r.isError, r.text).toBeFalsy();
    const s = r.structuredContent;
    expect(s["configuration"]).toBe("trade");
    expect([s["revenue"], s["cost"], s["margin"], s["quantity"]]).toEqual([350, 230, 120, 4]);
    expect(r.structuredContent.items.map((i) => [i["departmentRef"], i["revenue"], i["margin"]])).toEqual([
      [D1, 300, 110],
      [D2, 50, 10],
    ]);
    const q = seen.find((x) => x.fn === "Turnovers")!;
    expect(q.args["Dimensions"]).toBe("Подразделение");
    expect(fake.violations).toEqual([]);
  });

  it("подразделение × месяц и сравнение с прошлым периодом: prev и рост выручки по строке", async () => {
    const { call } = await tradeSetup();
    const r = await call({
      groupBy: ["department", "period"],
      period: "month",
      compareWith: "previousPeriod",
    });
    expect(r.isError, r.text).toBeFalsy();
    const rows = r.structuredContent.items.map((i) => [
      i["departmentRef"],
      i["period"],
      i["revenue"],
      (i["prev"] as { revenue: number } | null)?.revenue ?? null,
      i["growth"],
    ]);
    expect(rows).toEqual([
      [D1, "2026-10", 300, 100, 200],
      [D2, "2026-10", 50, 100, -50],
    ]);
    expect(r.structuredContent["compare"]).toMatchObject({
      from: "2026-09-01",
      to: "2026-09-30",
      revenue: 200,
    });
  });

  it("sortBy margin и abc; отбор departments; неподдерживаемое сочетание — отказ схемы", async () => {
    const { call } = await tradeSetup();
    const r = await call({ groupBy: ["item"], sortBy: "margin", abc: true });
    const items = r.structuredContent.items;
    // I1: 250 − 160 = 90; I2: 100 − 70 = 30. Класс — по накопленной доле с самой позицией: I1 — 75 % → A, I2 — 100 % → C.
    expect(items.map((i) => [i["itemRef"], i["margin"], i["abcClass"]])).toEqual([
      [I1, 90, "A"],
      [I2, 30, "C"],
    ]);
    const only = await call({ groupBy: ["item"], departments: [D2] });
    expect(only.structuredContent["revenue"]).toBe(50);
  });

  // Ревью 1.1: у марта 6 недельных частей, у апреля 5. Лишняя мартовская часть «2026-03-30 (неделя)» сохраняла метку,
  // совпадавшую с меткой первой апрельской части, и prev строки складывал 1 марта и 30–31 марта. Старые тесты
  // сравнивали месяц с месяцем — частей поровну, лишней не бывало.
  it("сравнение по неделям: лишняя часть прошлого периода не складывается с чужой строкой — она в «ушедших»", async () => {
    const { call } = await tradeSetup("ut11", (start) =>
      start.includes("2026-03-01")
        ? [sale({ dep: D1, item: I1, q: 1, rev: 1, cost: 0 })]
        : start.includes("2026-03-30")
          ? [sale({ dep: D1, item: I1, q: 1, rev: 61, cost: 0 })]
          : start.includes("2026-04-01")
            ? [sale({ dep: D1, item: I1, q: 1, rev: 10, cost: 0 })]
            : [],
    );
    const r = await call({
      from: "2026-04-01",
      to: "2026-04-30",
      groupBy: ["period"],
      period: "week",
      compareWith: "previousPeriod",
    });
    expect(r.isError, r.text).toBeFalsy();
    const first = r.structuredContent.items.find((i) => i["period"] === "2026-03-30 (неделя)")!;
    expect((first["prev"] as { revenue: number }).revenue).toBe(1);
    expect(r.structuredContent["compare"]).toMatchObject({ revenue: 62 });
    expect(String(r.structuredContent["note"])).toContain("только в прошлом периоде");
  });
});

// ─── Бухгалтерия ───

const A = { r: g("1", 1), v: g("1", 2), c: g("1", 3) };
const NG = g("2", 1);
const BP_ROWS = [
  {
    Account_Key: A.r,
    ExtDimension1: NG,
    ExtDimension1_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
    ExtDimension3: I1,
    ExtDimension3_Type: "StandardODATA.Catalog_Номенклатура",
    СуммаTurnoverCr: 1200,
    СуммаTurnoverDr: 0,
    КоличествоTurnoverCr: 3,
  },
  // Реформация баланса в конце года: Дт 90.01 — не вычитается из выручки.
  {
    Account_Key: A.r,
    ExtDimension1: NG,
    ExtDimension1_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
    СуммаTurnoverCr: 0,
    СуммаTurnoverDr: 1200,
  },
  {
    Account_Key: A.v,
    ExtDimension1: NG,
    ExtDimension1_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
    СуммаTurnoverDr: 200,
    СуммаTurnoverCr: 0,
  },
  {
    Account_Key: A.c,
    ExtDimension1: NG,
    ExtDimension1_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
    СуммаTurnoverDr: 700,
    СуммаTurnoverCr: 0,
  },
];
async function bpSetup() {
  const fake = fake1C(
    { READ_ONLY: "true" },
    {
      schema: "bp30",
      data: (q) =>
        q.base === "ChartOfAccounts_Хозрасчетный" && !q.fn
          ? [
              {
                Ref_Key: A.r,
                Code: "90.01.1",
                Description: "Выручка",
                ЗапретитьИспользоватьВПроводках: false,
              },
              { Ref_Key: A.v, Code: "90.03", Description: "НДС", ЗапретитьИспользоватьВПроводках: false },
              {
                Ref_Key: A.c,
                Code: "90.02.1",
                Description: "Себестоимость",
                ЗапретитьИспользоватьВПроводках: false,
              },
            ]
          : q.fn === "BalanceAndTurnovers"
            ? BP_ROWS.map((r) => ({
                СуммаOpeningBalanceDr: 0,
                СуммаOpeningBalanceCr: 0,
                СуммаClosingBalanceDr: 0,
                СуммаClosingBalanceCr: 0,
                ...r,
              }))
            : q.base === "Catalog_НоменклатурныеГруппы"
              ? [{ Ref_Key: NG, Description: "Товары" }]
              : undefined,
    },
  );
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  return async (args: object) => {
    const r = (await client.callTool({
      name: "read.analytics.get_sales_report",
      arguments: { from: "2026-10-01", to: "2026-10-31", ...args },
    })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown> & { items: Array<Record<string, unknown>> };
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
}

describe("get_sales_report в Бухгалтерии — счёт 90", () => {
  it("по номенклатурной группе: Кт 90.01, НДС Дт 90.03, себестоимость Дт 90.02; реформация не обнуляет выручку", async () => {
    const call = await bpSetup();
    const r = await call({ groupBy: ["itemGroup"] });
    expect(r.isError, r.text).toBeFalsy();
    expect(r.structuredContent.items).toEqual([
      expect.objectContaining({
        itemGroup: "Товары",
        itemGroupRef: NG,
        revenue: 1200,
        revenueExVat: 1000,
        cost: 700,
        margin: 500,
        marginExVat: 300,
        quantity: 3,
      }),
    ]);
  });

  it("по номенклатуре: НДС и себестоимость без номенклатуры не делятся — не выводятся, сказано в note", async () => {
    const call = await bpSetup();
    const r = await call({ groupBy: ["item"] });
    expect(r.isError, r.text).toBeFalsy();
    const it0 = r.structuredContent.items.find((i) => i["itemRef"] === I1)!;
    expect(it0["revenue"]).toBe(1200);
    expect(it0).not.toHaveProperty("margin");
    expect(it0).not.toHaveProperty("revenueExVat");
    expect(String(r.structuredContent["note"])).toContain("без нужного субконто");
  });

  // Ревью 1.1: у метрики без субконто разреза все строки сортировались нулями — порядок по ключу, класс C у всех, без
  // отказа и заметки. Старые тесты Бухгалтерии не задавали sortBy, а в УТ все метрики есть.
  it("sortBy по метрике, которой в разрезе нет (маржа по номенклатуре), — отказ, а не нули и класс C", async () => {
    const call = await bpSetup();
    for (const sortBy of ["margin", "marginExVat", "revenueExVat"]) {
      const r = await call({ groupBy: ["item"], sortBy, abc: true });
      expect(r.isError, sortBy).toBe(true);
      expect(r.text).toContain("этой метрики в выбранном разрезе нет");
    }
    const ok = await call({ groupBy: ["itemGroup"], sortBy: "margin", abc: true });
    expect(ok.isError, ok.text).toBeFalsy();
    expect(ok.structuredContent.items[0]).toMatchObject({ margin: 500, abcClass: "A" });
  });

  it("разрезы УТ в Бухгалтерии — отказ с подсказкой get_sales_breakdown", async () => {
    const call = await bpSetup();
    const r = await call({ groupBy: ["partner"] });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("get_sales_breakdown");
  });
});

describe("get_purchases_report — закупки в разрезах (УТ)", () => {
  const S1 = g("5", 1);
  const S2 = g("5", 2);
  const buy = (partner: string, item: string, q: number, sum: number, disc = 0) => ({
    Партнер_Key: partner,
    Организация_Key: g("c", 1),
    АналитикаУчетаНоменклатуры: { Номенклатура_Key: item },
    Склад: W1,
    Склад_Type: "StandardODATA.Catalog_Склады",
    КоличествоTurnover: q,
    СуммаTurnover: sum,
    СуммаБезНДСTurnover: Math.round((sum / 1.2) * 100) / 100,
    СуммаСкидкиTurnover: disc,
  });
  const OCT_BUY = [buy(S1, I1, 5, 600, 10), buy(S1, I2, 1, 100), buy(S2, I1, 2, 300, 5)];
  async function setup() {
    const seen: FakeQuery[] = [];
    const fake = fake1C(
      { READ_ONLY: "true" },
      {
        schema: "ut11",
        data: (q) => {
          seen.push(q);
          if (q.base === "AccumulationRegister_Закупки" && q.fn === "Turnovers")
            return String(q.args["StartPeriod"]).includes("2026-10") ? OCT_BUY : [];
          return undefined;
        },
      },
    );
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const call = async (args: object) => {
      const r = (await client.callTool({
        name: "read.analytics.get_purchases_report",
        arguments: { from: "2026-10-01", to: "2026-10-31", ...args },
      })) as {
        isError?: boolean;
        structuredContent: Record<string, unknown> & { items: Array<Record<string, unknown>> };
        content: { text: string }[];
      };
      return { ...r, text: r.content.map((c) => c.text).join("") };
    };
    return { call, fake, seen };
  }

  it("по поставщикам: сумма, количество и скидка; поставщик — прямое измерение регистра", async () => {
    const { call, fake, seen } = await setup();
    const r = await call({ groupBy: ["partner"] });
    expect(r.isError, r.text).toBeFalsy();
    expect([
      r.structuredContent["amount"],
      r.structuredContent["quantity"],
      r.structuredContent["discount"],
    ]).toEqual([1000, 8, 15]);
    expect(r.structuredContent.items.map((i) => [i["partnerRef"], i["amount"], i["discount"]])).toEqual([
      [S1, 700, 10],
      [S2, 300, 5],
    ]);
    expect(r.structuredContent).not.toHaveProperty("margin");
    expect(seen.find((q) => q.fn === "Turnovers")!.args["Dimensions"]).toBe("Партнер");
    expect(fake.violations).toEqual([]);
  });

  it("sortBy amount и abc; в Бухгалтерии — отказ с подсказкой get_purchases_breakdown", async () => {
    const { call } = await setup();
    const r = await call({ groupBy: ["item"], sortBy: "amount", abc: true });
    expect(r.structuredContent.items.map((i) => [i["itemRef"], i["amount"], i["abcClass"]])).toEqual([
      [I1, 900, "A"],
      [I2, 100, "C"],
    ]);
    const bp = fake1C({ READ_ONLY: "true" }, { schema: "bp30" });
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(bp.ctx).connect(a);
    const c = new Client({ name: "t", version: "1" });
    await c.connect(b);
    const x = (await c.callTool({
      name: "read.analytics.get_purchases_report",
      arguments: { from: "2026-10-01", to: "2026-10-31", groupBy: ["partner"] },
    })) as { isError?: boolean; content: { text: string }[] };
    expect(x.isError).toBe(true);
    expect(x.content[0]!.text).toContain("get_purchases_breakdown");
  });
});

describe("скидки в разрезах продаж (УТ)", () => {
  it("discount — ручная и автоматическая скидка из регистра продаж; sortBy discount", async () => {
    const withDisc = OCT.map((r, i) => ({
      ...r,
      СуммаРучнойСкидкиTurnover: i + 1,
      СуммаАвтоматическойСкидкиTurnover: 1,
    }));
    const fake = fake1C(
      { READ_ONLY: "true" },
      {
        schema: "ut11",
        data: (q) =>
          q.base === "AccumulationRegister_ВыручкаИСебестоимостьПродаж" && q.fn === "Turnovers"
            ? withDisc
            : undefined,
      },
    );
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const c = new Client({ name: "t", version: "1" });
    await c.connect(b);
    const r = (await c.callTool({
      name: "read.analytics.get_sales_report",
      arguments: { from: "2026-10-01", to: "2026-10-31", groupBy: ["department"], sortBy: "discount" },
    })) as { structuredContent: Record<string, unknown> & { items: Array<Record<string, unknown>> } };
    // Скидки: строки 1+1, 2+1, 3+1 → D1 (строки 1, 2) = 5, D2 (строка 3) = 4; итог 9.
    expect(r.structuredContent["discount"]).toBe(9);
    expect(r.structuredContent.items.map((i) => [i["departmentRef"], i["discount"]])).toEqual([
      [D1, 5],
      [D2, 4],
    ]);
  });
});
