import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { aggregateSales, turnoversPath } from "../src/odata/item-sales.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Продажи УТ по регистру «Выручка и себестоимость продаж» на синтетическом корпусе. Метрика приёмки: нетто, продано и
 * возвраты совпадают с ручным расчётом до штуки, исключённый склад не попадает в итог. Данные выдуманы, схема —
 * снимок УТ 11.5.
 */

beforeEach(() => {
  // Окно отчётов о рознице кончается не позже «сегодня»: тесты не зависят от дня запуска.
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date("2026-11-15T12:00:00"));
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const G = (i: number) => uuid("d", i);
const I = (i: number) => uuid("a", i);
const W = (i: number) => uuid("b", i);
const C = (i: number) => uuid("c", i);
const EMPTY = "00000000-0000-0000-0000-000000000000";

const GROUPS = [
  { Ref_Key: G(1), Description: "Телефоны", Parent_Key: EMPTY },
  { Ref_Key: G(2), Description: "Смартфоны", Parent_Key: G(1) },
  { Ref_Key: G(3), Description: "01. Apple", Parent_Key: G(2) },
  { Ref_Key: G(4), Description: "Аксессуары", Parent_Key: EMPTY },
  { Ref_Key: G(5), Description: "01. Apple", Parent_Key: G(4) },
  { Ref_Key: G(9), Description: "Старые модели", Parent_Key: G(3) },
].map((g) => ({ ...g, IsFolder: true }));

const ITEMS = [
  { Ref_Key: I(1), Description: "Смартфон Тест 11", Артикул: "T11", Parent_Key: G(3) },
  { Ref_Key: I(2), Description: "Смартфон Тест 12", Артикул: "T12", Parent_Key: G(3) },
  { Ref_Key: I(3), Description: "Чехол Тест", Артикул: "CASE", Parent_Key: G(5) },
  { Ref_Key: I(4), Description: "Настройка телефона", Артикул: "SRV", Parent_Key: G(4) },
  { Ref_Key: I(5), Description: "Смартфон Тест 8", Артикул: "T8", Parent_Key: G(9) },
].map((i) => ({ ...i, IsFolder: false }));

const WAREHOUSES = [
  { Ref_Key: W(1), Description: "Магазин 1_основной", ТипСклада: "РозничныйМагазин", IsFolder: false },
  { Ref_Key: W(2), Description: "Магазин 1_подменный фонд", ТипСклада: "РозничныйМагазин", IsFolder: false },
  { Ref_Key: W(4), Description: "Офис основной", ТипСклада: "ОптовыйСклад", IsFolder: false },
];

const WH_TYPE = "StandardODATA.Catalog_Склады";
const sale = (item: string, wh: string | null, operation: string, quantity: number, ch = EMPTY) => ({
  АналитикаУчетаНоменклатуры_Key: uuid("e", Math.abs(quantity * 1000) + item.charCodeAt(30)),
  АналитикаУчетаНоменклатуры: { Номенклатура_Key: item, Характеристика_Key: ch },
  Склад: wh ?? "",
  Склад_Type: wh ? WH_TYPE : "StandardODATA.Undefined",
  ХозяйственнаяОперация: operation,
  КоличествоTurnover: quantity,
  // Цена 100,10 за штуку с НДС 20 %: возврат уменьшает выручку так же, как количество.
  СуммаВыручкиTurnover: Math.round(quantity * 100.1 * 100) / 100,
  СуммаВыручкиБезНДСTurnover: Math.round((quantity * 100.1 * 100) / 1.2) / 100,
});

// Ручной расчёт рядом с данными.
const SALES = [
  sale(I(1), W(1), "РеализацияВРозницу", 5),
  sale(I(1), W(2), "РеализацияВРозницу", 2), // подменный фонд
  sale(I(1), W(1), "ВозвратТоваровОтКлиентаПрошлыхПериодов", -1),
  sale(I(1), W(4), "РеализацияКлиенту", 3),
  sale(I(1), W(4), "СторноРеализации", -1),
  sale(I(2), W(1), "РеализацияВРозницу", 4, C(1)),
  sale(I(2), W(1), "РеализацияВРозницу", 1, C(2)),
  sale(I(3), W(1), "РеализацияКомиссионногоТовара", 2),
  sale(I(4), null, "РеализацияКлиенту", 1), // услуга без склада
  sale(I(5), W(1), "РеализацияВРозницу", 1),
  sale(I(5), W(1), "ВозвратТоваровОтКлиентаПрошлыхПериодов", -1), // нетто 0, но движение было
];
// Продано 5+2+3+4+1+2+1+1 = 19; возвращено 1+1+1 = 3; нетто 16.

const REPORTS = [
  { Ref_Key: uuid("f", 1), Склад_Key: W(1), Date: "2026-10-29T22:10:00" },
  { Ref_Key: uuid("f", 2), Склад_Key: W(1), Date: "2026-10-30T23:59:59" },
];

const guidsIn = (text: string) =>
  [...text.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());

function data(seen: FakeQuery[], reports = REPORTS, sales = SALES) {
  return (q: FakeQuery) => {
    seen.push(q);
    const filter = q.params.get("$filter") ?? "";
    if (q.fn === "Turnovers") {
      const condition = q.args["Condition"];
      if (!condition) return sales;
      const allowed = new Set(guidsIn(condition));
      return sales.filter((r) => allowed.has(String(r.Склад).toLowerCase()));
    }
    if (q.base === "Document_ОтчетОРозничныхПродажах") {
      const wh = guidsIn(filter);
      const ge = /Date ge datetime'([^']+)'/.exec(filter)?.[1] ?? "";
      const le = /Date le datetime'([^']+)'/.exec(filter)?.[1] ?? "9999";
      return reports
        .filter((r) => (!wh.length || wh.includes(r.Склад_Key)) && r.Date >= ge && r.Date <= le)
        .sort((a, b) => b.Date.localeCompare(a.Date));
    }
    if (q.base === "Catalog_Склады") {
      const refs = new Set(guidsIn(filter));
      return WAREHOUSES.filter((w) => refs.has(w.Ref_Key.toLowerCase()));
    }
    if (q.base === "Catalog_Номенклатура") {
      if (filter.includes("IsFolder eq true")) return GROUPS;
      const refs = new Set(guidsIn(filter));
      if (filter.includes("Parent_Key eq")) return ITEMS.filter((i) => refs.has(i.Parent_Key.toLowerCase()));
      if (filter.includes("Ref_Key eq"))
        return [...ITEMS, ...GROUPS].filter((i) => refs.has(i.Ref_Key.toLowerCase()));
      if (filter.includes("IsFolder eq false")) return ITEMS;
    }
    return undefined;
  };
}

interface SalesItem {
  ref: string;
  characteristicRef?: string;
  warehouseRef?: string;
  quantity: number;
  sold: number;
  returns: number;
}
interface Sales {
  totalQuantity: number;
  sold: number;
  returns: number;
  count: number;
  items: SalesItem[];
  byOperation?: { operation: string; quantity: number; sold: number; returns: number }[];
  byWarehouse?: { warehouse: string | null; warehouseRef?: string; quantity: number; positions: number }[];
  retailReports?: {
    warehouses: number;
    latest: string | null;
    completeThrough: string | null;
    unknown?: number;
    behind: number;
    oldest: { warehouseRef: string; lastReport: string | null }[];
  };
  warehouseFilter?: { ref: string; type?: string }[];
  note?: string;
  nextOffset?: number;
}

async function setup(
  schema: FakeSchema = "ut11",
  {
    reports = REPORTS,
    sales = SALES,
    env = {},
  }: { reports?: typeof REPORTS; sales?: typeof SALES; env?: Record<string, string> } = {},
) {
  const seen: FakeQuery[] = [];
  const fake = fake1C({ READ_ONLY: "true", ...env }, { schema, data: data(seen, reports, sales) });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (args: object = {}) => {
    const r = (await client.callTool({
      name: "read.analytics.get_sales_register",
      arguments: { from: "2026-10-01", to: "2026-10-31", ...args },
    })) as { isError?: boolean; structuredContent: Sales; content: { text: string }[] };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake, seen };
}

describe("get_sales_register: итоги совпадают с ручным расчётом", () => {
  it("все склады: нетто, продано и возвраты; возвраты и сторно вычитаются", async () => {
    const { call, fake } = await setup();
    const r = await call();
    expect(r.isError, r.text).toBeFalsy();
    const s = r.structuredContent;
    expect([s.totalQuantity, s.sold, s.returns]).toEqual([16, 19, 3]);
    // Выручка нетто: 16 штук × 100,10 = 1601,60; без НДС — сумма строк, каждая округлена до копейки.
    const sx = s as Sales & { revenue: number; revenueExVat: number };
    expect(sx.revenue).toBe(1601.6);
    expect(sx.revenueExVat).toBe(
      SALES.reduce((t, r) => t + Math.round(r.СуммаВыручкиБезНДСTurnover * 100), 0) / 100,
    );
    const i1 = s.items.find((i) => i.ref === I(1))!;
    // 5 + 2 − 1 + 3 − 1.
    expect([i1.quantity, i1.sold, i1.returns]).toEqual([8, 10, 2]);
    // Характеристики — раздельно; нулевое нетто с движением показано.
    expect(s.items.filter((i) => i.ref === I(2)).map((i) => [i.characteristicRef, i.quantity])).toEqual([
      [C(1), 4],
      [C(2), 1],
    ]);
    expect(s.items.find((i) => i.ref === I(5))).toMatchObject({ quantity: 0, sold: 1, returns: 1 });
    expect(s.byOperation).toEqual([
      expect.objectContaining({
        operation: "РеализацияВРозницу",
        quantity: 13,
        sold: 13,
        returns: 0,
        revenue: 1301.3,
      }),
      expect.objectContaining({
        operation: "РеализацияКлиенту",
        quantity: 4,
        sold: 4,
        returns: 0,
        revenue: 400.4,
      }),
      expect.objectContaining({
        operation: "РеализацияКомиссионногоТовара",
        quantity: 2,
        sold: 2,
        returns: 0,
        revenue: 200.2,
      }),
      expect.objectContaining({
        operation: "СторноРеализации",
        quantity: -1,
        sold: 0,
        returns: 1,
        revenue: -100.1,
      }),
      expect.objectContaining({
        operation: "ВозвратТоваровОтКлиентаПрошлыхПериодов",
        quantity: -2,
        sold: 0,
        returns: 2,
        revenue: -200.2,
      }),
    ]);
    expect(fake.violations).toEqual([]);
  });

  it("период целиком: начало в 00:00:00, конец — сам день в 23:59:59, а не полночь следующего", async () => {
    const { call, seen } = await setup();
    await call({ from: "2026-10-01", to: "2026-10-31" });
    const q = seen.find((x) => x.fn === "Turnovers")!;
    expect(q.args["StartPeriod"]).toBe("datetime'2026-10-01T00:00:00'");
    expect(q.args["EndPeriod"]).toBe("datetime'2026-10-31T23:59:59'");
    expect(turnoversPath("R", "2026-12-31", "2026-12-31")).toContain(
      "EndPeriod=datetime'2026-12-31T23:59:59'",
    );
  });

  it("набор складов: исключённый склад и продажи без склада не попадают в итог, отбор — в 1С через cast", async () => {
    const { call, seen, fake } = await setup();
    const s = (await call({ warehouses: [W(1), W(4)], groupBy: "warehouse" })).structuredContent;
    // Основной: I1 5−1, I2 4+1, I3 2, I5 1−1 = 11; офис: 3−1 = 2. Подменный фонд (2) и услуга без склада (1) — нет.
    expect([s.totalQuantity, s.sold, s.returns]).toEqual([13, 16, 3]);
    expect(s.byWarehouse).toEqual([
      expect.objectContaining({ warehouseRef: W(1), quantity: 11, positions: 5 }),
      expect.objectContaining({ warehouseRef: W(4), quantity: 2, positions: 1 }),
    ]);
    expect(s.warehouseFilter?.map((w) => w.type)).toEqual(["РозничныйМагазин", "ОптовыйСклад"]);
    const condition = seen.find((x) => x.fn === "Turnovers")!.args["Condition"]!;
    expect(condition).toContain(`Склад eq cast(guid'${W(1)}', 'Catalog_Склады')`);
    expect(condition).not.toContain(W(2));
    expect(fake.violations).toEqual([]);
  });

  it("без отбора складов продажа без склада входит в итог и видна как warehouse: null", async () => {
    const { call } = await setup();
    const s = (await call({ groupBy: "warehouse" })).structuredContent;
    expect(s.byWarehouse?.find((w) => w.warehouse === null)?.quantity).toBe(1);
    expect(s.note).toContain("без склада");
  });

  it("группа по пути с подгруппами и итоги по операциям только по ней", async () => {
    const { call } = await setup();
    const s = (await call({ group: "Телефоны/Смартфоны/01. Apple" })).structuredContent;
    // I1 8, I2 5, I5 0 (подгруппа «Старые модели»); чехол и услуга не входят.
    expect([s.totalQuantity, s.sold, s.returns]).toEqual([13, 16, 3]);
    expect(s.byOperation?.find((o) => o.operation === "РеализацияКомиссионногоТовара")).toBeUndefined();
  });

  it("одноимённые группы — отказ со списком путей", async () => {
    const { call } = await setup();
    const r = await call({ group: "01. Apple" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Телефоны/Смартфоны/01. Apple");
    expect(r.text).toContain("Аксессуары/01. Apple");
  });

  it("отчёты о рознице: отставание — от самого свежего отчёта, а не от конца периода", async () => {
    const { call } = await setup();
    const s = (await call()).structuredContent;
    // Розница — на W(1) и W(2); отчёт есть только у W(1), последний 30.10. Отстаёт только W(2): у него отчётов нет.
    expect(s.retailReports).toEqual({
      warehouses: 2,
      latest: "2026-10-30",
      completeThrough: null,
      behind: 1,
      oldest: [expect.objectContaining({ warehouseRef: W(2), lastReport: null })],
    });
  });

  it("период по сегодня и в будущем: склады, отчитавшиеся вчера, не отстают; лишних запросов нет (ревью)", async () => {
    // «Сегодня» — 15.11; оба склада отчитались по 14.11, сегодняшняя смена ещё не закрыта.
    const reports = [W(1), W(2)].map((w, i) => ({
      Ref_Key: uuid("f", 60 + i),
      Склад_Key: w,
      Date: "2026-11-14T22:00:00",
    }));
    const { call, seen } = await setup("ut11", { reports });
    for (const to of ["2026-11-15", "2026-11-30"]) {
      const before = seen.filter((q) => q.base === "Document_ОтчетОРозничныхПродажах").length;
      const s = (await call({ from: "2026-11-01", to })).structuredContent;
      expect(s.retailReports, to).toEqual({
        warehouses: 2,
        latest: "2026-11-14",
        completeThrough: "2026-11-14",
        behind: 0,
        oldest: [],
      });
      // Одно окно на неделю, без поиска по одному складу.
      expect(seen.filter((q) => q.base === "Document_ОтчетОРозничныхПродажах").length - before).toBe(1);
    }
  });

  it("все склады отчитались по конец периода — completeThrough = to, список отстающих пуст", async () => {
    const reports = [W(1), W(2)].map((w, i) => ({
      Ref_Key: uuid("f", 50 + i),
      Склад_Key: w,
      Date: "2026-10-31T22:00:00",
    }));
    const { call } = await setup("ut11", { reports });
    expect((await call()).structuredContent.retailReports).toEqual({
      warehouses: 2,
      latest: "2026-10-31",
      completeThrough: "2026-10-31",
      behind: 0,
      oldest: [],
    });
  });

  it("отчётов за год больше лимита строк — продажи всё равно считаются (ревью: вызов падал целиком)", async () => {
    // 400 отчётов за год по W(1), лимит строк 100: читается только последняя неделя и один запрос на отставший склад.
    // Отчёты за 2025 год, период в прошлом: окно — последняя неделя декабря, где отчётов нет.
    const reports = Array.from({ length: 400 }, (_, i) => {
      const d = new Date(Date.UTC(2025, 0, 1) + i * 19 * 3600_000).toISOString().slice(0, 19);
      return { Ref_Key: uuid("f", 1000 + i), Склад_Key: W(1), Date: d };
    });
    const { call } = await setup("ut11", { reports, env: { ODATA_ANALYTICS_MAX_ROWS: "100" } });
    const r = await call({ from: "2025-01-01", to: "2025-12-31" });
    expect(r.isError, r.text).toBeFalsy();
    expect(r.structuredContent.totalQuantity).toBe(16);
    const last = reports
      .map((x) => x.Date.slice(0, 10))
      .sort()
      .at(-1);
    expect(r.structuredContent.retailReports?.latest).toBe(last);
  });

  it("отчётов даже за неделю больше лимита — продажи с примечанием, а не отказ", async () => {
    const reports = Array.from({ length: 150 }, (_, i) => ({
      Ref_Key: uuid("f", 3000 + i),
      Склад_Key: W(1),
      Date: "2026-10-30T10:00:00",
    }));
    const { call } = await setup("ut11", { reports, env: { ODATA_ANALYTICS_MAX_ROWS: "100" } });
    const r = await call();
    expect(r.isError, r.text).toBeFalsy();
    expect(r.structuredContent.totalQuantity).toBe(16);
    expect(r.structuredContent.retailReports).toMatchObject({ completeThrough: null, unknown: 2 });
    expect(r.structuredContent.note).toContain("узнать не удалось");
  });

  it("неизвестный склад, перевёрнутый период — отказ", async () => {
    const { call } = await setup();
    const unknown = await call({ warehouses: [W(77)] });
    expect(unknown.isError).toBe(true);
    expect(unknown.text).toContain(W(77));
    const swapped = await call({ from: "2026-10-31", to: "2026-10-01" });
    expect(swapped.isError).toBe(true);
    expect(swapped.text).toContain("позже конца");
  });
});

describe("get_sales_register: поле returned занято обёрткой страниц", () => {
  it("на короткой странице возвраты не затираются числом строк (на живой базе было 1 вместо 66)", async () => {
    const { call } = await setup();
    const s = (await call({ limit: 1 })).structuredContent as Sales & { returned?: number };
    expect([s.totalQuantity, s.sold, s.returns]).toEqual([16, 19, 3]);
    // Обёртка пишет в returned число строк этого ответа — поле не наше.
    expect(s.returned).toBe(1);
    expect(s.items).toHaveLength(1);
  });
});

describe("get_sales_register: каждый параметр меняет результат", () => {
  it("warehouses, group, query, groupBy, период", async () => {
    const { call } = await setup();
    const sig = (s: Sales) =>
      JSON.stringify([s.totalQuantity, s.count, s.items.map((i) => [i.ref, i.warehouseRef])]);
    const base = sig((await call()).structuredContent);
    for (const args of [
      { warehouses: [W(4)] },
      { group: "Аксессуары" },
      { query: "t12" },
      { groupBy: "warehouse" },
    ]) {
      expect(sig((await call(args)).structuredContent), JSON.stringify(args)).not.toBe(base);
    }
  });
});

describe("get_sales_register: страницы", () => {
  it("offset даёт все позиции ровно по разу, итоги — по всему отбору", async () => {
    const { call } = await setup();
    const seen: string[] = [];
    let offset = 0;
    for (let guard = 0; guard < 10; guard++) {
      const s = (await call({ limit: 2, offset })).structuredContent;
      expect(s.totalQuantity).toBe(16);
      seen.push(...s.items.map((i) => `${i.ref}|${i.characteristicRef ?? ""}`));
      if (s.nextOffset === undefined) break;
      offset = s.nextOffset;
    }
    expect(new Set(seen).size).toBe(seen.length);
    expect(seen).toHaveLength(6);
  });
});

describe("get_sales_register: неопубликованный регистр", () => {
  it("в Бухгалтерии — понятный отказ с именем регистра", async () => {
    const { call } = await setup("bp30");
    const r = await call();
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Выручка и себестоимость продаж");
  });
});

describe("aggregateSales", () => {
  it("строка с нулём продаж и возвратов не показывается, дробные количества без хвостов float", () => {
    const lines = aggregateSales(
      [
        { itemRef: I(1), operation: "x", quantity: 0.1, revenueCents: 10, revenueExVatCents: 8 },
        { itemRef: I(1), operation: "x", quantity: 0.2, revenueCents: 20, revenueExVatCents: 17 },
        { itemRef: I(2), operation: "x", quantity: 0, revenueCents: 0, revenueExVatCents: 0 },
      ],
      false,
    );
    expect(lines).toEqual([
      { itemRef: I(1), quantity: 0.3, sold: 0.3, returns: 0, revenue: 0.3, revenueExVat: 0.25 },
    ]);
  });
});
