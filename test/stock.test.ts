import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { groupPath, resolveItemGroup, type ItemGroup } from "../src/odata/inventory.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Остатки УТ по регистру «Товары на складах» на синтетическом корпусе. Метрика приёмки: итоги совпадают с ручным
 * расчётом до штуки, а исключённый склад не попадает в итог. Имена и данные выдуманы, схема — снимок УТ 11.5.
 */

afterEach(() => {
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const G = (i: number) => uuid("d", i);
const I = (i: number) => uuid("a", i);
const W = (i: number) => uuid("b", i);
const C = (i: number) => uuid("c", i);
const EMPTY = "00000000-0000-0000-0000-000000000000";

// Две ветки с одинаковым «01. Apple» и третья, где путь «Смартфоны/01. Apple» повторяется.
const GROUPS = [
  { Ref_Key: G(1), Description: "Телефоны", Parent_Key: EMPTY },
  { Ref_Key: G(2), Description: "Смартфоны", Parent_Key: G(1) },
  { Ref_Key: G(3), Description: "01. Apple", Parent_Key: G(2) },
  { Ref_Key: G(4), Description: "Аксессуары", Parent_Key: EMPTY },
  { Ref_Key: G(5), Description: "01. Apple", Parent_Key: G(4) },
  { Ref_Key: G(6), Description: "Распродажа", Parent_Key: EMPTY },
  { Ref_Key: G(7), Description: "Смартфоны", Parent_Key: G(6) },
  { Ref_Key: G(8), Description: "01. Apple", Parent_Key: G(7) },
  { Ref_Key: G(9), Description: "Старые модели", Parent_Key: G(3) },
].map((g) => ({ ...g, IsFolder: true }));

const ITEMS = [
  { Ref_Key: I(1), Description: "Смартфон Тест 11 64GB Black", Артикул: "T11-64", Parent_Key: G(3) },
  { Ref_Key: I(2), Description: "Смартфон Тест 12", Артикул: "T12", Parent_Key: G(3) },
  { Ref_Key: I(3), Description: "Чехол Тест", Артикул: "CASE-1", Parent_Key: G(5) },
  { Ref_Key: I(4), Description: "Смартфон Тест 10 уценка", Артикул: "T10-U", Parent_Key: G(8) },
  { Ref_Key: I(5), Description: "Смартфон Тест 8", Артикул: "T8", Parent_Key: G(9) },
].map((i) => ({ ...i, IsFolder: false }));

// У магазина три склада одного типа; офисный склад — другого; W(9) — группа складов.
const WAREHOUSES = [
  { Ref_Key: W(1), Description: "Магазин 1_основной", ТипСклада: "РозничныйМагазин", IsFolder: false },
  { Ref_Key: W(2), Description: "Магазин 1_подменный фонд", ТипСклада: "РозничныйМагазин", IsFolder: false },
  {
    Ref_Key: W(3),
    Description: "Магазин 1_ремонт оборудования",
    ТипСклада: "РозничныйМагазин",
    IsFolder: false,
  },
  { Ref_Key: W(4), Description: "Офис основной", ТипСклада: "ОптовыйСклад", IsFolder: false },
  { Ref_Key: W(9), Description: "Магазины", ТипСклада: "", IsFolder: true },
];

const bal = (item: string, wh: string, inStock: number, toShip = 0, ch = EMPTY, series = EMPTY) => ({
  Номенклатура_Key: item,
  Характеристика_Key: ch,
  Склад_Key: wh,
  Назначение_Key: EMPTY,
  Помещение_Key: EMPTY,
  Серия_Key: series,
  КонтролируемаяПартия_Key: EMPTY,
  ВНаличииBalance: inStock,
  КОтгрузкеBalance: toShip,
});

// Ручной расчёт рядом с данными: по каждой строке видно, во что она входит.
const BALANCE = [
  bal(I(1), W(1), 3, 1, EMPTY, C(91)), // серия 1 на основном
  bal(I(1), W(1), 2, 0, EMPTY, C(92)), // серия 2 на основном: вместе 5 / 1
  bal(I(1), W(2), 4), // подменный фонд
  bal(I(1), W(3), 1), // ремонт
  bal(I(2), W(1), 2, 0, C(1)), // характеристика 1
  bal(I(2), W(1), 6, 2, C(2)), // характеристика 2
  bal(I(2), W(4), 10), // офис, без характеристики
  bal(I(3), W(1), 7), // чехол в другой «01. Apple»
  bal(I(4), W(1), 1), // распродажа
  bal(I(5), W(4), 2), // подгруппа «Старые модели»
  bal(I(5), W(1), 0, 0), // нулевая строка не попадает в ответ
];

const guidsIn = (text: string) =>
  [...text.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());

/** Ответы «базы»: остатки с отбором Condition по складам и справочники по $filter. */
function data(balance: typeof BALANCE, balanceOnDate?: typeof BALANCE) {
  return (q: FakeQuery) => {
    const filter = q.params.get("$filter") ?? "";
    if (q.fn === "Balance") {
      const rows = q.args["Period"] && balanceOnDate ? balanceOnDate : balance;
      const condition = q.args["Condition"];
      if (!condition) return rows;
      const allowed = new Set(guidsIn(condition));
      return rows.filter((r) => allowed.has(r.Склад_Key.toLowerCase()));
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

interface StockItem {
  item: string;
  ref: string;
  article?: string;
  group?: string;
  characteristicRef?: string;
  warehouse?: string;
  warehouseRef?: string;
  inStock: number;
  toShip: number;
}
interface Stock {
  totalInStock: number;
  totalToShip: number;
  count: number;
  items: StockItem[];
  warehouseFilter?: { ref: string; name: string; type?: string }[];
  byWarehouse?: {
    warehouse: string | null;
    warehouseRef?: string;
    type?: string;
    inStock: number;
    toShip: number;
  }[];
  itemGroup?: { name: string; ref: string };
  note?: string;
  nextOffset?: number;
  truncated?: boolean;
}

async function setup({
  schema = "ut11",
  balance = BALANCE,
  balanceOnDate,
}: { schema?: FakeSchema; balance?: typeof BALANCE; balanceOnDate?: typeof BALANCE } = {}) {
  const fake = fake1C({ READ_ONLY: "true" }, { schema, data: data(balance, balanceOnDate) });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  // Настоящий клиент: ответ проходит outputSchema.
  const call = async (args: object = {}, name = "read.analytics.get_stock") => {
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Stock;
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake };
}

const qtyOf = (s: Stock, ref: string, extra: Partial<StockItem> = {}) =>
  s.items
    .filter((i) => i.ref === ref && Object.entries(extra).every(([k, v]) => i[k as keyof StockItem] === v))
    .reduce((t, i) => t + i.inStock, 0);

describe("get_stock: итоги совпадают с ручным расчётом", () => {
  it("все склады: серии сворачиваются, характеристики раздельно, нулевые строки не показаны", async () => {
    const { call, fake } = await setup();
    const r = await call();
    expect(r.isError, r.text).toBeFalsy();
    const s = r.structuredContent;
    // 3+2+4+1 + 2+6+10 + 7 + 1 + 2 = 38; к отгрузке 1 + 2 = 3.
    expect(s.totalInStock).toBe(38);
    expect(s.totalToShip).toBe(3);
    expect(qtyOf(s, I(1))).toBe(10);
    expect(s.items.find((i) => i.ref === I(1))).toMatchObject({
      inStock: 10,
      toShip: 1,
      article: "T11-64",
      group: "01. Apple",
    });
    expect(qtyOf(s, I(2), { characteristicRef: C(1) })).toBe(2);
    expect(qtyOf(s, I(2), { characteristicRef: C(2) })).toBe(6);
    // Без характеристики — отдельная строка без поля characteristicRef.
    expect(s.items.filter((i) => i.ref === I(2) && !i.characteristicRef).map((i) => i.inStock)).toEqual([10]);
    // I(1); I(2)×3 по характеристикам; I(3); I(4); I(5) — нулевая строка I(5) на W(1) не добавила позицию.
    expect(s.count).toBe(7);
    expect(s.note).toContain("характеристик не опубликован");
    expect(fake.violations).toEqual([]);
  });

  it("набор складов: исключённые склады не попадают в итог, склады перечислены с типом", async () => {
    const { call, fake } = await setup();
    const r = await call({ warehouses: [W(1), W(4)], groupBy: "warehouse" });
    expect(r.isError, r.text).toBeFalsy();
    const s = r.structuredContent;
    // Основной магазина: 5 + 2 + 6 + 7 + 1 = 21; офис: 10 + 2 = 12. Подменный фонд (4) и ремонт (1) не входят.
    expect(s.totalInStock).toBe(33);
    expect(s.items.some((i) => i.warehouseRef === W(2) || i.warehouseRef === W(3))).toBe(false);
    expect(s.warehouseFilter).toEqual([
      { ref: W(1), name: "Магазин 1_основной", type: "РозничныйМагазин" },
      { ref: W(4), name: "Офис основной", type: "ОптовыйСклад" },
    ]);
    expect(s.byWarehouse).toEqual([
      {
        warehouse: "Магазин 1_основной",
        warehouseRef: W(1),
        type: "РозничныйМагазин",
        positions: 5,
        inStock: 21,
        toShip: 3,
      },
      {
        warehouse: "Офис основной",
        warehouseRef: W(4),
        type: "ОптовыйСклад",
        positions: 2,
        inStock: 12,
        toShip: 0,
      },
    ]);
    // Отбор — на стороне 1С, в Condition только выбранные склады.
    const conditions = fake.requests.filter((q) => q.resource === "AccumulationRegister_ТоварыНаСкладах");
    expect(conditions).toHaveLength(1);
    expect(fake.violations).toEqual([]);
  });

  it("больше 10 складов — несколько запросов, строки не задваиваются", async () => {
    const many = Array.from({ length: 11 }, (_, i) => W(100 + i));
    const extra = many.map((w) => ({
      Ref_Key: w,
      Description: `Склад ${w.slice(-3)}`,
      ТипСклада: "ОптовыйСклад",
      IsFolder: false,
    }));
    WAREHOUSES.push(...extra);
    try {
      const balance = [...BALANCE, ...many.map((w) => bal(I(1), w, 1))];
      const { call, fake } = await setup({ balance });
      const s = (await call({ warehouses: [W(1), ...many] })).structuredContent;
      // Основной магазина 21 + по 1 на каждом из 11 складов.
      expect(s.totalInStock).toBe(32);
      expect(fake.requests.filter((q) => q.resource === "AccumulationRegister_ТоварыНаСкладах")).toHaveLength(
        2,
      );
    } finally {
      WAREHOUSES.splice(WAREHOUSES.length - extra.length, extra.length);
    }
  });

  it("повтор склада и Ref_Key в фигурных скобках — тот же склад, в warehouseFilter один раз", async () => {
    const { call } = await setup();
    const s = (await call({ warehouses: [W(4), W(4).toUpperCase(), `{${W(4)}}`] })).structuredContent;
    expect(s.totalInStock).toBe(12);
    expect(s.warehouseFilter).toHaveLength(1);
  });

  it("неизвестный склад и группа складов — отказ, а не тихий ноль", async () => {
    const { call } = await setup();
    const unknown = await call({ warehouses: [W(1), W(77)] });
    expect(unknown.isError).toBe(true);
    expect(unknown.text).toContain(W(77));
    const folder = await call({ warehouses: [W(9)] });
    expect(folder.isError).toBe(true);
    expect(folder.text).toContain("группы складов");
  });

  it("asOf уходит в Period следующего дня, остатки берутся на эту дату", async () => {
    const { call } = await setup({ balanceOnDate: [bal(I(1), W(1), 42)] });
    const s = (await call({ asOf: "2026-10-31" })).structuredContent;
    expect(s.totalInStock).toBe(42);
    const now = (await call()).structuredContent;
    expect(now.totalInStock).toBe(38);
  });
});

describe("get_stock: группа номенклатуры", () => {
  it("полный путь выбирает одну из одноимённых групп, подгруппы входят", async () => {
    const { call } = await setup();
    const s = (await call({ group: "Телефоны/Смартфоны/01. Apple" })).structuredContent;
    // I(1) 10 + I(2) 18 + I(5) 2 из подгруппы «Старые модели»; чехол и распродажа не входят.
    expect(s.totalInStock).toBe(30);
    expect(s.itemGroup).toMatchObject({ name: "01. Apple", ref: G(3) });
  });

  it("одно имя на несколько групп — отказ со списком путей", async () => {
    const { call } = await setup();
    const r = await call({ group: "01. Apple" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Телефоны/Смартфоны/01. Apple");
    expect(r.text).toContain("Аксессуары/01. Apple");
    expect(r.text).toContain("Распродажа/Смартфоны/01. Apple");
    const partial = await call({ group: "Смартфоны/01. Apple" });
    expect(partial.isError).toBe(true);
    expect(partial.text).not.toContain("Аксессуары");
  });

  it("Ref_Key группы и набор складов вместе", async () => {
    const { call } = await setup();
    const s = (await call({ group: G(3), warehouses: [W(1)] })).structuredContent;
    // На основном магазина из G(3) с подгруппами: I(1) 5 + I(2) 8.
    expect(s.totalInStock).toBe(13);
  });

  it("otherSimilar.count — позиции, а не строки складов", async () => {
    // «Смартфон» латинской «C» в одной позиции на двух складах; запрос кириллицей находит буквальные совпадения.
    const lookalike = {
      Ref_Key: I(60),
      Description: "Cмартфон двойник",
      Артикул: "X",
      Parent_Key: G(3),
      IsFolder: false,
    };
    ITEMS.push(lookalike);
    try {
      const { call } = await setup({ balance: [...BALANCE, bal(I(60), W(1), 1), bal(I(60), W(4), 1)] });
      const s = (await call({ query: "смартфон", groupBy: "warehouse" })).structuredContent as Stock & {
        otherSimilar?: { count: number; inStock: number };
      };
      expect(s.otherSimilar).toMatchObject({ count: 1, inStock: 2 });
    } finally {
      ITEMS.pop();
    }
  });

  it("query по артикулу внутри группы", async () => {
    const { call } = await setup();
    const s = (await call({ group: G(3), query: "t12" })).structuredContent;
    expect(s.totalInStock).toBe(18);
    expect(new Set(s.items.map((i) => i.ref))).toEqual(new Set([I(2)]));
  });
});

describe("get_stock: большая ветка групп", () => {
  it("больше 30 папок — справочник читается целиком, а не пачками по 10; итог тот же", async () => {
    const extraGroups = Array.from({ length: 40 }, (_, i) => ({
      Ref_Key: G(200 + i),
      Description: `Серия ${i}`,
      Parent_Key: G(3),
      IsFolder: true,
    }));
    const extraItems = extraGroups.map((g, i) => ({
      Ref_Key: I(200 + i),
      Description: `Смартфон серии ${i}`,
      Артикул: `S-${i}`,
      Parent_Key: g.Ref_Key,
      IsFolder: false,
    }));
    GROUPS.push(...extraGroups);
    ITEMS.push(...extraItems);
    try {
      const balance = [...BALANCE, ...extraItems.map((it) => bal(it.Ref_Key, W(4), 1))];
      const { call, fake } = await setup({ balance });
      const s = (await call({ group: "Телефоны/Смартфоны/01. Apple" })).structuredContent;
      // Ветка G(3) без новых папок — 30, плюс по 1 у каждой из 40 позиций в подпапках.
      expect(s.totalInStock).toBe(70);
      const catalogReads = fake.requests.filter((q) => q.resource === "Catalog_Номенклатура");
      // Группы, вся номенклатура и названия групп страницы — а не 5 пачек по 10 папок.
      expect(catalogReads.length).toBeLessThanOrEqual(3);
    } finally {
      GROUPS.splice(GROUPS.length - extraGroups.length, extraGroups.length);
      ITEMS.splice(ITEMS.length - extraItems.length, extraItems.length);
    }
  });
});

describe("get_stock: каждый параметр меняет результат", () => {
  it("warehouses, group, query, groupBy, asOf", async () => {
    const { call } = await setup({ balanceOnDate: [bal(I(1), W(1), 1)] });
    const base = (await call()).structuredContent;
    const variants = [
      { warehouses: [W(4)] },
      { group: "Аксессуары" },
      { query: "уценка" },
      { groupBy: "warehouse" },
      { asOf: "2026-01-01" },
    ];
    for (const args of variants) {
      const s = (await call(args)).structuredContent;
      expect(
        JSON.stringify([s.totalInStock, s.count, s.items.map((i) => [i.ref, i.warehouseRef])]),
        JSON.stringify(args),
      ).not.toBe(
        JSON.stringify([base.totalInStock, base.count, base.items.map((i) => [i.ref, i.warehouseRef])]),
      );
    }
  });
});

describe("get_stock: страницы", () => {
  it("страницы с offset дают все позиции ровно по разу, итоги — по всему отбору", async () => {
    const { call } = await setup();
    const seen: string[] = [];
    let offset = 0;
    for (let guard = 0; guard < 10; guard++) {
      const s = (await call({ limit: 3, offset })).structuredContent;
      expect(s.totalInStock).toBe(38);
      expect(s.count).toBe(7);
      seen.push(...s.items.map((i) => `${i.ref}|${i.characteristicRef ?? ""}`));
      if (s.nextOffset === undefined) break;
      offset = s.nextOffset;
    }
    expect(seen).toHaveLength(7);
    expect(new Set(seen).size).toBe(7);
  });
});

describe("неопубликованные регистры: понятный отказ", () => {
  it("get_stock в Бухгалтерии — не опубликован регистр «Товары на складах»", async () => {
    const { call } = await setup({ schema: "bp30" });
    const r = await call();
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Товары на складах");
    expect((r as unknown as { _meta?: Record<string, string> })._meta?.["odata-1c/errorKind"]).toBe(
      "not_published",
    );
  });

  it("get_inventory в УТ — отказ ведёт к get_stock, а не к публикации Хозрасчётного", async () => {
    const { call } = await setup();
    const r = await call({}, "read.analytics.get_inventory");
    expect(r.isError).toBe(true);
    expect(r.text).toContain("read.analytics.get_stock");
    expect(r.text).not.toContain("Добавьте их в 1С");
  });
});

describe("resolveItemGroup: пути", () => {
  const groups: ItemGroup[] = GROUPS.map((g) => ({
    ref: g.Ref_Key,
    name: g.Description,
    ...(g.Parent_Key !== EMPTY ? { parent: g.Parent_Key } : {}),
  }));

  it("путь — суффикс цепочки, части целиком, регистр и ё не важны", () => {
    expect(resolveItemGroup("телефоны / смартфоны / 01. apple", groups).ref).toBe(G(3));
    expect(resolveItemGroup("Аксессуары/01. Apple", groups).ref).toBe(G(5));
    expect(() => resolveItemGroup("Смарт/01. Apple", groups)).toThrow(/не найдена/);
  });

  it("путь из текста отказа находит группу, даже если в названии родителя есть «/» (ревью)", () => {
    const withSlash = [
      ...groups,
      { ref: G(70), name: "Кабели/переходники" },
      { ref: G(71), name: "USB", parent: G(70) },
      { ref: G(72), name: "USB", parent: G(4) },
    ];
    let message = "";
    try {
      resolveItemGroup("USB", withSlash);
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain("Кабели/переходники/USB");
    expect(resolveItemGroup("Кабели/переходники/USB", withSlash).ref).toBe(G(71));
    expect(resolveItemGroup("Кабели / переходники / USB", withSlash).ref).toBe(G(71));
  });

  it("название с «/» целиком важнее пути", () => {
    const withSlash = [
      ...groups,
      { ref: G(50), name: "Кабели/переходники" },
      { ref: G(51), name: "Кабели" },
      { ref: G(52), name: "переходники", parent: G(51) },
    ];
    expect(resolveItemGroup("Кабели/переходники", withSlash).ref).toBe(G(50));
  });

  it("путь к группе строится от корня, цикл в данных не зацикливает", () => {
    expect(
      groupPath(
        groups.find((g) => g.ref === G(9))!,
        groups,
      ),
    ).toBe("Телефоны/Смартфоны/01. Apple/Старые модели");
    const loop: ItemGroup[] = [
      { ref: G(60), name: "А", parent: G(61) },
      { ref: G(61), name: "Б", parent: G(60) },
    ];
    expect(groupPath(loop[0]!, loop)).toBe("Б/А");
  });
});
