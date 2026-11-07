import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import {
  aggregateStock,
  dimensionOf,
  foldText,
  matchesItemQuery,
  strictText,
  resolveItemGroup,
  subtreeOf,
} from "../src/odata/inventory.js";
import { fake1C } from "./helpers/fake-1c.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const ITEM = (i: number) => uuid("a", i);
const STORE = (i: number) => uuid("b", i);
const CP = uuid("c", 1);
const GROUP = (i: number) => uuid("d", i);
const NOM = "StandardODATA.Catalog_Номенклатура";
const WH = "StandardODATA.Catalog_Склады";
const CPT = "StandardODATA.Catalog_Контрагенты";
const PARTY = "StandardODATA.Document_ПоступлениеТоваровУслуг";

type Dim = [string, string];
const row = (dims: Dim[], qty: number, sum: number) => ({
  Account_Key: "x",
  ...Object.fromEntries(
    dims.flatMap(([ref, type], i) => [
      [`ExtDimension${i + 1}`, ref],
      [`ExtDimension${i + 1}_Type`, type],
    ]),
  ),
  КоличествоBalanceDr: qty,
  КоличествоBalanceCr: 0,
  СуммаBalanceDr: sum,
  СуммаBalanceCr: 0,
});

const CATALOG: Record<
  string,
  { Description: string; Артикул?: string; Parent_Key?: string; IsFolder?: boolean }
> = {
  [ITEM(1)]: { Description: "Бумага", Артикул: "A-1", Parent_Key: GROUP(2) },
  [ITEM(2)]: { Description: "Ручка", Артикул: "  ", Parent_Key: GROUP(1) },
  [ITEM(3)]: { Description: "Стол", Parent_Key: GROUP(3) },
  [GROUP(1)]: { Description: "Канцелярия", IsFolder: true },
  [GROUP(2)]: { Description: "Бумажное", Parent_Key: GROUP(1), IsFolder: true },
  [GROUP(3)]: { Description: "Мебель", IsFolder: true },
  [STORE(1)]: { Description: "Основной" },
  [STORE(2)]: { Description: "Розница" },
};

async function setup(balance: object[]) {
  const f = fake1C({ READ_ONLY: "true" });
  const realGet = f.conn.client.getCollection.bind(f.conn.client);
  const seenReads: string[] = [];
  vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
    const u = decodeURIComponent(url);
    seenReads.push(u);
    if (u.includes("/Balance")) return { value: balance };
    if (/^ChartOfAccounts/.test(u)) return realGet(url);
    if (
      u.startsWith("Catalog_Номенклатура") &&
      u.includes("IsFolder eq false") &&
      !u.includes("Parent_Key eq")
    ) {
      return {
        value: Object.entries(CATALOG)
          .filter(([ref, v]) => !v.IsFolder && ref.startsWith("a"))
          .map(([Ref_Key, v]) => ({ Ref_Key, ...v })),
      };
    }
    if (u.startsWith("Catalog_Номенклатура") && u.includes("IsFolder eq true")) {
      return {
        value: Object.entries(CATALOG)
          .filter(([, v]) => v.IsFolder)
          .map(([Ref_Key, v]) => ({ Ref_Key, ...v })),
      };
    }
    if (u.startsWith("Catalog_Номенклатура") && u.includes("Parent_Key eq")) {
      const groups = [...u.matchAll(/Parent_Key eq guid'([^']+)'/g)].map((m) => m[1]!);
      return {
        value: Object.entries(CATALOG)
          .filter(([, v]) => !v.IsFolder && v.Parent_Key && groups.includes(v.Parent_Key))
          .map(([Ref_Key, v]) => ({ Ref_Key, ...v })),
      };
    }
    if (u.startsWith("Catalog_Номенклатура") || u.startsWith("Catalog_Склады")) {
      const refs = [...u.matchAll(/guid'([^']+)'/g)].map((m) => m[1]!);
      return { value: refs.filter((r) => CATALOG[r]).map((Ref_Key) => ({ Ref_Key, ...CATALOG[Ref_Key] })) };
    }
    return realGet(url);
  }) as never);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(f.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  // Настоящий клиент: ответ проходит outputSchema.
  const call = async (args: object = {}) => {
    const r = (await client.callTool({ name: "read.analytics.get_inventory", arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, never>;
      content: { text: string }[];
    };
    return r;
  };
  return { call, reads: () => seenReads };
}

interface Item {
  item: string;
  ref: string;
  article?: string;
  group?: string;
  groupRef?: string;
  warehouse?: string;
  quantity: number;
  amount: number;
}

describe("субконто по типу, а не по позиции", () => {
  it("dimensionOf находит значение на любой позиции", () => {
    expect(dimensionOf(row([[ITEM(1), NOM]], 1, 1), ["Catalog_Номенклатура"])?.ref).toBe(ITEM(1));
    expect(
      dimensionOf(
        row(
          [
            [CP, CPT],
            [ITEM(1), NOM],
          ],
          1,
          1,
        ),
        ["Catalog_Номенклатура"],
      )?.ref,
    ).toBe(ITEM(1));
    expect(dimensionOf(row([[STORE(1), WH]], 1, 1), ["Catalog_Номенклатура"])).toBeUndefined();
    expect(
      dimensionOf(
        row(
          [
            [ITEM(1), NOM],
            [STORE(1), WH],
          ],
          1,
          1,
        ),
        ["Catalog_Склады"],
      )?.ref,
    ).toBe(STORE(1));
  });

  it("контрагент на 10.07 и склад на 41.12 не становятся товаром, итог прежний", async () => {
    const { call } = await setup([
      row(
        [
          [ITEM(1), NOM],
          [PARTY, PARTY],
        ],
        2,
        100,
      ), // 41.01
      row(
        [
          [CP, CPT],
          [ITEM(1), NOM],
          [PARTY, PARTY],
        ],
        3,
        50,
      ), // 10.07
      row([[STORE(2), WH]], 0, 70), // 41.12
    ]);
    const s = (await call()).structuredContent as unknown as {
      items: Item[];
      totalAmount: number;
      count: number;
      withoutNomenclature: { amount: number };
      note: string;
    };
    expect(s.items).toHaveLength(1);
    expect(s.items[0]).toMatchObject({ item: "Бумага", quantity: 5, amount: 150 });
    expect(s.count).toBe(1);
    expect(s.withoutNomenclature.amount).toBe(70);
    expect(s.totalAmount).toBe(220);
    expect(s.note).toContain("withoutNomenclature");
  });
});

describe("артикул и группа", () => {
  it("есть в ответе, пустой артикул не выводится, ответа по складам нет", async () => {
    const { call } = await setup([row([[ITEM(1), NOM]], 1, 10), row([[ITEM(2), NOM]], 1, 5)]);
    const s = (await call()).structuredContent as unknown as {
      items: Item[];
      byWarehouse?: unknown;
      groupBy?: string;
    };
    expect(s.items[0]).toMatchObject({
      item: "Бумага",
      article: "A-1",
      group: "Бумажное",
      groupRef: GROUP(2),
    });
    expect(s.items[1]).toMatchObject({ item: "Ручка", group: "Канцелярия" });
    expect(s.items[1]).not.toHaveProperty("article");
    expect(s.byWarehouse).toBeUndefined();
    expect(s.groupBy).toBeUndefined();
  });
});

describe("разбивка по складам", () => {
  const rows = () => [
    row(
      [
        [ITEM(1), NOM],
        [STORE(1), WH],
        [PARTY, PARTY],
      ],
      2,
      100,
    ), // склад на позиции 2
    row(
      [
        [ITEM(1), NOM],
        [PARTY, PARTY],
        [STORE(2), WH],
      ],
      1,
      40,
    ), // склад на позиции 3
    row([[STORE(2), WH]], 0, 70), // склад на позиции 1, номенклатуры нет
    row(
      [
        [ITEM(2), NOM],
        [PARTY, PARTY],
      ],
      1,
      5,
    ), // без склада
  ];

  it("позиция склада не важна; суммы сходятся с итогом", async () => {
    const { call } = await setup(rows());
    const s = (await call({ groupBy: "warehouse" })).structuredContent as unknown as {
      items: Item[];
      totalAmount: number;
      count: number;
      warehouseTracked: boolean;
      byWarehouse: { warehouse: string | null; positions: number; amount: number }[];
      withoutNomenclature: { amount: number };
      note: string;
    };
    expect(s.warehouseTracked).toBe(true);
    expect(s.items.map((i) => [i.item, i.warehouse, i.amount])).toEqual([
      ["Бумага", "Основной", 100],
      ["Бумага", "Розница", 40],
      ["Ручка", undefined, 5],
    ]);
    expect(s.count).toBe(3);
    expect(s.byWarehouse).toEqual([
      expect.objectContaining({ warehouse: "Розница", positions: 1, amount: 110 }),
      expect.objectContaining({ warehouse: "Основной", positions: 1, amount: 100 }),
      expect.objectContaining({ warehouse: null, positions: 1, amount: 5 }),
    ]);
    expect(s.totalAmount).toBe(215);
    expect(s.byWarehouse.reduce((t, w) => t + w.amount, 0)).toBe(s.totalAmount);
    expect(s.note).toContain("без склада");
  });

  it("без разбивки те же строки сводятся по номенклатуре", async () => {
    const { call } = await setup(rows());
    const s = (await call()).structuredContent as unknown as { items: Item[]; totalAmount: number };
    expect(s.items.map((i) => [i.item, i.quantity, i.amount])).toEqual([
      ["Бумага", 3, 140],
      ["Ручка", 1, 5],
    ]);
    expect(s.totalAmount).toBe(215);
  });

  it("склад не ведётся: прямо сказано, разбивки нет", async () => {
    const { call } = await setup([
      row(
        [
          [ITEM(1), NOM],
          [PARTY, PARTY],
        ],
        2,
        100,
      ),
    ]);
    const s = (await call({ groupBy: "warehouse" })).structuredContent as unknown as {
      warehouseTracked: boolean;
      byWarehouse?: unknown;
      note: string;
      items: Item[];
    };
    expect(s.warehouseTracked).toBe(false);
    expect(s.byWarehouse).toBeUndefined();
    expect(s.note).toContain("склада нет");
    expect(s.items[0]!.amount).toBe(100);
  });

  it("чтение по offset: все строки разбивки без потерь при равных суммах", async () => {
    const many = Array.from({ length: 30 }, (_, i) =>
      row(
        [
          [ITEM(100 + i), NOM],
          [STORE(1), WH],
        ],
        1,
        10,
      ),
    );
    const { call } = await setup(many);
    const seen: string[] = [];
    let offset = 0;
    for (let i = 0; i < 20; i++) {
      const s = (await call({ groupBy: "warehouse", limit: 7, offset })).structuredContent as unknown as {
        items: Item[];
        nextOffset?: number;
        count: number;
      };
      expect(s.count).toBe(30);
      seen.push(...s.items.map((x) => x.ref));
      if (s.nextOffset === undefined) break;
      offset = s.nextOffset;
    }
    expect(new Set(seen).size).toBe(30);
    expect(seen).toHaveLength(30);
  });
});

describe("отбор по группе", () => {
  const balance = () => [
    row([[ITEM(1), NOM]], 1, 10),
    row([[ITEM(2), NOM]], 1, 20),
    row([[ITEM(3), NOM]], 1, 40),
    row([[STORE(2), WH]], 0, 70),
  ];
  it("с подгруппами; итог только по группе; строки без номенклатуры не входят", async () => {
    const { call } = await setup(balance());
    const s = (await call({ group: "Канцелярия" })).structuredContent as unknown as {
      items: Item[];
      totalAmount: number;
      itemGroup: { name: string; ref: string };
      withoutNomenclature?: unknown;
    };
    expect(s.items.map((i) => i.item).sort()).toEqual(["Бумага", "Ручка"]);
    expect(s.totalAmount).toBe(30);
    expect(s.itemGroup).toMatchObject({ name: "Канцелярия", ref: GROUP(1) });
    expect(s.withoutNomenclature).toBeUndefined();
  });
  it("подгруппа и Ref_Key дают другой результат", async () => {
    const { call } = await setup(balance());
    const sub = (await call({ group: GROUP(2) })).structuredContent as unknown as { items: Item[] };
    expect(sub.items.map((i) => i.item)).toEqual(["Бумага"]);
    const other = (await call({ group: "мебель" })).structuredContent as unknown as { items: Item[] };
    expect(other.items.map((i) => i.item)).toEqual(["Стол"]);
  });
  it("нет такой группы — отказ, а не пустой остаток", async () => {
    const { call } = await setup(balance());
    const r = await call({ group: "Нет такой" });
    expect(r.isError).toBe(true);
    expect(r.content[0]!.text).toContain("не найдена");
  });
  it("неоднозначное название — отказ со списком", () => {
    const groups = [
      { ref: GROUP(1), name: "Бумага офисная" },
      { ref: GROUP(2), name: "Бумага цветная" },
    ];
    expect(() => resolveItemGroup("Бумага", groups)).toThrow(/несколько групп/);
    expect(resolveItemGroup("бумага цветная", groups).ref).toBe(GROUP(2));
  });
  it("цикл родителей не зацикливает обход", () => {
    const groups = [
      { ref: "a", name: "a", parent: "b" },
      { ref: "b", name: "b", parent: "a" },
    ];
    expect([...subtreeOf("a", groups)].sort()).toEqual(["a", "b"]);
  });
});

describe("конфигурации без типов субконто", () => {
  it("номенклатура — первая позиция, как раньше", () => {
    const lines = aggregateStock(
      [{ ExtDimension1: ITEM(1), КоличествоBalanceDr: 2, СуммаBalanceDr: 100 }],
      false,
    );
    expect(lines).toEqual([expect.objectContaining({ itemRef: ITEM(1), cents: 10_000 })]);
  });
  it("строка с типами, но без номенклатуры, остаётся без неё", () => {
    expect(aggregateStock([row([[STORE(1), WH]], 0, 10)], false)[0]!.itemRef).toBeUndefined();
  });
});

describe("признак учёта по складам", () => {
  it("нет остатков — признака нет: неизвестно, ведётся ли склад", async () => {
    const { call } = await setup([]);
    const s = (await call({ groupBy: "warehouse" })).structuredContent as unknown as Record<string, unknown>;
    expect(s).not.toHaveProperty("warehouseTracked");
    expect(s).not.toHaveProperty("note");
  });
  it("отбор по группе без остатков не отменяет ведение склада", async () => {
    const { call } = await setup([
      row(
        [
          [ITEM(3), NOM],
          [STORE(1), WH],
        ],
        1,
        10,
      ),
    ]);
    const s = (await call({ groupBy: "warehouse", group: "Канцелярия" })).structuredContent as unknown as {
      warehouseTracked: boolean;
      items: unknown[];
    };
    expect(s.warehouseTracked).toBe(true);
    expect(s.items).toEqual([]);
  });
  it("byWarehouse только на первой странице", async () => {
    const many = Array.from({ length: 12 }, (_, i) =>
      row(
        [
          [ITEM(100 + i), NOM],
          [STORE(1), WH],
        ],
        1,
        10,
      ),
    );
    const { call } = await setup(many);
    const first = (await call({ groupBy: "warehouse", limit: 5 })).structuredContent as unknown as Record<
      string,
      unknown
    >;
    const next = (await call({ groupBy: "warehouse", limit: 5, offset: 5 }))
      .structuredContent as unknown as Record<string, unknown>;
    expect(first).toHaveProperty("byWarehouse");
    expect(next).not.toHaveProperty("byWarehouse");
    expect(next["warehouseTracked"]).toBe(true);
  });
  it("пробелы в group — отказ проверки аргументов", async () => {
    const { call } = await setup([]);
    const r = await call({ group: "   " }).catch((e: Error) => ({
      isError: true,
      content: [{ text: e.message }],
    }));
    expect(r.isError).toBe(true);
  });
  it("отбор по группе не читает позиции по одной", async () => {
    const many = Array.from({ length: 200 }, (_, i) => row([[ITEM(1000 + i), NOM]], 1, 1));
    const { call, reads } = await setup(many);
    await call({ group: "Канцелярия", limit: 5 });
    expect(reads().filter((u) => u.startsWith("Catalog_Номенклатура")).length).toBeLessThanOrEqual(3);
  });
});

describe("разбивка без типов субконто и на следующих страницах", () => {
  it("типов нет: не утверждает, что склад не ведётся", async () => {
    const { call } = await setup([
      { ExtDimension1: ITEM(1), ExtDimension2: STORE(1), КоличествоBalanceDr: 1, СуммаBalanceDr: 10 },
    ]);
    const s = (await call({ groupBy: "warehouse" })).structuredContent as unknown as {
      note: string;
      items: Item[];
    } & Record<string, unknown>;
    expect(s).not.toHaveProperty("warehouseTracked");
    expect(s.note).toContain("определить нельзя");
    expect(s.note).not.toContain("ведутся без разбивки");
    expect(s.items[0]!.item).toBe("Бумага");
  });
  it("примечание про byWarehouse только там, где он есть", async () => {
    const many = [
      ...Array.from({ length: 8 }, (_, i) =>
        row(
          [
            [ITEM(100 + i), NOM],
            [STORE(1), WH],
          ],
          1,
          10,
        ),
      ),
      row([[ITEM(200), NOM]], 1, 5),
    ];
    const { call } = await setup(many);
    const first = (await call({ groupBy: "warehouse", limit: 4 })).structuredContent as unknown as {
      note: string;
    };
    const next = (await call({ groupBy: "warehouse", limit: 4, offset: 4 })).structuredContent as unknown as {
      note?: string;
    };
    expect(first.note).toContain("byWarehouse");
    expect(next.note ?? "").not.toContain("byWarehouse");
  });
});

describe("поиск по товару", () => {
  const balance = () => [
    row([[ITEM(1), NOM]], 1, 10),
    row([[ITEM(2), NOM]], 2, 20),
    row([[ITEM(3), NOM]], 3, 40),
    row([[STORE(2), WH]], 0, 70),
  ];
  const names = (s: unknown) => (s as { items: Item[] }).items.map((i) => i.item);

  it("по названию и по артикулу дают разный результат, итог только по найденному", async () => {
    const { call } = await setup(balance());
    const byName = (await call({ query: "ручка" })).structuredContent as unknown as {
      items: Item[];
      totalAmount: number;
      count: number;
      query: string;
      withoutNomenclature?: unknown;
    };
    expect(names(byName)).toEqual(["Ручка"]);
    expect(byName.totalAmount).toBe(20);
    expect(byName.count).toBe(1);
    expect(byName.query).toBe("ручка");
    expect(byName.withoutNomenclature).toBeUndefined();
    const byArticle = (await call({ query: "a-1" })).structuredContent;
    expect(names(byArticle)).toEqual(["Бумага"]);
    expect((byArticle as unknown as { items: Item[] }).items[0]).toMatchObject({
      article: "A-1",
      group: "Бумажное",
    });
  });
  it("ничего не найдено — пусто без ошибки, справочник читается одним запросом", async () => {
    const { call, reads } = await setup(balance());
    const r = await call({ query: "нет такого" });
    expect(r.isError).toBeFalsy();
    expect(names(r.structuredContent)).toEqual([]);
    expect((r.structuredContent as unknown as { totalAmount: number }).totalAmount).toBe(0);
    expect(reads().filter((u) => u.startsWith("Catalog_Номенклатура"))).toHaveLength(1);
  });
  it("вместе с группой — пересечение", async () => {
    const { call } = await setup(balance());
    expect(names((await call({ query: "стол", group: "Канцелярия" })).structuredContent)).toEqual([]);
    expect(names((await call({ query: "бумага", group: "Канцелярия" })).structuredContent)).toEqual([
      "Бумага",
    ]);
  });
  it("вместе с разбивкой по складам", async () => {
    const { call } = await setup([
      row(
        [
          [ITEM(1), NOM],
          [STORE(1), WH],
        ],
        1,
        10,
      ),
      row(
        [
          [ITEM(2), NOM],
          [STORE(2), WH],
        ],
        1,
        5,
      ),
    ]);
    const s = (await call({ query: "бумага", groupBy: "warehouse" })).structuredContent as unknown as {
      items: Item[];
      byWarehouse: unknown[];
    };
    expect(s.items.map((i) => i.warehouse)).toEqual(["Основной"]);
    expect(s.byWarehouse).toHaveLength(1);
  });
  it("в запросе к 1С нет substringof: скобки не сломают его", async () => {
    const { call, reads } = await setup(balance());
    const r = await call({ query: "бумага (А4)" });
    expect(r.isError).toBeFalsy();
    expect(reads().join(" ")).not.toContain("substringof");
  });
  it("слова в любом порядке, без учёта регистра", () => {
    const item = { name: "Бумага офисная (A4)", article: "БМ-1" };
    expect(matchesItemQuery(item, "a4 БУМАГА")).toBe(true);
    expect(matchesItemQuery(item, "бм-1")).toBe(true);
    expect(matchesItemQuery(item, "(a4)")).toBe(true);
    expect(matchesItemQuery(item, "бумага ручка")).toBe(false);
  });
});

describe("поиск без ложных нулей", () => {
  it("ё как е, похожие буквы разных раскладок", () => {
    expect(matchesItemQuery({ name: "Щётка" }, "щетка")).toBe(true);
    expect(matchesItemQuery({ name: "Елка" }, "ёлка")).toBe(true);
    expect(matchesItemQuery({ name: "Бумага (A4)" }, "бумаг а4")).toBe(true);
    expect(matchesItemQuery({ name: "Бумага (А4)" }, "бумаг a4")).toBe(true);
    expect(matchesItemQuery({ article: "KM-5" }, "км-5")).toBe(true);
    expect(foldText("А4")).toBe(foldText("a4"));
  });
  it("словоформы не распознаются, но ответ объясняет пустой результат", async () => {
    const { call } = await setup([row([[ITEM(1), NOM]], 1, 10)]);
    const s = (await call({ query: "бумаги" })).structuredContent as unknown as {
      items: Item[];
      note: string;
    };
    expect(s.items).toEqual([]);
    expect(s.note).toContain("основу слова");
    const hit = (await call({ query: "бумаг" })).structuredContent as unknown as { items: Item[] };
    expect(hit.items).toHaveLength(1);
  });
  it("слишком длинный запрос — отказ проверки аргументов", async () => {
    const { call } = await setup([]);
    const r = await call({ query: "я".repeat(201) }).catch((e: Error) => ({
      isError: true,
      content: [{ text: e.message }],
    }));
    expect(r.isError).toBe(true);
  });
  it("соседние страницы не перечитывают справочник", async () => {
    const many = Array.from({ length: 12 }, (_, i) =>
      row(
        [
          [ITEM(1), NOM],
          [STORE(100 + i), WH],
        ],
        1,
        10,
      ),
    );
    const { call, reads } = await setup(many);
    await call({ query: "бумаг", groupBy: "warehouse", limit: 5 });
    await call({ query: "бумаг", groupBy: "warehouse", limit: 5, offset: 5 });
    expect(reads().filter((u) => u.includes("IsFolder eq false"))).toHaveLength(1);
  });
});

describe("уровни сопоставления и свежесть справочника", () => {
  it("«TP» не выбирает чужую группу по похожим буквам", () => {
    const groups = [{ ref: GROUP(1), name: "Стройматериалы" }];
    expect(() => resolveItemGroup("TP", groups)).toThrow(/не найдена/);
    expect(resolveItemGroup("строй", groups).ref).toBe(GROUP(1));
  });
  it("название группы целиком находится и в другой раскладке, пробел в конце не мешает", () => {
    expect(resolveItemGroup("A4", [{ ref: GROUP(1), name: "А4" }]).ref).toBe(GROUP(1));
    expect(resolveItemGroup("Мебель", [{ ref: GROUP(1), name: "Мебель " }]).ref).toBe(GROUP(1));
  });
  it("буквальные совпадения важнее похожих букв, вторым шагом примечание", async () => {
    CATALOG[ITEM(50)] = { Description: "Кабель 3м" };
    CATALOG[ITEM(51)] = { Description: "Лента 3M" };
    try {
      const { call } = await setup([row([[ITEM(50), NOM]], 1, 10), row([[ITEM(51), NOM]], 1, 20)]);
      const both = (await call({ query: "3m" })).structuredContent as unknown as {
        items: Item[];
        totalAmount: number;
        otherSimilar: { count: number; amount: number; examples: string[] };
        note: string;
      };
      expect(both.items.map((i) => i.item)).toEqual(["Лента 3M"]);
      expect(both.totalAmount).toBe(20);
      expect(both.otherSimilar).toEqual({ count: 1, amount: 10, examples: ["Кабель 3м"] });
      expect(both.note).toContain("otherSimilar");
      const { call: call2 } = await setup([row([[ITEM(50), NOM]], 1, 10)]);
      const look = (await call2({ query: "3m" })).structuredContent as unknown as {
        items: Item[];
        note: string;
      };
      expect(look.items.map((i) => i.item)).toEqual(["Кабель 3м"]);
      expect(look.note).toContain("похожих букв");
    } finally {
      delete CATALOG[ITEM(50)];
      delete CATALOG[ITEM(51)];
    }
  });
  it("товар есть в справочнике без остатка и товара нет: разные примечания", async () => {
    const { call } = await setup([row([[ITEM(2), NOM]], 1, 10)]);
    const noStock = (await call({ query: "бумаг" })).structuredContent as unknown as { note: string };
    expect(noStock.note).toContain("остатка на счетах нет");
    const unknown = (await call({ query: "бумаги" })).structuredContent as unknown as { note: string };
    expect(unknown.note).toContain("не найдено");
    const inGroup = (await call({ query: "стол", group: "Канцелярия" })).structuredContent as unknown as {
      note: string;
    };
    expect(inGroup.note).toContain("В выбранной группе такого товара нет");
  });
  it("первая страница читает свежий справочник: новый товар виден сразу", async () => {
    const { call, reads } = await setup([row([[ITEM(1), NOM]], 1, 10)]);
    await call({ query: "бумаг" });
    await call({ query: "бумаг" });
    expect(reads().filter((u) => u.includes("IsFolder eq false"))).toHaveLength(2);
  });
  it("strictText: регистр и ё", () => {
    expect(strictText("ЁЛКА")).toBe("елка");
    expect(strictText("A4")).not.toBe(strictText("А4"));
  });
});

describe("группа: название целиком важнее части", () => {
  it("«A4» латиницей выбирает группу «А4», а не «Бумага A4»", () => {
    const groups = [
      { ref: GROUP(1), name: "А4" },
      { ref: GROUP(2), name: "Бумага A4" },
    ];
    expect(resolveItemGroup("A4", groups).ref).toBe(GROUP(1));
  });
});

describe("примечание при пустом поиске", () => {
  it("товар другой раскладкой без остатка: не советует искать основу слова", async () => {
    CATALOG[ITEM(60)] = { Description: "Бумага A4" };
    try {
      const { call } = await setup([row([[ITEM(2), NOM]], 1, 10)]);
      const s = (await call({ query: "бумага а4" })).structuredContent as unknown as { note: string };
      expect(s.note).toContain("другой раскладки");
      expect(s.note).not.toContain("основу слова");
    } finally {
      delete CATALOG[ITEM(60)];
    }
  });
  it("отбор по группе читает только товары, без подгрупп", async () => {
    const { call, reads } = await setup([row([[ITEM(1), NOM]], 1, 10)]);
    await call({ group: "Канцелярия" });
    const query = reads().find((u) => u.includes("Parent_Key eq"))!;
    expect(query).toContain("IsFolder eq false");
  });
});
