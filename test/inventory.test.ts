import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { dimensionOf, resolveItemGroup, subtreeOf } from "../src/odata/inventory.js";
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
  vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
    const u = decodeURIComponent(url);
    if (u.includes("/Balance")) return { value: balance };
    if (/^ChartOfAccounts/.test(u)) return realGet(url);
    if (u.startsWith("Catalog_Номенклатура") && u.includes("IsFolder eq true")) {
      return {
        value: Object.entries(CATALOG)
          .filter(([, v]) => v.IsFolder)
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
  return { call };
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
