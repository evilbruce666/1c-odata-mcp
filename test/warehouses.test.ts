import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/** Список складов на синтетическом справочнике: три склада на магазин, папки, помеченные на удаление. */

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const W = (i: number) => uuid("b", i);
const EMPTY = "00000000-0000-0000-0000-000000000000";

const WAREHOUSES = [
  {
    Ref_Key: W(90),
    Description: "Магазины",
    IsFolder: true,
    Parent_Key: EMPTY,
    DeletionMark: false,
    ТипСклада: "",
  },
  {
    Ref_Key: W(91),
    Description: "Север",
    IsFolder: true,
    Parent_Key: W(90),
    DeletionMark: false,
    ТипСклада: "",
  },
  { Ref_Key: W(1), Description: "Магазин 1_основной", ТипСклада: "РозничныйМагазин", Parent_Key: W(91) },
  {
    Ref_Key: W(2),
    Description: "Магазин 1_подменный фонд",
    ТипСклада: "РозничныйМагазин",
    Parent_Key: W(91),
  },
  {
    Ref_Key: W(3),
    Description: "Магазин 1_ремонт оборудования",
    ТипСклада: "РозничныйМагазин",
    Parent_Key: W(91),
  },
  { Ref_Key: W(4), Description: "Офис основной", ТипСклада: "ОптовыйСклад", Parent_Key: EMPTY },
  {
    Ref_Key: W(5),
    Description: "Старый склад",
    ТипСклада: "ОптовыйСклад",
    Parent_Key: EMPTY,
    DeletionMark: true,
  },
].map((w) => ({ IsFolder: false, DeletionMark: false, ...w }));

async function setup(schema: FakeSchema = "ut11") {
  const fake = fake1C(
    { READ_ONLY: "true" },
    {
      schema,
      data: (q: FakeQuery) => (q.base === "Catalog_Склады" ? WAREHOUSES : undefined),
    },
  );
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (args: object = {}) => {
    const r = (await client.callTool({ name: "read.catalog.list_warehouses", arguments: args })) as {
      isError?: boolean;
      structuredContent: {
        count: number;
        items: { ref: string; name: string; type?: string; path?: string; isFolder: boolean }[];
      };
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake };
}

describe("read.catalog.list_warehouses", () => {
  it("по умолчанию — склады без папок и помеченных; путь папки и тип", async () => {
    const { call, fake } = await setup();
    const s = (await call()).structuredContent;
    expect(s.count).toBe(4);
    expect(s.items.find((w) => w.ref === W(1))).toMatchObject({
      name: "Магазин 1_основной",
      type: "РозничныйМагазин",
      path: "Магазины/Север",
      isFolder: false,
    });
    expect(s.items.find((w) => w.ref === W(4))?.path).toBeUndefined();
    expect(s.items.some((w) => w.ref === W(5))).toBe(false);
    expect(fake.violations).toEqual([]);
  });

  it("набор «основные склады магазинов» собирается из полей: три склада одного типа различаются названием", async () => {
    const { call } = await setup();
    const s = (await call({ type: "РозничныйМагазин", query: "_ОСНОВНОЙ" })).structuredContent;
    expect(s.items.map((w) => w.ref)).toEqual([W(1)]);
    const retail = (await call({ type: "РозничныйМагазин" })).structuredContent;
    expect(retail.count).toBe(3);
  });

  it("папки, всё и помеченные на удаление — по запросу", async () => {
    const { call } = await setup();
    expect((await call({ kind: "folders" })).structuredContent.items.map((w) => w.ref)).toEqual([
      W(90),
      W(91),
    ]);
    expect((await call({ kind: "all" })).structuredContent.count).toBe(6);
    expect((await call({ includeDeleted: true })).structuredContent.count).toBe(5);
  });

  it("страницы: каждый склад ровно по разу, count — всего", async () => {
    const { call } = await setup();
    const seen: string[] = [];
    let offset = 0;
    for (let guard = 0; guard < 10; guard++) {
      const s = (await call({ limit: 1, offset })).structuredContent as {
        count: number;
        items: { ref: string }[];
        nextOffset?: number;
      };
      expect(s.count).toBe(4);
      seen.push(...s.items.map((w) => w.ref));
      if (s.nextOffset === undefined) break;
      offset = s.nextOffset;
    }
    expect(new Set(seen).size).toBe(4);
    expect(seen).toHaveLength(4);
  });

  it("справочник не опубликован — понятный отказ", async () => {
    const { call, fake } = await setup();
    const available = await fake.conn.available();
    vi.spyOn(fake.conn, "available").mockResolvedValue(
      new Set([...available].filter((s) => !s.startsWith("Catalog_Склады"))),
    );
    const r = await call();
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Склады");
  });

  it("работает и на схеме Бухгалтерии", async () => {
    const { call, fake } = await setup("bp30");
    const r = await call();
    expect(r.isError, r.text).toBeFalsy();
    expect(fake.violations).toEqual([]);
  });
});
