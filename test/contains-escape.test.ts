import { describe, expect, it, vi, afterEach } from "vitest";
import { contains } from "../src/odata/query.js";
import { fake1C } from "./helpers/fake-1c.js";
import { createServer } from "../src/mcp/server.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("contains: символы ПОДОБНО ищутся буквально", () => {
  it.each([
    ["ООО Ромашка", "ООО Ромашка"],
    ["(А4)", "(А4)"],
    ["a_b", "a[_]b"],
    ["100%", "100[%]"],
    ["[кол]", "[[]кол]"],
    ["c:\\x", "c:\\\\x"],
  ])("%s → %s", (input, literal) => {
    expect(contains("Description", input)).toBe(`substringof('${literal}', Description)`);
  });
  it("кавычка экранируется, как раньше", () => {
    expect(contains("Description", "О'Нил")).toBe("substringof('О''Нил', Description)");
  });
});

describe("выбор склада при записи", () => {
  type Tool = {
    inputSchema: { parse: (a: object) => object };
    handler: (a: object, e: object) => Promise<{ isError?: boolean; content: { text: string }[] }>;
  };
  type Row = {
    Ref_Key: string;
    Description: string;
    Code?: string;
    DeletionMark?: boolean;
    IsFolder?: boolean;
  };
  const ref = (i: number) => `aaaaaaaa-0000-4000-8000-${String(i).padStart(12, "0")}`;

  /** Эмуляция 1С: условия на Description, Code, Ref_Key и признаки применяются к строкам. */
  function setup(warehouses: Row[]) {
    const f = fake1C({ READ_ONLY: "true" });
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    const queries: string[] = [];
    vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
      const u = decodeURIComponent(url);
      if (!u.startsWith("Catalog_Склады")) return realGet(url);
      queries.push(u);
      let rows = warehouses;
      if (u.includes("DeletionMark eq false")) rows = rows.filter((r) => !r.DeletionMark);
      if (u.includes("IsFolder eq false")) rows = rows.filter((r) => !r.IsFolder);
      const eqName = /Description eq '([^']*)'/.exec(u)?.[1];
      if (eqName !== undefined) rows = rows.filter((r) => r.Description === eqName);
      const eqCode = /Code eq '([^']*)'/.exec(u)?.[1];
      if (eqCode !== undefined) rows = rows.filter((r) => r.Code === eqCode);
      const eqRef = /Ref_Key eq guid'([^']*)'/.exec(u)?.[1];
      if (eqRef !== undefined) rows = rows.filter((r) => r.Ref_Key === eqRef);
      const part = /substringof\('([^']*)', Description\)/.exec(u)?.[1];
      if (part !== undefined)
        rows = rows.filter((r) => r.Description.toLowerCase().includes(part.toLowerCase()));
      const top = Number(/\$top=(\d+)/.exec(u)?.[1] ?? 1000);
      return { value: rows.slice(0, top) };
    }) as never);
    const tool = (createServer(f.ctx) as unknown as { _registeredTools: Record<string, Tool> })
      ._registeredTools["write.warehouse.create_writeoff"]!;
    const call = (warehouse: string) =>
      tool.handler(
        tool.inputSchema.parse({
          warehouse,
          lines: [{ nomenclatureRef: ref(9), quantity: 1 }],
          date: "2026-11-01",
        }),
        {},
      );
    const callNoWarehouse = () =>
      tool.handler(
        tool.inputSchema.parse({ lines: [{ nomenclatureRef: ref(9), quantity: 1 }], date: "2026-11-01" }),
        {},
      );
    return { call, callNoWarehouse, queries };
  }
  const text = (r: { content: { text: string }[] }) => r.content[0]!.text;

  it("несколько совпадений по части: отказ с кодами и ссылками", async () => {
    const { call } = setup([
      { Ref_Key: ref(1), Description: "Склад основной", Code: "01" },
      { Ref_Key: ref(2), Description: "Склад розничный", Code: "02" },
    ]);
    const r = await call("Склад");
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("подходит несколько");
    expect(text(r)).toContain("код 02");
    expect(text(r)).toContain(ref(2));
    expect(text(r)).toContain("Ничего не записано");
  });
  it("точное название важнее части, даже за пределом выборки", async () => {
    const many: Row[] = Array.from({ length: 25 }, (_, i) => ({
      Ref_Key: ref(i + 10),
      Description: `Склад ${i}`,
    }));
    const { call } = setup([...many, { Ref_Key: ref(1), Description: "Склад" }]);
    expect(text(await call("Склад"))).not.toContain("подходит несколько");
  });
  it("копия с пометкой удаления и папка не мешают", async () => {
    const { call } = setup([
      { Ref_Key: ref(1), Description: "Основной" },
      { Ref_Key: ref(2), Description: "Основной", DeletionMark: true },
      { Ref_Key: ref(3), Description: "Основной", IsFolder: true },
    ]);
    expect(text(await call("Основной"))).not.toContain("подходит несколько");
  });
  it("настоящие дубли: отказ, и по коду или Ref_Key выбрать можно", async () => {
    const rows: Row[] = [
      { Ref_Key: ref(1), Description: "Основной", Code: "01" },
      { Ref_Key: ref(2), Description: "Основной", Code: "02" },
    ];
    const { call } = setup(rows);
    expect(text(await call("Основной"))).toContain("подходит несколько");
    expect(text(await call("02"))).not.toContain("подходит несколько");
    expect(text(await call(ref(2)))).not.toContain("подходит несколько");
    expect(text(await call(ref(2)))).not.toContain("не найден");
  });
  it("пробелы вокруг имени не меняют выбор", async () => {
    const { call, queries } = setup([
      { Ref_Key: ref(1), Description: "Склад" },
      { Ref_Key: ref(2), Description: "Склад розничный" },
    ]);
    await call("  Склад ");
    expect(queries[0]).toContain("Description eq 'Склад'");
  });
  it("строка из пробелов — отказ, а не выбор единственного", async () => {
    const { call } = setup([{ Ref_Key: ref(1), Description: "Основной" }]);
    const r = await call("   ");
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("пустое значение");
  });
  it("без названия: помеченный склад и папка не считаются, рабочий выбирается", async () => {
    const { queries, callNoWarehouse } = setup([
      { Ref_Key: ref(1), Description: "Основной" },
      { Ref_Key: ref(2), Description: "Копия", DeletionMark: true },
      { Ref_Key: ref(3), Description: "Папка", IsFolder: true },
    ]);
    await callNoWarehouse();
    expect(queries[0]).toContain("DeletionMark eq false");
    expect(queries[0]).toContain("IsFolder eq false");
  });
  it("название из цифр у помеченного склада: отказ, а не склад с таким кодом", async () => {
    const { call } = setup([
      { Ref_Key: ref(1), Description: "12", DeletionMark: true },
      { Ref_Key: ref(2), Description: "Другой", Code: "12" },
    ]);
    const r = await call("12");
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("помечен на удаление");
  });
  it("точное название на удалении: другой склад по части не подставляется", async () => {
    const { call } = setup([
      { Ref_Key: ref(1), Description: "Метка", DeletionMark: true },
      { Ref_Key: ref(2), Description: "Метка новая" },
    ]);
    const r = await call("Метка");
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("помечен на удаление");
    expect(text(await call("Метка новая"))).not.toContain("помечен на удаление");
  });
  it("нет такого склада — отказ", async () => {
    const { call } = setup([{ Ref_Key: ref(1), Description: "Основной" }]);
    expect(text(await call("Другой"))).toContain("не найден");
  });
  it("справочник не опубликован: заданный склад не игнорируется", async () => {
    const f = fake1C({ READ_ONLY: "true" });
    const meta = await f.conn.getMetadata();
    const entities = new Map([...meta.entities].filter(([k]) => !k.startsWith("Catalog_Склады")));
    vi.spyOn(f.conn, "available").mockResolvedValue(new Set(entities.keys()));
    const tool = (createServer(f.ctx) as unknown as { _registeredTools: Record<string, Tool> })
      ._registeredTools["write.warehouse.create_writeoff"]!;
    const r = await tool.handler(
      tool.inputSchema.parse({
        warehouse: "Склад",
        lines: [{ nomenclatureRef: ref(9), quantity: 1 }],
        date: "2026-11-01",
      }),
      {},
    );
    expect(text(r)).toContain("не опубликован");
  });
});
