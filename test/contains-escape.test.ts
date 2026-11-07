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

describe("запись не берёт первое совпадение по части названия", () => {
  type Tool = {
    inputSchema: { parse: (a: object) => object };
    handler: (a: object, e: object) => Promise<{ isError?: boolean; content: { text: string }[] }>;
  };
  function setup(warehouses: { Ref_Key: string; Description: string }[]) {
    const f = fake1C({ READ_ONLY: "true" });
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
      const u = decodeURIComponent(url);
      if (u.startsWith("Catalog_Склады")) return { value: warehouses };
      return realGet(url);
    }) as never);
    return (createServer(f.ctx) as unknown as { _registeredTools: Record<string, Tool> })._registeredTools[
      "write.warehouse.create_writeoff"
    ]!;
  }
  const ref = (i: number) => `aaaaaaaa-0000-4000-8000-${String(i).padStart(12, "0")}`;
  const args = (warehouse: string) => ({
    warehouse,
    lines: [{ nomenclatureRef: ref(9), quantity: 1 }],
    date: "2026-11-01",
  });

  it("несколько совпадений: отказ со списком, ничего не записано", async () => {
    const tool = setup([
      { Ref_Key: ref(1), Description: "Склад основной" },
      { Ref_Key: ref(2), Description: "Склад розничный" },
    ]);
    const r = await tool.handler(tool.inputSchema.parse(args("Склад")), {});
    expect(r.isError).toBe(true);
    expect(r.content[0]!.text).toContain("подходит несколько");
    expect(r.content[0]!.text).toContain("Склад розничный");
    expect(r.content[0]!.text).toContain("Ничего не записано");
  });
  it("точное название среди похожих проходит этот шаг", async () => {
    const tool = setup([
      { Ref_Key: ref(1), Description: "Склад" },
      { Ref_Key: ref(2), Description: "Склад розничный" },
    ]);
    const r = await tool.handler(tool.inputSchema.parse(args("склад")), {});
    expect(r.content[0]!.text).not.toContain("подходит несколько");
  });
  it("единственное совпадение по части проходит этот шаг", async () => {
    const tool = setup([{ Ref_Key: ref(1), Description: "Склад основной" }]);
    const r = await tool.handler(tool.inputSchema.parse(args("основн")), {});
    expect(r.content[0]!.text).not.toContain("подходит несколько");
  });
});
