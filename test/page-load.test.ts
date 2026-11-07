import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

async function connect(f: ReturnType<typeof fake1C>) {
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(f.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  return client;
}

describe("должники и кредиторы с aging: настоящий клиент", () => {
  it.each([
    ["read.analytics.get_debtors", "debtors"],
    ["read.analytics.get_creditors", "creditors"],
  ])("%s: aging и overdueOnly отвечают, count — число", async (name, list) => {
    const client = await connect(fake1C({ READ_ONLY: "true" }));
    for (const args of [{}, { aging: true }, { overdueOnly: true }, { aging: true, limit: 5 }]) {
      const r = (await client.callTool({ name, arguments: args })) as {
        isError?: boolean;
        content: { text: string }[];
        structuredContent: Record<string, unknown>;
      };
      expect(r.isError, `${JSON.stringify(args)}: ${r.content[0]!.text.slice(0, 200)}`).toBeFalsy();
      expect(typeof r.structuredContent["count"]).toBe("number");
      expect(Array.isArray(r.structuredContent[list])).toBe(true);
    }
  });
});

describe("aging: обход всех страниц со строками", () => {
  const ref = (i: number) => `cccccccc-0000-4000-8000-${String(i).padStart(12, "0")}`;
  const balance = Array.from({ length: 60 }, (_, i) => ({
    Account_Key: "",
    ExtDimension1: ref(i),
    ExtDimension1_Type: "StandardODATA.Catalog_Контрагенты",
    СуммаBalanceDr: 100 + (i % 7),
    СуммаBalanceCr: 0,
  }));
  const total = balance.reduce((s, r) => s + r.СуммаBalanceDr, 0);

  function setup() {
    const f = fake1C({ READ_ONLY: "true" });
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    const nameFilters: string[] = [];
    vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
      const u = decodeURIComponent(url);
      if (u.includes("/Balance")) return { value: balance.map((r) => ({ ...r })) };
      if (u.startsWith("Catalog_Контрагенты")) {
        nameFilters.push(u);
        const refs = [...u.matchAll(/guid'([^']+)'/g)].map((m) => m[1]!);
        return { value: refs.map((Ref_Key) => ({ Ref_Key, Description: `Имя ${Ref_Key}` })) };
      }
      return realGet(url);
    }) as never);
    return { f, nameFilters };
  }

  interface Debtor {
    ref: string;
    counterparty: string;
    amount: number;
    debt: number;
    overdue: number;
  }

  it.each([[25], [1000]])(
    "limit %s: все строки с именами, суммы сходятся, продолжение закрывается",
    async (limit) => {
      const { f } = setup();
      const client = await connect(f);
      const seen: Debtor[] = [];
      let offset = 0;
      for (let i = 0; i < 80; i++) {
        const r = (await client.callTool({
          name: "read.analytics.get_debtors",
          arguments: { aging: true, limit, offset },
        })) as {
          isError?: boolean;
          content: { text: string }[];
          structuredContent: {
            debtors: Debtor[];
            count: number;
            totalReceivable: number;
            nextOffset?: number;
          };
        };
        expect(r.isError, r.content[0]!.text.slice(0, 200)).toBeFalsy();
        const s = r.structuredContent;
        expect(s.count).toBe(60);
        expect(s.totalReceivable).toBe(total);
        seen.push(...s.debtors);
        if (s.nextOffset === undefined) break;
        offset = s.nextOffset;
      }
      expect(new Set(seen.map((d) => d.ref)).size).toBe(60);
      expect(seen).toHaveLength(60);
      expect(seen.reduce((t, d) => t + d.amount, 0)).toBe(total);
      for (const d of seen) {
        expect(d.counterparty).toBe(`Имя ${d.ref}`);
        expect(d.debt).toBe(d.amount);
        expect(typeof d.overdue).toBe("number");
      }
    },
  );

  it("overdueOnly с должниками: ответ строится, отказа нет", async () => {
    const { f } = setup();
    const client = await connect(f);
    const r = (await client.callTool({
      name: "read.analytics.get_debtors",
      arguments: { overdueOnly: true, limit: 10 },
    })) as { isError?: boolean; content: { text: string }[]; structuredContent: { count: number } };
    expect(r.isError, r.content[0]!.text.slice(0, 200)).toBeFalsy();
    expect(typeof r.structuredContent.count).toBe("number");
  });
});

describe("нагрузка на 1С при чтении по offset", () => {
  const stub = (f: ReturnType<typeof fake1C>, rows: object[]) => {
    const stats = { contracts: [] as string[] };
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
      const u = decodeURIComponent(url);
      if (u.startsWith("Catalog_Контрагенты")) {
        const top = Number(/\$top=(\d+)/.exec(u)?.[1] ?? 1000);
        const skip = Number(/\$skip=(\d+)/.exec(u)?.[1] ?? 0);
        return { value: rows.slice(skip, skip + top) };
      }
      if (u.startsWith("Catalog_ДоговорыКонтрагентов")) {
        stats.contracts.push(u);
        return { value: [] };
      }
      return realGet(url);
    }) as never);
    return stats;
  };
  const counterparties = (n: number) =>
    Array.from({ length: n }, (_, i) => ({
      Ref_Key: `cccccccc-0000-4000-8000-${String(i).padStart(12, "0")}`,
      Description: `Контрагент ${i}`,
      IsFolder: false,
    }));

  it("find_counterparty: 1000 строк, все на месте, роли только для своей страницы", async () => {
    const f = fake1C({ READ_ONLY: "true" });
    const stats = stub(f, counterparties(1000));
    const client = await connect(f);
    const seen: string[] = [];
    let offset = 0;
    for (let i = 0; i < 100; i++) {
      const r = (await client.callTool({
        name: "read.counterparty.find_counterparty",
        arguments: { query: "Контрагент", limit: 100, offset },
      })) as { structuredContent: { rows: { ref: string }[]; nextOffset?: number } };
      seen.push(...r.structuredContent.rows.map((x) => x.ref));
      if (r.structuredContent.nextOffset === undefined) break;
      offset = r.structuredContent.nextOffset;
    }
    expect(new Set(seen).size).toBe(seen.length);
    expect(seen.length).toBeGreaterThanOrEqual(1000);
    // По 5 пачек на страницу; раньше их становилось всё больше, до 50 на последней.
    expect(stats.contracts.length).toBeLessThanOrEqual(60);
  });

  it("ровно ODATA_MAX_ROWS строк: всё прочитано, признака неполноты нет", async () => {
    const f = fake1C({ READ_ONLY: "true", ODATA_MAX_ROWS: "30" });
    stub(f, counterparties(30));
    const client = await connect(f);
    const r = (await client.callTool({
      name: "read.counterparty.find_counterparty",
      arguments: { query: "Контрагент", limit: 100 },
    })) as { structuredContent: { rows: unknown[]; truncated: boolean; note?: string } };
    expect(r.structuredContent.rows).toHaveLength(30);
    expect(r.structuredContent.truncated).toBe(false);
    expect(r.structuredContent.note).toBeUndefined();
  });

  it("строк больше предела выборки: неполнота названа", async () => {
    const f = fake1C({ READ_ONLY: "true", ODATA_MAX_ROWS: "30" });
    stub(f, counterparties(31));
    const client = await connect(f);
    const r = (await client.callTool({
      name: "read.counterparty.find_counterparty",
      arguments: { query: "Контрагент", limit: 100 },
    })) as { structuredContent: { rows: unknown[]; truncated: boolean; note?: string } };
    expect(r.structuredContent.rows).toHaveLength(30);
    expect(r.structuredContent.truncated).toBe(true);
    expect(r.structuredContent.note).toContain("30");
  });

  it("роли строк до начала страницы не запрашиваются", async () => {
    const f = fake1C({ READ_ONLY: "true" });
    const rows = counterparties(60);
    const stats = stub(f, rows);
    const client = await connect(f);
    await client.callTool({
      name: "read.counterparty.find_counterparty",
      arguments: { query: "Контрагент", limit: 10, offset: 50 },
    });
    const all = stats.contracts.join(" ");
    expect(all).toContain(rows[50]!.Ref_Key);
    expect(all).not.toContain(rows[0]!.Ref_Key);
    expect(stats.contracts).toHaveLength(1);
  });
});
