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
