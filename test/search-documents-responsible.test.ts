import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeSchema } from "./helpers/fake-1c.js";

// 1.1: отбор search_documents по ответственному — «Ответственный», иначе «Менеджер»; нет обоих — неприменим.
const U = "cccccccc-0000-4000-8000-000000000001";
function search(schema: FakeSchema) {
  const f = fake1C({}, { schema });
  const tools = (
    createServer(f.ctx) as unknown as {
      _registeredTools: Record<string, { handler: (args: object, extra: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  const getCollection = vi.spyOn(f.conn.client, "getCollection");
  const run = (args: object) =>
    tools["read.document.search_documents"]!.handler(
      { postedOnly: false, limit: 50, offset: 0, ...args },
      {},
    );
  const filters = () => getCollection.mock.calls.map((c) => decodeURIComponent(String(c[0])));
  return { f, run, filters };
}
afterEach(() => vi.restoreAllMocks());

describe("search_documents: responsibleRef", () => {
  it.each([
    ["bp30", "Document_РеализацияТоваровУслуг", "Ответственный_Key"],
    ["ut11", "Document_РеализацияТоваровУслуг", "Менеджер_Key"],
    ["ut11", "Document_ПеремещениеТоваров", "Ответственный_Key"],
  ] as const)("%s %s — отбор по %s на стороне 1С", async (schema, entitySet, field) => {
    const { f, run, filters } = search(schema);
    const r = await run({ entitySet, responsibleRef: `{${U}}` });
    expect(r.isError, JSON.stringify(r.content)).toBeFalsy();
    expect(filters().some((q) => q.startsWith(entitySet) && q.includes(`${field} eq guid'${U}'`))).toBe(true);
    // При отборе — ссылка и имя ответственного у найденных документов.
    const rows = (r.structuredContent as { rows: Array<{ responsibleRef?: string; responsible?: string }> })
      .rows;
    expect(rows.length).toBeGreaterThan(0);
    expect(rows.every((x) => x.responsibleRef === U && typeof x.responsible === "string")).toBe(true);
    expect(f.violations).toEqual([]);
  });

  it("ut11: явный вид без «Ответственный» и «Менеджер» (ПКО) — отказ без запроса к документам", async () => {
    const { f, run } = search("ut11");
    const r = await run({ entitySet: "Document_ПриходныйКассовыйОрдер", responsibleRef: U });
    expect(r.isError).toBe(true);
    expect(JSON.stringify(r.content)).toContain("фильтр по ответственному неприменим");
    expect(f.requests.some((q) => q.resource.startsWith("Document_ПриходныйКассовыйОрдер"))).toBe(false);
  });

  it("без responsibleRef фильтра по ответственному нет", async () => {
    const { run, filters } = search("bp30");
    await run({ entitySet: "Document_РеализацияТоваровУслуг" });
    expect(filters().some((q) => q.includes("Ответственный_Key eq"))).toBe(false);
  });
});
