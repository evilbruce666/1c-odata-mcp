import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

// #45: minAmount фильтрует по СуммаДокумента, а у перемещения и инвентаризации этого поля нет.
function searchDocuments(ctx: ReturnType<typeof fake1C>["ctx"]) {
  const tools = (
    createServer(ctx) as unknown as {
      _registeredTools: Record<string, { handler: (args: object, extra: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  // Значения по умолчанию из схемы (MCP подставляет их до вызова обработчика).
  return (args: object) =>
    tools["read.document.search_documents"]!.handler({ postedOnly: false, limit: 50, ...args }, {});
}
const textOf = (r: CallToolResult) => JSON.stringify(r.content);
afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("search_documents: minAmount (#45)", () => {
  it.each(["Document_ПеремещениеТоваров", "Document_ИнвентаризацияТоваровНаСкладе"])(
    "%s без СуммаДокумента — понятный отказ без запроса к документам",
    async (entitySet) => {
      const f = fake1C();
      const result = await searchDocuments(f.ctx)({ entitySet, minAmount: 1000 });
      expect(result.isError).toBe(true);
      expect(textOf(result)).toContain("minAmount");
      expect(textOf(result)).toContain("неприменим");
      expect(f.requests.some((r) => r.resource.startsWith(entitySet))).toBe(false);
      expect(f.violations).toEqual([]);
    },
  );

  it("у реализации фильтр по сумме работает, как раньше", async () => {
    const f = fake1C();
    const getCollection = vi.spyOn(f.conn.client, "getCollection");
    const result = await searchDocuments(f.ctx)({
      entitySet: "Document_РеализацияТоваровУслуг",
      minAmount: 1000,
    });
    expect(result.isError).toBeFalsy();
    expect(decodeURIComponent(String(getCollection.mock.calls[0]?.[0]))).toContain("СуммаДокумента ge 1000");
    expect(f.violations).toEqual([]);
  });

  it("без minAmount перемещение ищется как раньше", async () => {
    const f = fake1C();
    const result = await searchDocuments(f.ctx)({ entitySet: "Document_ПеремещениеТоваров" });
    expect(result.isError).toBeFalsy();
    expect(f.violations).toEqual([]);
  });
});
