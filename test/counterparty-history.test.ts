import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

// История покупателя и поставщика возвращала помеченные на удаление документы без признака пометки: по строке
// нельзя было отличить их от живых, и агент мог включить их в сумму. Тесты истории не было вовсе.
const CP = "dddddddd-0000-4000-8000-000000000001";
afterEach(() => vi.restoreAllMocks());

describe.each([
  ["write" as const, "read.counterparty.get_customer_history", "Document_РеализацияТоваровУслуг"],
  ["write" as const, "read.counterparty.get_supplier_history", "Document_ПоступлениеТоваровУслуг"],
])("%s: %s", (_, tool, set) => {
  it("в строках есть deletionMark: помеченный отличим от живого", async () => {
    const f = fake1C();
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (!decodeURIComponent(path).startsWith(`${set}?`) && !decodeURIComponent(path).startsWith(set))
        return { value: [] } as never;
      return {
        value: [
          {
            Ref_Key: "a",
            Number: "1",
            Date: "2026-11-02T00:00:00",
            Posted: true,
            DeletionMark: false,
            СуммаДокумента: 100,
          },
          {
            Ref_Key: "b",
            Number: "2",
            Date: "2026-11-01T00:00:00",
            Posted: true,
            DeletionMark: true,
            СуммаДокумента: 900,
          },
        ],
      } as never;
    });
    const tools = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools;
    const r = await tools[tool]!.handler({ ref: CP, limit: 10 }, {});
    expect(r.isError, JSON.stringify(r)).toBeFalsy();
    const rows = (r.structuredContent as { rows: Array<Record<string, unknown>> }).rows;
    expect(rows.map((x) => [x["number"], x["deletionMark"]])).toEqual([
      ["1", false],
      ["2", true],
    ]);
  });
});
