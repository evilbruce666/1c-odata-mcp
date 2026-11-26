import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { SNAPSHOT_REF as id } from "./helpers/metadata-snapshot.js";
import { fake1C } from "./helpers/fake-1c.js";

// Найдено проверкой совместимости (npm run compat) на схеме рабочей Бухгалтерии без регистра счетов расчётов:
// apply_bank_settlement_settings отвечал «Неверные аргументы: expected string, received undefined». Тест всех
// инструментов шёл только на снимке, где регистр есть.
afterEach(() => vi.restoreAllMocks());
const ARGS = {
  entitySet: "Document_ПоступлениеНаРасчетныйСчет",
  ref: id,
  contractRef: id,
  side: "customer",
  expectedDataVersion: "AAAAAAAAAAA=",
  confirm: false,
};

describe("apply_bank_settlement_settings без регистра счетов расчётов", () => {
  it("понятный отказ с причиной, без записи", async () => {
    const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const published = f.conn.available.bind(f.conn);
    vi.spyOn(f.conn, "available").mockImplementation(async () => {
      const set = new Set(await published());
      set.delete("InformationRegister_СчетаУчетаРасчетовСКонтрагентами");
      return set;
    });
    const tools = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools;
    const r = await tools["write.document.apply_bank_settlement_settings"]!.handler(ARGS, {});
    const text = JSON.stringify(r);
    expect(r.isError).toBe(true);
    expect(text).toMatch(/не опубликован в OData/);
    expect(text).not.toMatch(/expected string/);
    expect(f.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });
});
