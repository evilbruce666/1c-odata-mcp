import { afterEach, describe, expect, it, vi } from "vitest";
import { parseConfig } from "../src/config/env.js";
import { ServerContext } from "../src/context.js";
import { createServer } from "../src/mcp/server.js";
import { contractSettlementKind } from "../src/odata/settlement-accounts.js";
import type { CallToolResult } from "@modelcontextprotocol/server";

const single = { ODATA_BASE_URL: "https://example.invalid/odata/", ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };
const multi = {
  ODATA_DB_BY_BASE_URL: single.ODATA_BASE_URL,
  ODATA_DB_BY_USERNAME: "u",
  ODATA_DB_BY_PASSWORD: "p",
  ODATA_DB_RU_BASE_URL: single.ODATA_BASE_URL,
  ODATA_DB_RU_USERNAME: "u",
  ODATA_DB_RU_PASSWORD: "p",
};
afterEach(() => vi.restoreAllMocks());
describe("regulated currency per connection", () => {
  it.each([undefined, "", "  "])("preserves default643 for missing or blank global (%s)", (value) => {
    const ctx = new ServerContext(parseConfig({ ...single, ODATA_REGULATED_CURRENCY_CODE: value }));
    expect(ctx.db().behavior.regulatedCurrencyCode).toBe("643");
  });
  it("retains conditional-unit classification outside the bank mutation tool", () => {
    expect(contractSettlementKind({ РасчетыВУсловныхЕдиницах: true }, "840", "643")).toBe(
      "РасчетыВУсловныхЕдиницах",
    );
  });
  it("uses the global setting in single-base mode", () => {
    const ctx = new ServerContext(parseConfig({ ...single, ODATA_REGULATED_CURRENCY_CODE: " 933 " }));
    expect(ctx.db().behavior.regulatedCurrencyCode).toBe("933");
  });
  it("isolates Belarus933 and Russia643 in one process, retaining global behavior", () => {
    const cfg = parseConfig({ ...multi, ODATA_DB_BY_REGULATED_CURRENCY_CODE: " 933 " });
    const ctx = new ServerContext(cfg);
    for (const [name, code] of [
      ["by", "933"],
      ["ru", "643"],
    ]) {
      expect(ctx.db(name).behavior.regulatedCurrencyCode).toBe(code);
      expect(
        contractSettlementKind(
          { РасчетыВУсловныхЕдиницах: false },
          code!,
          ctx.db(name).behavior.regulatedCurrencyCode!,
        ),
      ).toBe("РасчетыВВалютеРегламентированногоУчета");
    }
    expect(ctx.db("by").behavior).not.toBe(cfg.behavior);
    expect(cfg.behavior.regulatedCurrencyCode).toBe("643");
  });
  it("per-base overrides global; blank per-base falls back to global", () => {
    const ctx = new ServerContext(
      parseConfig({
        ...multi,
        ODATA_REGULATED_CURRENCY_CODE: "933",
        ODATA_DB_RU_REGULATED_CURRENCY_CODE: "643",
        ODATA_DB_BY_REGULATED_CURRENCY_CODE: " ",
      }),
    );
    expect(ctx.db("ru").behavior.regulatedCurrencyCode).toBe("643");
    expect(ctx.db("by").behavior.regulatedCurrencyCode).toBe("933");
  });
  it.each(["933\n643", "933\0", "x".repeat(101)])(
    "rejects invalid global and per-base values %s",
    (value) => {
      expect(() => parseConfig({ ...single, ODATA_REGULATED_CURRENCY_CODE: value })).toThrow();
      expect(() => parseConfig({ ...multi, ODATA_DB_BY_REGULATED_CURRENCY_CODE: value })).toThrow();
    },
  );
  it("preserves custom catalog codes without inferring ISO currency", () => {
    const ctx = new ServerContext(parseConfig({ ...single, ODATA_REGULATED_CURRENCY_CODE: " LOCAL " }));
    expect(ctx.db().behavior.regulatedCurrencyCode).toBe("LOCAL");
  });
  it.each([
    [{ ...multi, ODATA_DB_BY_REGULATED_CURRENCY_CODE: "933" }, "by", "933", "database"],
    [{ ...single, ODATA_REGULATED_CURRENCY_CODE: "933" }, "default", "933", "global"],
    [single, "default", "643", "default"],
  ] as const)(
    "health exposes effective currency and its actual configuration source",
    async (env, name, code, source) => {
      const ctx = new ServerContext(parseConfig(env));
      vi.spyOn(ctx.db(name).client, "refreshMetadata").mockResolvedValue({
        entities: new Map(),
        odataVersion: "3.0",
      });
      vi.spyOn(ctx.db(name).client, "journalSummary").mockResolvedValue(undefined);
      const tools = (
        createServer(ctx) as unknown as {
          _registeredTools: Record<
            string,
            { handler: (args: object, extra: object) => Promise<CallToolResult> }
          >;
        }
      )._registeredTools;
      const result = await tools["read.system.health_check"]!.handler({ database: name }, {});
      expect(result.isError).toBeFalsy();
      expect(result.structuredContent).toMatchObject({ regulatedCurrency: { code, source } });
    },
  );
});
