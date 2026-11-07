import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { parseConfig } from "../src/config/env.js";
import { createServer, limitResponseSize, summaryFirst } from "../src/mcp/server.js";
import { withToolCall } from "../src/odata/call-context.js";
import { ok, responseCharsLimit } from "../src/tools/_shared.js";
import { fake1C } from "./helpers/fake-1c.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const textOf = (r: CallToolResult) => (r.content[0] as { text: string }).text;
const row = (i: number) => ({ id: i, text: "я".repeat(200) });
const rows = (n: number) => Array.from({ length: n }, (_, i) => row(i));
const atLimit = <T>(maxResponseChars: number, fn: () => T) =>
  withToolCall({ deadlineMs: 0, maxResponseChars }, async () => fn());

describe("ODATA_MAX_RESPONSE_CHARS", () => {
  const base = { ODATA_BASE_URL: "http://1c.test/", ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };
  it.each([
    [undefined, 45_000],
    ["10000", 10_000],
    ["45000", 45_000],
  ])("%s → %s", (value, expected) => {
    const cfg = parseConfig({ ...base, ...(value ? { ODATA_MAX_RESPONSE_CHARS: value } : {}) });
    expect(cfg.behavior.maxResponseChars).toBe(expected);
  });
  it.each(["9999", "45001", "много"])("%s — отказ при старте", (value) => {
    expect(() => parseConfig({ ...base, ODATA_MAX_RESPONSE_CHARS: value })).toThrow(
      /ODATA_MAX_RESPONSE_CHARS/,
    );
  });
  it("вне вызова действует предел по умолчанию", () => {
    expect(responseCharsLimit()).toBe(45_000);
  });
});

describe("настроенный предел строгий", () => {
  it("страницы не больше предела и ничего не теряется", async () => {
    await atLimit(10_000, () => {
      const seen: number[] = [];
      let offset = 0;
      for (let i = 0; i < 100; i++) {
        const r = limitResponseSize(ok({ total: 500, rows: rows(500) }), { offset, continuable: true });
        expect(textOf(r).length).toBeLessThanOrEqual(10_000);
        const s = r.structuredContent as { rows: { id: number }[]; nextOffset?: number };
        seen.push(...s.rows.map((x) => x.id));
        if (s.nextOffset === undefined) break;
        offset = s.nextOffset;
      }
      expect(seen).toEqual(rows(500).map((x) => x.id));
    });
  });

  it("настоящий инструмент на пределе", async () => {
    const f = fake1C({ ODATA_MAX_RESPONSE_CHARS: "10000" });
    const meta = await f.conn.getMetadata();
    const entities = new Map(
      Array.from({ length: 2000 }, (_, i) => [
        `Catalog_Объект${i}`,
        { entitySet: `Catalog_Объект${i}`, shortName: `Объект${i}`, class: "catalog" },
      ]),
    );
    vi.spyOn(f.conn, "getMetadata").mockResolvedValue({ ...meta, entities } as never);
    const tool = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<
          string,
          {
            inputSchema: { parse: (a: object) => object };
            handler: (a: object, e: object) => Promise<CallToolResult>;
          }
        >;
      }
    )._registeredTools["read.schema.list_entities"]!;
    const r = await tool.handler(tool.inputSchema.parse({}), {});
    expect(textOf(r).length).toBeLessThanOrEqual(10_000);
    expect((r.structuredContent as { truncated: boolean }).truncated).toBe(true);
  });

  it("без настройки ответ между 45 и 50 тысячами не режется, как раньше", () => {
    const r = ok({ rows: rows(100), pad: "x".repeat(24_000) });
    expect(textOf(r).length).toBeGreaterThan(45_000);
    expect(textOf(r).length).toBeLessThan(50_000);
    expect(limitResponseSize(r)).toBe(r);
  });
});

describe("вложенные списки", () => {
  const nested = () => ({
    status: "warning",
    checks: [
      { name: "первая", items: rows(1) },
      { name: "большая", items: rows(1000) },
    ],
  });

  it("элемент с большим списком внутри укладывается, срез назван", () => {
    const r = limitResponseSize(ok(nested()));
    expect(textOf(r).length).toBeLessThanOrEqual(50_000);
    const s = r.structuredContent as {
      truncated: boolean;
      truncatedLists: { path: string; total: number; returned: number }[];
      nextOffset?: number;
      note: string;
      checks: { name: string; items: unknown[] }[];
    };
    expect(s.truncated).toBe(true);
    expect(s.nextOffset).toBeUndefined();
    expect(s.truncatedLists[0]).toMatchObject({ path: "checks[1].items", total: 1000 });
    expect(s.checks[1]!.items).toHaveLength(s.truncatedLists[0]!.returned);
    expect(s.checks.map((c) => c.name)).toEqual(["первая", "большая"]);
    expect(s.note).toContain("checks[1].items");
  });

  it("объект без списков верхнего уровня: табличная часть документа", () => {
    const data = { document: { Number: "1", Товары: rows(1000) } };
    const s = limitResponseSize(ok(data)).structuredContent as {
      document: { Number: string; Товары: unknown[] };
      truncatedLists: { path: string }[];
    };
    expect(s.document.Number).toBe("1");
    expect(s.truncatedLists[0]!.path).toBe("document.Товары");
    expect(data.document.Товары).toHaveLength(1000); // исходные данные не меняются
  });

  it("настроенный предел действует и на вложенные", async () => {
    await atLimit(10_000, () => {
      expect(textOf(limitResponseSize(ok(nested()))).length).toBeLessThanOrEqual(10_000);
    });
  });

  it("большой ответ без списков возвращается как есть", () => {
    const r = ok({ text: "я".repeat(60_000) });
    expect(limitResponseSize(r)).toBe(r);
  });
});

describe("итоги первыми", () => {
  it("простые поля раньше объектов, самый большой последним, текст совпадает", () => {
    const r = summaryFirst(
      ok({ rows: rows(3), scan: { requests: 1 }, total: 3, note: "важно", truncated: false }),
    );
    expect(Object.keys(r.structuredContent!)).toEqual(["total", "note", "truncated", "scan", "rows"]);
    expect(JSON.parse(textOf(r))).toEqual(r.structuredContent);
    expect(textOf(r).indexOf('"total"')).toBeLessThan(textOf(r).indexOf('"rows"'));
  });

  it("служебные поля обрезки идут раньше списка", () => {
    const r = summaryFirst(limitResponseSize(ok({ total: 2000, rows: rows(2000) }), { continuable: true }));
    const keys = Object.keys(r.structuredContent!);
    expect(keys.at(-1)).toBe("rows");
    expect(keys.indexOf("nextOffset")).toBeLessThan(keys.indexOf("rows"));
  });

  it("ошибка не трогается", () => {
    const err = { content: [{ type: "text" as const, text: "нет" }], isError: true };
    expect(summaryFirst(err)).toBe(err);
  });
});
