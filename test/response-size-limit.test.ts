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

  it("много мелких списков тоже укладываются в настроенный предел, срезы названы", async () => {
    const lists = Object.fromEntries(Array.from({ length: 30 }, (_, i) => [`l${i}`, rows(8)]));
    await atLimit(10_000, () => {
      const r = limitResponseSize(ok(lists));
      expect(textOf(r).length).toBeLessThanOrEqual(10_000);
      const s = r.structuredContent as { truncated: boolean; truncatedLists: unknown[]; overLimit?: boolean };
      expect(s.truncated).toBe(true);
      expect(s.truncatedLists.length).toBeGreaterThan(0);
      expect(s.overLimit).toBeUndefined();
    });
  });

  it("сократить нечем: признак превышения, данные целы", () => {
    const s = limitResponseSize(ok({ text: "я".repeat(60_000) })).structuredContent as {
      text: string;
      overLimit: boolean;
      note: string;
    };
    expect(s.overLimit).toBe(true);
    expect(s.text).toHaveLength(60_000);
    expect(s.note).toContain("больше предела");
  });

  it("срезы с примечанием не выходят за норму", () => {
    const r = limitResponseSize(ok({ checks: [{ a: rows(1000) }, { b: rows(1000) }], note: "итог" }));
    expect(textOf(r).length).toBeLessThanOrEqual(50_000);
  });

  it("путь среза на второй странице — от начала списка", () => {
    const big = Array.from({ length: 6 }, (_, i) => ({ id: i, lines: rows(1000) }));
    const s = limitResponseSize(ok({ rows: big }), { offset: 3, continuable: true }).structuredContent as {
      truncatedLists: { path: string }[];
    };
    expect(s.truncatedLists[0]!.path).toBe("rows[3].lines");
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

describe("чтение по nextOffset при любом пределе: ни одна строка не теряется", () => {
  type Tool = {
    inputSchema: { parse: (a: object) => object };
    handler: (a: object, e: object) => Promise<CallToolResult>;
  };
  const SET = "Document_СписаниеСРасчетногоСчета";
  const LIMITS = [undefined, "10000", "11000", "12000", "15000", "20000", "30000", "44000", "45000"];

  async function walk(limit: string | undefined) {
    const f = fake1C(limit ? { ODATA_MAX_RESPONSE_CHARS: limit } : {});
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    const docs = Array.from({ length: 700 }, (_, i) => ({
      Ref_Key: `d${i}`,
      Number: `0000-${String(i).padStart(6, "0")}`,
      Date: `2026-0${(i % 9) + 1}-15T00:00:00`,
      Posted: true,
      DeletionMark: false,
      СуммаДокумента: i,
      НомерВходящегоДокумента: "описание длинного входящего номера документа для объёма",
    })).sort((a, b) => b.Date.localeCompare(a.Date) || a.Ref_Key.localeCompare(b.Ref_Key));
    vi.spyOn(f.conn.client, "getCollection").mockImplementation((async (url: string) => {
      const u = decodeURIComponent(url);
      if (!u.startsWith(SET)) return realGet(url);
      const top = Number(/\$top=(\d+)/.exec(u)![1]);
      const skip = Number(/\$skip=(\d+)/.exec(u)?.[1] ?? 0);
      return { value: docs.slice(skip, skip + top) };
    }) as never);
    const tool = (createServer(f.ctx) as unknown as { _registeredTools: Record<string, Tool> })
      ._registeredTools["read.document.search_documents"]!;
    const seen: string[] = [];
    let offset = 0;
    for (let i = 0; i < 300; i++) {
      const r = await tool.handler(tool.inputSchema.parse({ entitySet: SET, limit: 500, offset }), {});
      const s = r.structuredContent as {
        rows: { ref: string }[];
        offset: number;
        returned: number;
        nextOffset?: number;
      };
      expect(textOf(r).length).toBeLessThanOrEqual(limit && Number(limit) < 45_000 ? Number(limit) : 50_000);
      expect(s.offset).toBe(offset);
      expect(s.returned).toBe(s.rows.length);
      seen.push(...s.rows.map((x) => x.ref));
      if (s.nextOffset === undefined) break;
      expect(s.nextOffset).toBe(offset + s.rows.length);
      offset = s.nextOffset;
    }
    return seen;
  }

  it.each(LIMITS)("search_documents, предел %s: 700 из 700 без повторов", async (limit) => {
    const seen = await walk(limit);
    expect(seen).toHaveLength(700);
    expect(new Set(seen).size).toBe(700);
  });
});

describe("срез чужой страницы сдвигает продолжение", () => {
  const page = (offset: number) =>
    ok({
      offset,
      returned: 100,
      truncated: true,
      nextOffset: offset + 100,
      note: `Показаны документы ${offset + 1}–${offset + 100}. Следующая страница: offset=${offset + 100}.`,
      other: [1],
      rows: rows(100),
    });

  it("nextOffset — следующая непрочитанная строка, offset не затирается, диапазон один", async () => {
    await atLimit(10_000, () => {
      const s = limitResponseSize(page(200)).structuredContent as {
        offset: number;
        returned: number;
        nextOffset: number;
        note: string;
        rows: unknown[];
      };
      expect(s.rows.length).toBeLessThan(100);
      expect(s.offset).toBe(200);
      expect(s.returned).toBe(s.rows.length);
      expect(s.nextOffset).toBe(200 + s.rows.length);
      expect(s.note).toContain(`offset=${s.nextOffset}`);
      expect(s.note).not.toContain("offset=300");
    });
  });

  it("страница без среза не меняется", () => {
    const small = ok({ offset: 5, returned: 1, nextOffset: 6, rows: rows(1) });
    expect(limitResponseSize(small)).toBe(small);
  });
});
