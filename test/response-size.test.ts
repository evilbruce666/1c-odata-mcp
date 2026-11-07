import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer, limitResponseSize } from "../src/mcp/server.js";
import { bySplittingAccounts, resolveAccounts } from "../src/odata/accounting.js";
import { fetchMovements } from "../src/odata/account-card.js";
import { ODataError } from "../src/odata/errors.js";
import { fitToChars, pageOf } from "../src/tools/_shared.js";
import { fake1C } from "./helpers/fake-1c.js";

const NORM = 50_000;
type Tool = {
  inputSchema: { parse: (a: object) => object };
  handler: (a: object, e: object) => Promise<CallToolResult>;
};
const toolOf = (f: ReturnType<typeof fake1C>, name: string) =>
  (createServer(f.ctx) as unknown as { _registeredTools: Record<string, Tool> })._registeredTools[name]!;
interface Paged {
  returned: number;
  truncated: boolean;
  nextOffset?: number;
  offset: number;
  note?: string;
}
/** Читает все страницы инструмента; на каждой проверяет норму размера. */
async function readAll<T>(tool: Tool, args: object, take: (s: never) => T[]) {
  const out: T[] = [];
  let offset = 0;
  let first: (Paged & Record<string, unknown>) | undefined;
  for (let i = 0; i < 100; i++) {
    const r = await tool.handler(tool.inputSchema.parse({ ...args, offset }), {});
    expect(r.isError).toBeFalsy();
    expect((r.content[0] as { text: string }).text.length).toBeLessThanOrEqual(NORM);
    const s = r.structuredContent as Paged & Record<string, unknown>;
    first ??= s;
    const page = take(s as never);
    expect(s.returned).toBe(page.length);
    out.push(...page);
    if (!s.truncated) {
      expect(s.nextOffset).toBeUndefined();
      return { out, first };
    }
    expect(s.note).toContain(`offset=${s.nextOffset}`);
    offset = s.nextOffset!;
  }
  throw new Error("страницы не кончились");
}

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("pageOf", () => {
  const items = Array.from({ length: 10 }, (_, i) => i);
  it("limit и продолжение", () => {
    const p = pageOf(items, { offset: 2, limit: 3, head: {} });
    expect(p.items).toEqual([2, 3, 4]);
    expect(p.page).toEqual({ offset: 2, returned: 3, truncated: true, nextOffset: 5 });
    expect(p.note).toContain("offset=5");
  });
  it("последняя страница и пусто", () => {
    expect(pageOf(items, { offset: 8, head: {} }).page).toEqual({ offset: 8, returned: 2, truncated: false });
    expect(pageOf([], { offset: 0, head: {} })).toMatchObject({ items: [], page: { truncated: false } });
  });
  it("offset за краем — объяснение", () => {
    const p = pageOf(items, { offset: 99, head: {} });
    expect(p.items).toEqual([]);
    expect(p.note).toContain("за пределами");
  });
  it("лимит выборки: читать дальше нечего, но усечение названо", () => {
    const p = pageOf(items, { offset: 0, head: {}, capHit: true });
    expect(p.page).toEqual({ offset: 0, returned: 10, truncated: true });
    expect(p.note).toContain("лимит");
  });
  it("размер: строка больше предела всё равно отдаётся, offset сдвигается", () => {
    expect(fitToChars([{ t: "я".repeat(100_000) }], 0)).toHaveLength(1);
  });
});

describe("большие ответы укладываются в норму, ничего не теряется", () => {
  it("list_entities: тысячи объектов", async () => {
    const f = fake1C();
    const meta = await f.conn.getMetadata();
    const entities = new Map(
      Array.from({ length: 3000 }, (_, i) => [
        `Catalog_Объект${i}`,
        { entitySet: `Catalog_Объект${i}`, shortName: `Объект${i}`, class: i % 2 ? "catalog" : "document" },
      ]),
    );
    vi.spyOn(f.conn, "getMetadata").mockResolvedValue({ ...meta, entities } as never);
    const { out, first } = await readAll(
      toolOf(f, "read.schema.list_entities"),
      {},
      (s: { groups: Record<string, Array<{ entitySet: string }>> }) => Object.values(s.groups).flat(),
    );
    expect(out).toHaveLength(3000);
    expect(new Set(out.map((e) => e.entitySet)).size).toBe(3000);
    expect(first).toMatchObject({ total: 3000, truncated: true });
    expect(Object.values(first!["countsByClass"] as Record<string, number>).reduce((a, b) => a + b)).toBe(
      3000,
    );
  });

  it("list_entities: класс отбирается, страница группируется", async () => {
    const f = fake1C();
    const tool = toolOf(f, "read.schema.list_entities");
    const r = await tool.handler(tool.inputSchema.parse({ class: "catalog" }), {});
    const s = r.structuredContent as { groups: Record<string, unknown>; truncated: boolean };
    expect(Object.keys(s.groups)).toEqual(["Справочники"]);
    expect(s.truncated).toBe(false);
  });

  it("search_documents: limit 500, страницы без потерь", async () => {
    const f = fake1C();
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    const docs = Array.from({ length: 700 }, (_, i) => ({
      Ref_Key: `d${i}`,
      Number: `0000-${String(i).padStart(6, "0")}`,
      Date: `2026-0${(i % 9) + 1}-15T00:00:00`,
      Posted: true,
      DeletionMark: false,
      СуммаДокумента: i,
      НомерВходящегоДокумента: "описание длинного входящего номера документа для объёма",
    }));
    // Как настоящая 1С: заказанная сортировка, поэтому первые N строк всегда одни и те же.
    docs.sort((a, b) => b.Date.localeCompare(a.Date) || a.Ref_Key.localeCompare(b.Ref_Key));
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string, ...rest: never[]) => {
      const u = decodeURIComponent(url);
      if (!u.startsWith("Document_СписаниеСРасчетногоСчета")) return realGet(url, ...rest);
      const top = Number(/\$top=(\d+)/.exec(u)![1]);
      const skip = Number(/\$skip=(\d+)/.exec(u)?.[1] ?? 0);
      return { value: docs.slice(skip, skip + top) as never[] };
    });
    const { out, first } = await readAll(
      toolOf(f, "read.document.search_documents"),
      { entitySet: "Document_СписаниеСРасчетногоСчета", limit: 500 },
      (s: { rows: Array<{ ref: string }> }) => s.rows,
    );
    expect(first!.truncated).toBe(true);
    expect(new Set(out.map((r) => r.ref)).size).toBe(out.length);
    expect(out).toHaveLength(700);
  });

  it("search_documents: limit соблюдается и показывает продолжение", async () => {
    const f = fake1C();
    const docs = Array.from({ length: 5 }, (_, i) => ({ Ref_Key: `d${i}`, Date: "2026-01-01T00:00:00" }));
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string) => ({
      value: docs.slice(0, Number(/\$top=(\d+)/.exec(decodeURIComponent(url))![1])) as never[],
    }));
    const tool = toolOf(f, "read.document.search_documents");
    const r = await tool.handler(
      tool.inputSchema.parse({ entitySet: "Document_СписаниеСРасчетногоСчета", limit: 2 }),
      {},
    );
    expect(r.structuredContent).toMatchObject({ count: 2, returned: 2, truncated: true, nextOffset: 2 });
    const all = await tool.handler(
      tool.inputSchema.parse({ entitySet: "Document_СписаниеСРасчетногоСчета", limit: 5 }),
      {},
    );
    expect(all.structuredContent).toMatchObject({ count: 5, truncated: false });
  });

  it("get_deal_history: итоги по всем, страницы без потерь", async () => {
    const f = fake1C();
    const docs = Array.from({ length: 600 }, (_, i) => ({
      Ref_Key: `d${i}`,
      Number: `0000-${i}`,
      Date: "2026-03-01T00:00:00",
      Posted: true,
      СуммаДокумента: 10,
      ВидОперации: "ОплатаПокупателя",
      НазначениеПлатежа: "договор поставки ".repeat(12),
    }));
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string, ...rest: never[]) => {
      if (!decodeURIComponent(url).startsWith("Document_ПоступлениеНаРасчетныйСчет")) {
        return /^Document_/.test(decodeURIComponent(url)) ? { value: [] } : realGet(url, ...rest);
      }
      const u = decodeURIComponent(url);
      const top = Number(/\$top=(\d+)/.exec(u)![1]);
      const skip = Number(/\$skip=(\d+)/.exec(u)?.[1] ?? 0);
      return { value: docs.slice(skip, skip + top) as never[] };
    });
    const { out, first } = await readAll(
      toolOf(f, "read.analytics.get_deal_history"),
      { dealKey: "договор", limit: 1000 },
      (s: { items: Array<{ number: string }> }) => s.items,
    );
    expect(first).toMatchObject({ events: 600, inflow: 6000, truncated: true });
    expect(out).toHaveLength(600);
  });

  it("get_reconciliation: сальдо и обороты по всему периоду, строки страницами", async () => {
    const f = fake1C();
    const acc = (await resolveAccounts(f.conn, ["62.01"]))[0]!.key;
    const CP = "11111111-1111-4111-8111-111111111111";
    const rows = Array.from({ length: 800 }, (_, i) => ({
      Period: "2026-03-15T00:00:00",
      Recorder: `0000000${i % 9}-0000-4000-8000-${String(i).padStart(12, "0")}`,
      Recorder_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
      LineNumber: "1",
      AccountDr_Key: acc,
      ExtDimensionDr1: CP,
      ExtDimensionDr1_Type: "StandardODATA.Catalog_Контрагенты",
      Сумма: 1,
    }));
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string, ...rest: never[]) => {
      const u = decodeURIComponent(url);
      if (u.includes("RecordsWithExtDimensions")) return { value: rows as never[] };
      if (u.includes("/Balance")) return { value: [] };
      return realGet(url, ...rest);
    });
    const { out, first } = await readAll(
      toolOf(f, "read.counterparty.get_reconciliation"),
      { counterpartyRef: CP, from: "2026-01-01", to: "2026-12-31" },
      (s: { lines: Array<{ documentRef: string }> }) => s.lines,
    );
    expect(first).toMatchObject({ count: 800, debitTurnover: 800, truncated: true });
    expect(new Set(out.map((l) => l.documentRef)).size).toBe(800);
  });
});

describe("limitResponseSize: страницы для инструментов без своих", () => {
  const answer = (data: Record<string, unknown>): CallToolResult => ({
    content: [{ type: "text", text: JSON.stringify(data) }],
    structuredContent: data,
  });
  const rows = Array.from({ length: 2000 }, (_, i) => ({ n: i, text: "строка ".repeat(10) }));
  type Cut = {
    total: number;
    small: number[];
    rows: Array<{ n: number }>;
    note: string;
    listField: string;
    listTotal: number;
    returned: number;
    truncated: boolean;
    nextOffset?: number;
  };

  it("режет самый большой список, называет это, итоги и другие поля остаются", () => {
    const r = limitResponseSize(answer({ total: 2000, small: [1, 2], rows, note: "итоги верны" }), {
      continuable: true,
    });
    const text = (r.content[0] as { text: string }).text;
    expect(text.length).toBeLessThanOrEqual(NORM);
    const s = r.structuredContent as unknown as Cut;
    expect(s.total).toBe(2000);
    expect(s.small).toEqual([1, 2]);
    expect(s).toMatchObject({
      listField: "rows",
      listTotal: 2000,
      truncated: true,
      nextOffset: s.rows.length,
    });
    expect(s.note).toContain("итоги верны");
    expect(s.note).toContain(`offset=${s.nextOffset}`);
    expect(JSON.parse(text)).toEqual(s);
  });

  it("все страницы вместе — каждая строка ровно один раз", () => {
    const seen: number[] = [];
    let offset = 0;
    for (let i = 0; i < 50; i++) {
      const s = limitResponseSize(answer({ rows }), { offset, continuable: true })
        .structuredContent as unknown as Cut;
      seen.push(...s.rows.map((x) => x.n));
      if (!s.truncated) break;
      offset = s.nextOffset!;
    }
    expect(seen).toEqual(rows.map((x) => x.n));
  });

  it("инструмент без offset: усечение названо, продолжения не обещает", () => {
    const s = limitResponseSize(answer({ rows })).structuredContent as unknown as Cut;
    expect(s.truncated).toBe(true);
    expect(s.nextOffset).toBeUndefined();
    expect(s.note).toContain("Уточните");
  });

  it("в пределах нормы и ошибки не меняются", () => {
    const ok = answer({ rows: rows.slice(0, 3) });
    expect(limitResponseSize(ok)).toBe(ok);
    const err: CallToolResult = { content: [{ type: "text", text: "x".repeat(60_000) }], isError: true };
    expect(limitResponseSize(err)).toBe(err);
  });

  it("без списка ответ не трогает", () => {
    const big = answer({ text: "x".repeat(60_000) });
    expect(limitResponseSize(big)).toBe(big);
  });
});

describe("offset у инструментов со списком по limit", () => {
  const tools = () => {
    const f = fake1C();
    return {
      f,
      all: (
        createServer(f.ctx) as unknown as {
          _registeredTools: Record<string, Tool & { inputSchema: { shape: Record<string, unknown> } }>;
        }
      )._registeredTools,
    };
  };
  it("есть у всех инструментов чтения с limit, запись не затронута", () => {
    const { all } = tools();
    for (const [name, t] of Object.entries(all)) {
      if (name.startsWith("read.") && "limit" in t.inputSchema.shape) {
        expect(Object.keys(t.inputSchema.shape), name).toContain("offset");
      }
    }
    expect(Object.keys(all["read.system.health_check"]!.inputSchema.shape)).not.toContain("offset");
  });

  it("find_counterparty: вторая страница читает limit+offset строк и отдаёт продолжение", async () => {
    const { f, all } = tools();
    const rows = Array.from({ length: 30 }, (_, i) => ({
      Ref_Key: `cccccccc-0000-4000-8000-${String(i).padStart(12, "0")}`,
      Description: `Контрагент ${i}`,
      IsFolder: false,
    }));
    const spy = vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string) => {
      const u = decodeURIComponent(url);
      const top = Number(/\$top=(\d+)/.exec(u)?.[1] ?? 1000);
      const skip = Number(/\$skip=(\d+)/.exec(u)?.[1] ?? 0);
      return { value: (u.startsWith("Catalog_Контрагенты") ? rows.slice(skip, skip + top) : []) as never[] };
    });
    const tool = all["read.counterparty.find_counterparty"]!;
    const r = await tool.handler(tool.inputSchema.parse({ query: "Контрагент", limit: 10, offset: 10 }), {});
    if (r.isError) throw new Error(JSON.stringify(r.content));
    const s = r.structuredContent as { rows: Array<{ ref: string }>; offset: number; returned: number };
    expect(s).toMatchObject({ offset: 10, returned: 10 });
    expect(s.rows[0]!.ref).toBe("cccccccc-0000-4000-8000-000000000010");
    expect(r.structuredContent).toMatchObject({ count: 10, listTotal: 20 });
    expect(decodeURIComponent(String(spy.mock.calls[0]![0]))).toContain("Контрагент");
  });
});

describe("snapshot: данные изменились между страницами", () => {
  const items = Array.from({ length: 5 }, (_, i) => i);
  it("те же итоги — метка та же, без dataChanged", () => {
    const a = pageOf(items, { offset: 0, limit: 2, head: { total: 5 }, tracksChanges: true });
    const b = pageOf(items, {
      offset: 2,
      limit: 2,
      head: { total: 5 },
      tracksChanges: true,
      snapshot: a.page["snapshot"] as string,
    });
    expect(b.page["snapshot"]).toBe(a.page["snapshot"]);
    expect(b.page["dataChanged"]).toBeUndefined();
  });
  it("итоги изменились — dataChanged и совет читать заново", () => {
    const a = pageOf(items, { offset: 0, limit: 2, head: { total: 5 }, tracksChanges: true });
    const b = pageOf(items, {
      offset: 2,
      limit: 2,
      head: { total: 6 },
      tracksChanges: true,
      snapshot: a.page["snapshot"] as string,
    });
    expect(b.page["dataChanged"]).toBe(true);
    expect(b.note).toContain("offset=0");
  });
  it("search_documents не ставит метку: его head меняется от страницы к странице", () => {
    const p = pageOf(items, { offset: 0, limit: 2, head: { searched: 1 } });
    expect(p.page["snapshot"]).toBeUndefined();
  });
});

describe("таймаут большой выборки: разбиение вместо отказа", () => {
  const ACC = "aaaaaaaa-0000-4000-8000-000000000001";
  const timeout = () => new ODataError({ kind: "timeout", message: "Превышен таймаут запроса" });
  const row = (doc: string, acc: string) => ({
    Period: "2026-01-10T00:00:00",
    Recorder: doc,
    Recorder_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
    LineNumber: "1",
    AccountDr_Key: acc,
    Сумма: 1,
  });

  it("движения: период делится по месяцам, строки складываются", async () => {
    const { conn } = fake1C();
    const acc = ACC;
    const spy = vi.spyOn(conn.client, "getCollection").mockImplementation(async (url: string) => {
      const u = decodeURIComponent(url);
      const whole =
        u.includes("StartPeriod=datetime'2026-01-01") && u.includes("EndPeriod=datetime'2026-03-31");
      if (whole) throw timeout();
      const m = /StartPeriod=datetime'(\d{4}-\d{2})/.exec(u)![1];
      return { value: [row(`doc-${m}`, acc)] as never[] };
    });
    const { rows } = await fetchMovements(conn, { accountKeys: [acc], from: "2026-01-01", to: "2026-03-31" });
    expect(rows.map((r) => r["Recorder"])).toEqual(["doc-2026-01", "doc-2026-02", "doc-2026-03"]);
    expect(spy).toHaveBeenCalledTimes(4);
  });

  it("движения: один месяц не успел — ошибка таймаута, а не пустой результат", async () => {
    const { conn } = fake1C();
    vi.spyOn(conn.client, "getCollection").mockRejectedValue(timeout());
    await expect(
      fetchMovements(conn, { accountKeys: [ACC], from: "2026-01-01", to: "2026-01-31" }),
    ).rejects.toThrow("таймаут");
  });

  it("другая ошибка не разбивается", async () => {
    const { conn } = fake1C();
    const spy = vi
      .spyOn(conn.client, "getCollection")
      .mockRejectedValue(new ODataError({ kind: "bad_request", message: "плохо" }));
    await expect(
      fetchMovements(conn, { accountKeys: [ACC], from: "2026-01-01", to: "2026-03-31" }),
    ).rejects.toThrow("плохо");
    expect(spy).toHaveBeenCalledTimes(1);
  });

  it("счета: не успели — читаются половинами, строки не задваиваются", async () => {
    const seen: string[][] = [];
    const r = await bySplittingAccounts(["a", "b", "c", "d"], async (keys) => {
      seen.push([...keys]);
      if (keys.length > 1) throw timeout();
      return { rows: [{ k: keys[0] }], meta: { rowsScanned: 1, chunks: 1 } };
    });
    expect(r.rows.map((x) => x["k"])).toEqual(["a", "b", "c", "d"]);
    expect(r.meta.rowsScanned).toBe(4);
    expect(seen[0]).toEqual(["a", "b", "c", "d"]);
  });

  it("счёт один и не успел — ошибка", async () => {
    await expect(bySplittingAccounts(["a"], async () => Promise.reject(timeout()))).rejects.toThrow(
      "таймаут",
    );
  });
});
