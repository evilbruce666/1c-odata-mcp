import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { fetchMovements, STABLE_MOVEMENT_ORDER, type CardEntry } from "../src/odata/account-card.js";
import { resolveAccounts } from "../src/odata/accounting.js";
import { createServer } from "../src/mcp/server.js";
import { fitToChars, groupDims, groupEntries, MAX_CARD_PAGE } from "../src/tools/card.js";
import { fake1C } from "./helpers/fake-1c.js";

const ACC = "aaaaaaaa-0000-4000-8000-000000000001";
const row = (doc: string, line = 1) => ({
  Period: "2026-01-10T00:00:00",
  Recorder: doc,
  Recorder_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
  LineNumber: String(line),
  Active: true,
  AccountDr_Key: ACC,
  Сумма: 100,
});
const args = { accountKeys: [ACC], from: "2026-01-01", to: "2026-01-31" };

function replyWith(conn: ReturnType<typeof fake1C>["conn"], reply: (sorted: boolean) => object[]) {
  return vi.spyOn(conn.client, "getCollection").mockImplementation(async (url: string) => ({
    value: reply(decodeURIComponent(url).includes("$orderby")) as never[],
  }));
}

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("fetchMovements: сбой листания 1С не превращается в тихий недосчёт", () => {
  it("без повторов — один запрос, без сортировки", async () => {
    const { conn } = fake1C();
    const spy = replyWith(conn, () => [row("a"), row("b")]);
    const { rows } = await fetchMovements(conn, args);
    expect(rows).toHaveLength(2);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(decodeURIComponent(String(spy.mock.calls[0]![0]))).not.toContain("$orderby");
  });

  it("повтор вместо пропущенной строки — перечитывает с устойчивой сортировкой", async () => {
    const { conn } = fake1C();
    const spy = replyWith(conn, (sorted) => (sorted ? [row("a"), row("b")] : [row("a"), row("a")]));
    const { rows, meta } = await fetchMovements(conn, args);
    expect(rows.map((r) => r["Recorder"])).toEqual(["a", "b"]);
    expect(spy).toHaveBeenCalledTimes(2);
    expect(decodeURIComponent(String(spy.mock.calls[1]![0]))).toContain(STABLE_MOVEMENT_ORDER);
    expect(meta.chunks).toBe(2);
  });

  it("повторы и с сортировкой — отказ, а не неполная сумма", async () => {
    const { conn } = fake1C();
    replyWith(conn, () => [row("a"), row("a")]);
    await expect(fetchMovements(conn, args)).rejects.toThrow("с повторами");
  });

  it("разные строки одного документа — не повтор", async () => {
    const { conn } = fake1C();
    const spy = replyWith(conn, () => [row("a", 1), row("a", 2)]);
    const { rows } = await fetchMovements(conn, args);
    expect(rows).toHaveLength(2);
    expect(spy).toHaveBeenCalledTimes(1);
  });
});

describe("get_account_card: страницы, предел размера, группы", () => {
  const OURS = "aaaaaaaa-0000-4000-8000-000000000001";
  const CP1 = "11111111-1111-4111-8111-111111111111";
  const CP2 = "22222222-2222-4222-8222-222222222222";
  const CONTRACT = "33333333-3333-4333-8333-333333333333";
  const CORR = "bbbbbbbb-0000-4000-8000-000000000002";
  const entry = (i: number, cp = CP1, month = "01"): CardEntry => ({
    period: `2026-${month}-${String((i % 28) + 1).padStart(2, "0")}T00:00:00`,
    recorder: `doc-${i}`,
    side: i % 3 === 0 ? "credit" : "debit",
    cents: 100 * (i + 1),
    account: OURS,
    dims: [
      { ref: cp, set: "Catalog_Контрагенты" },
      { ref: CONTRACT, set: "Catalog_ДоговорыКонтрагентов" },
    ],
    corrAccount: CORR,
    corrDims: [],
  });
  const code = (k?: string) => (k === CORR ? "51" : k);
  const name = (d: { ref: string }) =>
    ({ [CP1]: "Ромашка", [CP2]: "Ёлка", [CONTRACT]: "Договор №1" })[d.ref] ?? d.ref;

  it("fitToChars: режет по размеру, но хотя бы одну строку отдаёт", () => {
    const items = Array.from({ length: 50 }, (_, i) => ({ n: i, text: "я".repeat(1000) }));
    const fit = fitToChars(items, 500, 10_000);
    expect(fit.length).toBeGreaterThan(5);
    expect(fit.length).toBeLessThan(11);
    expect(JSON.stringify({ entries: fit }, null, 2).length).toBeLessThanOrEqual(10_000);
    expect(fitToChars(items, 999_999, 10_000)).toHaveLength(1);
    expect(fitToChars([], 0)).toEqual([]);
  });

  it("groupEntries: суммы групп равны оборотам, сторно и кириллица не ломают", () => {
    const entries = [
      entry(1, CP1, "01"),
      entry(2, CP1, "02"),
      entry(3, CP2, "02"),
      { ...entry(4, CP2, "02"), cents: -500 },
    ];
    const groups = groupEntries(entries, ["month", "counterparty"], name, code);
    expect(groups.map((g) => [g["month"], g["counterparty"]])).toEqual([
      ["2026-01", "Ромашка"],
      ["2026-02", "Ёлка"],
      ["2026-02", "Ромашка"],
    ]);
    const sum = (f: "debit" | "credit") => groups.reduce((s, g) => s + Number(g[f]), 0);
    const expected = (side: string) =>
      entries.filter((e) => e.side === side).reduce((s, e) => s + e.cents, 0) / 100;
    expect(sum("debit")).toBeCloseTo(expected("debit"), 2);
    expect(sum("credit")).toBeCloseTo(expected("credit"), 2);
    expect(groups.reduce((s, g) => s + Number(g["count"]), 0)).toBe(entries.length);
  });

  it("groupEntries: строка без субконто остаётся в сводке, без поля группы", () => {
    const groups = groupEntries([{ ...entry(1), dims: [] }], ["contract"], name, code);
    expect(groups).toEqual([{ debit: 2, credit: 0, count: 1 }]);
  });

  it("несколько договоров одного контрагента — разные группы", () => {
    const other = "44444444-4444-4444-8444-444444444444";
    const withOther = {
      ...entry(1),
      dims: [{ ref: other, set: "Catalog_ДоговорыКонтрагентов" }],
    };
    const groups = groupEntries([entry(2), withOther], ["contract"], name, code);
    expect(groups).toHaveLength(2);
  });

  it("одноимённые договоры разных контрагентов не склеиваются, в группе есть ссылка", () => {
    const a = "66666666-6666-4666-8666-666666666666";
    const b = "77777777-7777-4777-8777-777777777777";
    const contract = (ref: string, i: number) => ({
      ...entry(i),
      dims: [{ ref, set: "Catalog_ДоговорыКонтрагентов" }],
    });
    const groups = groupEntries(
      [contract(a, 1), contract(b, 2), contract(a, 4)],
      ["contract"],
      () => "Основной договор",
      code,
    );
    expect(groups.map((g) => [g["contract"], g["contractRef"], g["count"]])).toEqual([
      ["Основной договор", a, 2],
      ["Основной договор", b, 1],
    ]);
  });

  it("groupDims: без группировки по субконто наименования не запрашиваются", () => {
    expect(groupDims([entry(1)], ["month", "corrAccount"])).toEqual([]);
    expect(groupDims([entry(1)], ["contract"]).map((d) => d.ref)).toEqual([CONTRACT]);
  });

  // ─── Инструмент целиком, на поддельной 1С ──────────────────────────────────
  const REG_ROWS = 250;
  async function card(args: Record<string, unknown>, rowCount = REG_ROWS) {
    const f = fake1C();
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    const acc = (await resolveAccounts(f.conn, ["62.01"]))[0]!.key;
    const rows = Array.from({ length: rowCount }, (_, i) => ({
      Period: `2026-0${(i % 3) + 1}-15T00:00:00`,
      Recorder: `0000000${i % 10}-0000-4000-8000-${String(i).padStart(12, "0")}`,
      Recorder_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
      LineNumber: "1",
      Active: true,
      ...(i % 2 ? { AccountDr_Key: acc } : { AccountCr_Key: acc }),
      Сумма: i + 1,
      Содержание: "Реализация услуг покупателю с длинным описанием операции",
    }));
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string, ...rest: never[]) => {
      if (decodeURIComponent(url).includes("RecordsWithExtDimensions")) {
        const skip = Number(/\$skip=(\d+)/.exec(url)?.[1] ?? 0);
        const top = Number(/\$top=(\d+)/.exec(url)?.[1] ?? rows.length);
        return { value: rows.slice(skip, skip + top) as never[] };
      }
      return realGet(url, ...rest);
    });
    const tools = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<
          string,
          {
            inputSchema: { parse: (a: object) => object };
            handler: (a: object, e: object) => Promise<CallToolResult>;
          }
        >;
      }
    )._registeredTools;
    const tool = tools["read.accounting.get_account_card"]!;
    const r = await tool.handler(
      tool.inputSchema.parse({ account: "62.01", from: "2026-01-01", to: "2026-03-31", ...args }),
      {},
    );
    return {
      text: (r.content[0] as { text: string }).text,
      isError: r.isError,
      c: r.structuredContent as never as CardResult,
    };
  }
  interface CardResult {
    count: number;
    returned: number;
    truncated: boolean;
    nextOffset?: number;
    offset: number;
    debitTurnover: number;
    creditTurnover: number;
    note?: string;
    entries: Array<{ documentRef: string; debit?: number; credit?: number; balance: number }>;
    groups?: Array<{ month?: string; debit: number; credit: number; count: number }>;
  }

  it("все страницы вместе — каждое движение ровно один раз, суммы равны оборотам", async () => {
    const seen: string[] = [];
    let dr = 0;
    let cr = 0;
    let offset = 0;
    let first: CardResult | undefined;
    for (let page = 0; page < 20; page++) {
      const { c, text } = await card({ offset });
      first ??= c;
      expect(text.length).toBeLessThanOrEqual(50_000);
      expect(c.returned).toBe(c.entries.length);
      for (const e of c.entries) {
        seen.push(e.documentRef);
        dr += e.debit ?? 0;
        cr += e.credit ?? 0;
      }
      if (!c.truncated) {
        expect(c.nextOffset).toBeUndefined();
        break;
      }
      expect(c.note).toContain(`offset=${c.nextOffset}`);
      offset = c.nextOffset!;
    }
    expect(first!.truncated).toBe(true);
    expect(seen).toHaveLength(first!.count);
    expect(new Set(seen).size).toBe(first!.count);
    expect(dr).toBeCloseTo(first!.debitTurnover, 2);
    expect(cr).toBeCloseTo(first!.creditTurnover, 2);
  });

  it("большой limit не раздувает ответ", async () => {
    const { c, text } = await card({ limit: 2000 });
    expect(c.returned).toBeLessThanOrEqual(MAX_CARD_PAGE);
    expect(c.truncated).toBe(true);
    expect(text.length).toBeLessThanOrEqual(50_000);
  });

  it("limit меньше страницы соблюдается", async () => {
    const { c } = await card({ limit: 3 });
    expect(c.returned).toBe(3);
    expect(c.nextOffset).toBe(3);
  });

  it("остаток после движения на второй странице считается от начала периода", async () => {
    const whole = (await card({ limit: 20 })).c.entries[10]!;
    const paged = (await card({ limit: 5, offset: 10 })).c.entries[0]!;
    expect(paged).toEqual(whole);
  });

  it("summaryOnly и groupBy: без строк, группы сходятся с оборотами", async () => {
    const { c } = await card({ summaryOnly: true, groupBy: ["month"] });
    expect(c.entries).toEqual([]);
    expect(c.returned).toBe(0);
    expect(c.truncated).toBe(false);
    expect(c.groups!.map((g) => g.month)).toEqual(["2026-01", "2026-02", "2026-03"]);
    expect(c.groups!.reduce((s, g) => s + g.debit, 0)).toBeCloseTo(c.debitTurnover, 2);
    expect(c.groups!.reduce((s, g) => s + g.credit, 0)).toBeCloseTo(c.creditTurnover, 2);
    expect(c.groups!.reduce((s, g) => s + g.count, 0)).toBe(c.count);
  });

  it("groupBy со строками: группы по всему периоду, а не по странице", async () => {
    const { c } = await card({ limit: 2, groupBy: ["month"] });
    expect(c.returned).toBe(2);
    expect(c.groups!.reduce((s, g) => s + g.count, 0)).toBe(c.count);
  });

  it("пустой период: нули, без страниц и ошибок", async () => {
    const { c, isError } = await card({ groupBy: ["month"] }, 0);
    expect(isError).toBeFalsy();
    expect(c).toMatchObject({ count: 0, returned: 0, truncated: false, entries: [], groups: [] });
  });

  it("групп больше, чем помещается: отказ, а не обрезанная сводка", async () => {
    const f = fake1C();
    const acc = (await resolveAccounts(f.conn, ["62.01"]))[0]!.key;
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    const rows = Array.from({ length: 900 }, (_, i) => ({
      Period: "2026-01-15T00:00:00",
      Recorder: `doc-${i}`,
      LineNumber: "1",
      AccountDr_Key: acc,
      ExtDimensionDr1: `55555555-5555-4555-8555-${String(i).padStart(12, "0")}`,
      ExtDimensionDr1_Type: "StandardODATA.Catalog_ДоговорыКонтрагентов",
      Сумма: 1,
    }));
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string, ...rest: never[]) =>
      decodeURIComponent(url).includes("RecordsWithExtDimensions")
        ? { value: rows as never[] }
        : realGet(url, ...rest),
    );
    const tool = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools["read.accounting.get_account_card"]!;
    const r = await tool.handler(
      {
        account: "62.01",
        from: "2026-01-01",
        to: "2026-01-31",
        limit: 1,
        offset: 0,
        summaryOnly: true,
        groupBy: ["contract"],
      },
      {},
    );
    expect(r.isError).toBe(true);
    expect((r.content[0] as { text: string }).text).toContain("900 групп");
  });

  it("offset за пределами карточки: пусто и объяснение", async () => {
    const { c } = await card({ offset: 9999 });
    expect(c).toMatchObject({ returned: 0, truncated: false, entries: [] });
    expect(c.note).toContain("за пределами");
  });
});
