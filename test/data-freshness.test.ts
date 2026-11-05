import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer, requireDatabase, withDatabaseInfo } from "../src/mcp/server.js";
import { addDays, localToday } from "../src/odata/settlements-aging.js";
import { lastClosedMonths } from "../src/tools/freshness.js";
import { KEY_CLOSING_OPERATION, type ClosingOp } from "../src/tools/month-close.js";
import { fake1C } from "./helpers/fake-1c.js";

const SALES = "Document_РеализацияТоваровУслуг";
const BANK = "Document_ПоступлениеНаРасчетныйСчет";
const today = localToday();
const iso = (d: string) => `${d}T12:00:00`;

type Tool = {
  inputSchema: { parse: (a: object) => object };
  handler: (a: object, e: object) => Promise<CallToolResult>;
};
const toolsOf = (ctx: ReturnType<typeof fake1C>["ctx"]) =>
  (createServer(ctx) as unknown as { _registeredTools: Record<string, Tool> })._registeredTools;

interface Row {
  Date: string;
  Posted?: boolean;
  Number?: string;
}
/** Минимальный ответ 1С по документам: понимает Date le, Date ge и сортировку по убыванию. */
function replyWith(f: ReturnType<typeof fake1C>, docs: Record<string, Row[]>) {
  return vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string) => {
    const u = decodeURIComponent(url);
    const set = Object.keys(docs).find((s) => u.startsWith(s));
    let rows = set ? docs[set]! : [];
    const le = /Date le datetime'([^']+)'/.exec(u)?.[1];
    const ge = /Date ge datetime'([^']+)'/.exec(u)?.[1];
    if (le) rows = rows.filter((r) => r.Date <= le);
    if (ge) rows = rows.filter((r) => r.Date >= ge);
    const sorted = [...rows].sort((a, b) => b.Date.localeCompare(a.Date));
    const top = Number(/\$top=(\d+)/.exec(u)?.[1] ?? sorted.length);
    return { value: sorted.slice(0, top) as never[] };
  });
}
async function freshness(f: ReturnType<typeof fake1C>) {
  const tool = toolsOf(f.ctx)["read.system.data_freshness"]!;
  const r = await tool.handler(tool.inputSchema.parse({}), {});
  return r.structuredContent as {
    asOf: string;
    documents: Array<{
      entitySet: string;
      lastDate: string | null;
      lastNumber?: string;
      countInWindow: number;
      unpostedInWindow: number;
      futureDatedUntil?: string;
    }>;
  };
}

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("read.system.data_freshness", () => {
  it("последний документ, число за окно и непроведённые", async () => {
    const f = fake1C();
    replyWith(f, {
      [SALES]: [
        { Date: iso(addDays(today, -2)), Posted: true, Number: "0000-000005 " },
        { Date: iso(addDays(today, -5)), Posted: false, Number: "0000-000004" },
        { Date: iso(addDays(today, -90)), Posted: true, Number: "0000-000001" },
      ],
    });
    const { documents, asOf } = await freshness(f);
    expect(asOf).toBe(today);
    expect(documents.find((d) => d.entitySet === SALES)).toMatchObject({
      lastDate: addDays(today, -2),
      lastNumber: "0000-000005",
      countInWindow: 2,
      unpostedInWindow: 1,
    });
  });

  it("будущая дата не выдаётся за последний документ", async () => {
    const f = fake1C();
    replyWith(f, {
      [SALES]: [
        { Date: iso(addDays(today, 25)), Posted: true },
        { Date: iso(addDays(today, -1)), Posted: true },
      ],
    });
    const sales = (await freshness(f)).documents.find((d) => d.entitySet === SALES)!;
    expect(sales.lastDate).toBe(addDays(today, -1));
    expect(sales.futureDatedUntil).toBe(addDays(today, 25));
    expect(sales.countInWindow).toBe(1);
  });

  it("только будущие документы: последней даты нет, будущая показана", async () => {
    const f = fake1C();
    replyWith(f, { [SALES]: [{ Date: iso(addDays(today, 3)), Posted: true }] });
    const sales = (await freshness(f)).documents.find((d) => d.entitySet === SALES)!;
    expect(sales).toMatchObject({ lastDate: null, countInWindow: 0, futureDatedUntil: addDays(today, 3) });
  });

  it("вид без свежих документов — без лишнего запроса за окно", async () => {
    const f = fake1C();
    const spy = replyWith(f, { [BANK]: [{ Date: iso("2017-06-27"), Posted: true }] });
    const bank = (await freshness(f)).documents.find((d) => d.entitySet === BANK)!;
    expect(bank).toMatchObject({ lastDate: "2017-06-27", countInWindow: 0, unpostedInWindow: 0 });
    const windows = spy.mock.calls.filter(
      ([u]) => decodeURIComponent(String(u)).startsWith(BANK) && String(u).includes("Date%20ge"),
    );
    expect(windows).toHaveLength(0);
  });

  it("пустая база: нули и null, без ошибки", async () => {
    const f = fake1C();
    replyWith(f, {});
    const { documents } = await freshness(f);
    expect(documents.length).toBeGreaterThan(5);
    for (const d of documents) expect(d).toMatchObject({ lastDate: null, countInWindow: 0 });
  });

  it("ошибка одного вида не теряет остальные", async () => {
    const f = fake1C();
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string) => {
      if (decodeURIComponent(url).startsWith(BANK)) throw new Error("сеть");
      return { value: [] };
    });
    const tool = toolsOf(f.ctx)["read.system.data_freshness"]!;
    const r = await tool.handler(tool.inputSchema.parse({}), {});
    const c = r.structuredContent as { documents: unknown[]; failed: Array<{ entitySet: string }> };
    expect(r.isError).toBeFalsy();
    expect(c.failed.map((x) => x.entitySet)).toEqual([BANK]);
    expect(c.documents.length).toBeGreaterThan(5);
  });
});

describe("data_freshness: окно регламентных операций", () => {
  it("операций больше окна — об этом сказано, и список закрытий остаётся", async () => {
    const f = fake1C();
    const realGet = f.conn.client.getCollection.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (url: string, ...rest: never[]) => {
      const u = decodeURIComponent(url);
      if (u.startsWith("Document_РегламентнаяОперация")) {
        return {
          value: Array.from({ length: 300 }, () => ({
            Date: "2026-09-30T23:59:59",
            ВидОперации: "РасчетНалогаУСН",
            Состояние: "Выполнено",
          })) as never[],
        };
      }
      return u.startsWith("Document_") ? { value: [] } : realGet(url, ...rest);
    });
    const tool = toolsOf(f.ctx)["read.system.data_freshness"]!;
    const r = await tool.handler(tool.inputSchema.parse({}), {});
    expect((r.structuredContent as { closingNote: string }).closingNote).toContain("300");
  });
});

describe("lastClosedMonths: те же правила, что check_month_close", () => {
  const op = (month: string, kind: string, state: string, org = "A"): ClosingOp => ({
    month,
    org,
    kind,
    state,
  });
  it("берёт последний закрытый месяц, а не последний с операциями", () => {
    const m = lastClosedMonths([
      op("2026-08", KEY_CLOSING_OPERATION, "Выполнено"),
      op("2026-09", KEY_CLOSING_OPERATION, "Выполнено"),
      op("2026-10", "РасчетНалогаУСН", "Выполнено"),
    ]);
    expect(m.get("A")).toBe("2026-09");
  });
  it("месяц с невыполненной операцией не считается закрытым", () => {
    const m = lastClosedMonths([
      op("2026-08", KEY_CLOSING_OPERATION, "Выполнено"),
      op("2026-09", KEY_CLOSING_OPERATION, "Выполнено"),
      op("2026-09", "РасчетНалогаУСН", "НеВыполнено"),
    ]);
    expect(m.get("A")).toBe("2026-08");
  });
  it("у каждой организации своё закрытие; нет закрытых — нет записи", () => {
    const m = lastClosedMonths([
      op("2026-08", KEY_CLOSING_OPERATION, "Выполнено", "A"),
      op("2026-09", KEY_CLOSING_OPERATION, "Выполнено", "B"),
      op("2026-09", "РасчетНалогаУСН", "Выполнено", "C"),
    ]);
    expect([...m]).toEqual([
      ["A", "2026-08"],
      ["B", "2026-09"],
    ]);
  });
});

describe("withDatabaseInfo: какая база ответила", () => {
  const conn = { cfg: { name: "ip" } };
  const ctx = (names: string[]) =>
    ({ db: () => conn, databases: () => names.map((name) => ({ name })) }) as never;
  const result = (data: object): CallToolResult => ({
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
    structuredContent: data as Record<string, unknown>,
  });

  it("база подставлена при нескольких настроенных — пометка и другие базы, в тексте тоже", () => {
    const r = withDatabaseInfo(ctx(["ip", "ooo"]), {}, result({ rows: [] }));
    expect(r.structuredContent).toMatchObject({ database: "ip", databaseDefaulted: true });
    const text = JSON.parse((r.content[0] as { text: string }).text) as { databaseNote: string };
    expect(text.databaseNote).toContain("ooo");
  });
  it("база указана явно — без пометки", () => {
    const r = withDatabaseInfo(ctx(["ip", "ooo"]), { database: "ip" }, result({ rows: [] }));
    expect(r.structuredContent).toEqual({ database: "ip", rows: [] });
  });
  it("одна база настроена — без пометки", () => {
    const r = withDatabaseInfo(ctx(["ip"]), {}, result({ rows: [] }));
    expect(r.structuredContent).toEqual({ database: "ip", rows: [] });
  });
  it("ошибка и ответ с собственным текстом не меняются", () => {
    const err: CallToolResult = { content: [{ type: "text", text: "нет" }], isError: true };
    expect(withDatabaseInfo(ctx(["ip", "ooo"]), {}, err)).toBe(err);
    const custom: CallToolResult = {
      content: [{ type: "text", text: "сводка" }],
      structuredContent: { rows: [] },
    };
    const r = withDatabaseInfo(ctx(["ip", "ooo"]), {}, custom);
    expect((r.content[0] as { text: string }).text).toBe("сводка");
  });

  it.each(["read.accounting.get_contract_settlement_settings", "read.accounting.get_settlement_settings"])(
    "%s называет базу",
    async (name) => {
      const f = fake1C();
      const id = "11111111-1111-4111-8111-111111111111";
      const tool = toolsOf(f.ctx)[name]!;
      const r = await tool.handler(
        tool.inputSchema.parse({
          organizationRef: id,
          counterpartyRef: id,
          contractRef: id,
          settlementKind: "РасчетыВВалютеРегламентированногоУчета",
          scope: "contract",
        }),
        {},
      );
      expect(r.isError).toBeFalsy();
      expect((r.structuredContent as { database: string }).database).toBeTruthy();
    },
  );

  it("каждый инструмент чтения называет базу (кроме list_databases)", async () => {
    const f = fake1C();
    const tools = toolsOf(f.ctx);
    const sample = tools["read.document.search_documents"]!;
    const r = await sample.handler(
      sample.inputSchema.parse({ entitySet: SALES, from: "2026-01-01", to: "2026-01-31" }),
      {},
    );
    expect((r.structuredContent as { database: string }).database).toBeTruthy();
    const list = tools["read.system.list_databases"]!;
    const l = await list.handler({}, {});
    expect((l.structuredContent as { database?: string }).database).toBeUndefined();
  });
});

describe("запись при нескольких базах требует database", () => {
  const multi = (f: ReturnType<typeof fake1C>) =>
    new Proxy(f.ctx, {
      get: (t, k) =>
        k === "databases"
          ? () => [
              { name: "ip", isDefault: true },
              { name: "ooo", isDefault: false },
            ]
          : k === "db"
            ? () => f.conn
            : Reflect.get(t, k),
    }) as typeof f.ctx;

  it("отказ без обращения к инструменту; с базой и с одной базой проходит", async () => {
    const inner = vi.fn(async () => ({ content: [] }));
    const two = { databases: () => [{ name: "a" }, { name: "b" }] } as never;
    const one = { databases: () => [{ name: "a" }] } as never;
    const refused = await requireDatabase(two, inner)({}, {});
    expect(refused.isError).toBe(true);
    expect((refused.content[0] as { text: string }).text).toContain("a, b");
    expect(inner).not.toHaveBeenCalled();
    await requireDatabase(two, inner)({ database: "a" }, {});
    await requireDatabase(one, inner)({}, {});
    expect(inner).toHaveBeenCalledTimes(2);
  });

  it("каждый инструмент записи: отказ без database, ни одного запроса; и есть параметр database", async () => {
    const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const tools = toolsOf(multi(f));
    const spy = vi.spyOn(f.conn.client, "getCollection");
    const names = Object.keys(tools).filter((n) => n.startsWith("write."));
    expect(names.length).toBeGreaterThan(30);
    for (const name of names) {
      const tool = tools[name] as Tool & { inputSchema: { shape: Record<string, unknown> } };
      expect(Object.keys(tool.inputSchema.shape), name).toContain("database");
      for (const confirm of [false, true]) {
        const r = await tool.handler({ confirm }, {});
        expect(r.isError, `${name} confirm=${confirm}`).toBe(true);
        expect((r.content[0] as { text: string }).text, name).toContain("укажите database");
      }
    }
    expect(spy).not.toHaveBeenCalled();
  });

  it("чтение при нескольких базах без database работает", async () => {
    const f = fake1C();
    const tool = toolsOf(multi(f))["read.system.list_organizations"]!;
    const r = await tool.handler(tool.inputSchema.parse({}), {});
    expect(r.isError).toBeFalsy();
    expect(r.structuredContent).toMatchObject({ databaseDefaulted: true });
  });
});
