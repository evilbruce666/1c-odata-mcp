import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { matchesIncomingNumber, matchesOwnNumber, parseNumberQuery } from "../src/odata/doc-number.js";
import { fake1C } from "./helpers/fake-1c.js";

// Форматы — из настоящих баз БП 3.0: свой номер с префиксом и нулями разной длины, входящий — любой.
describe("номер документа", () => {
  const q = (s: string) => parseNumberQuery(s)!;

  it.each([
    ["33", "0000-000033"],
    ["№33", "0000-000033"],
    ["№ 0033", "0000-000033"],
    ["1", "00000000001"],
    ["24", "0000-00000024"],
    ["33", "БП-000033"],
    ["0000-000033", "0000-000033"],
    ["бп-000033", "БП-000033"],
  ])("«%s» находит свой номер %s", (query, value) => expect(matchesOwnNumber(value, q(query))).toBe(true));

  it.each([
    ["33", "0000-000133"],
    ["33", "0000-000330"],
    ["3", "0000-000033"],
    ["БП-33", "ОО-000033"],
  ])("«%s» не находит свой номер %s", (query, value) =>
    expect(matchesOwnNumber(value, q(query))).toBe(false),
  );

  it.each([
    ["1171733", "1171733 "],
    ["ГК-22537", "ГК-22537"],
    ["гк-22537", "ГК-22537"],
    ["20290930-29", "20290930-29"],
    ["5893", "05893"],
    ["0000-05893", "0000-05893"],
  ])("«%s» находит входящий номер %j", (query, value) =>
    expect(matchesIncomingNumber(value, q(query))).toBe(true),
  );

  it.each([
    ["29", "20290930-29"], // хвост входящего номера — не номер
    ["22537", "ГК-22537"],
    ["5893", "0000-05893"],
    ["33", "1171733"],
  ])("«%s» не находит входящий номер %j", (query, value) =>
    expect(matchesIncomingNumber(value, q(query))).toBe(false),
  );

  it("пустой запрос — undefined", () => {
    expect(parseNumberQuery("  № ")).toBeUndefined();
  });
});

function search(ctx: ReturnType<typeof fake1C>["ctx"]) {
  const tools = (
    createServer(ctx) as unknown as {
      _registeredTools: Record<string, { handler: (args: object, extra: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  return (args: object) =>
    tools["read.document.search_documents"]!.handler({ postedOnly: false, limit: 50, ...args }, {});
}
const urls = (spy: { mock: { calls: unknown[][] } }) =>
  spy.mock.calls.map((c) => decodeURIComponent(String(c[0])));

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("search_documents: номер, вид и поиск по основным документам", () => {
  it("номер в одном типе — отбор на сервере по своему и входящему номеру, только поля сводки", async () => {
    const f = fake1C();
    const spy = vi.spyOn(f.conn.client, "getCollection");
    const r = await search(f.ctx)({ entitySet: "Document_ПоступлениеТоваровУслуг", number: "№33" });
    expect(r.isError).toBeFalsy();
    const url = urls(spy)[0]!;
    expect(url).toContain("endswith(Number, '33')");
    expect(url).toContain("substringof('33', НомерВходящегоДокумента)");
    expect(url).toContain("$select=Ref_Key,Number,Date");
    expect(f.violations).toEqual([]);
  });

  it("клиентская проверка отсекает ложные совпадения подстроки", async () => {
    const f = fake1C();
    vi.spyOn(f.conn.client, "getCollection").mockResolvedValue({
      value: [
        { Ref_Key: "a", Number: "0000-000033", Date: "2026-09-10T00:00:00", НомерВходящегоДокумента: "" },
        { Ref_Key: "b", Number: "0000-000133", Date: "2026-09-09T00:00:00", НомерВходящегоДокумента: "" },
        { Ref_Key: "c", Number: "0000-000050", Date: "2025-01-01T00:00:00", НомерВходящегоДокумента: "33 " },
        { Ref_Key: "d", Number: "0000-000051", Date: "2025-01-02T00:00:00", НомерВходящегоДокумента: "1133" },
      ],
    });
    const r = await search(f.ctx)({ entitySet: "Document_ПоступлениеТоваровУслуг", number: "33" });
    const rows = (r.structuredContent as { rows: Array<{ ref: string; matchedBy: string }> }).rows;
    expect(rows.map((x) => [x.ref, x.matchedBy])).toEqual([
      ["a", "number"],
      ["c", "incomingNumber"],
    ]);
  });

  it("свой номер — выше входящего, даже если входящий новее", async () => {
    const f = fake1C();
    vi.spyOn(f.conn.client, "getCollection").mockResolvedValue({
      value: [
        { Ref_Key: "new", Number: "0000-000750", Date: "2026-09-04T00:00:00", НомерВходящегоДокумента: "33" },
        { Ref_Key: "own", Number: "0000-000033", Date: "2024-01-10T00:00:00", НомерВходящегоДокумента: "" },
      ],
    });
    const r = await search(f.ctx)({ entitySet: "Document_СписаниеСРасчетногоСчета", number: "33" });
    expect((r.structuredContent as { rows: Array<{ ref: string }> }).rows.map((x) => x.ref)).toEqual([
      "own",
      "new",
    ]);
  });

  it("без entitySet и kind — по всем основным видам, слитно и по дате", async () => {
    const f = fake1C();
    const spy = vi.spyOn(f.conn.client, "getCollection");
    const r = await search(f.ctx)({ from: "2026-09-10", to: "2026-09-10" });
    expect(r.isError).toBeFalsy();
    const sc = r.structuredContent as { searched: Array<{ entitySet: string }>; scope: string };
    const sets = sc.searched.map((s) => s.entitySet);
    expect(sets).toEqual(
      expect.arrayContaining(["Document_РеализацияТоваровУслуг", "Document_ПлатежноеПоручение"]),
    );
    expect(sets.length).toBeGreaterThan(10);
    expect(sc.scope).toContain("entitySet");
    expect(urls(spy).every((u) => u.includes("Date ge datetime'2026-09-10T00:00:00'"))).toBe(true);
    expect(f.violations).toEqual([]);
  });

  it("kind=bank — только банковские документы", async () => {
    const f = fake1C();
    const r = await search(f.ctx)({ kind: "bank" });
    const sets = (r.structuredContent as { searched: Array<{ entitySet: string }> }).searched.map(
      (s) => s.entitySet,
    );
    expect(sets.sort()).toEqual(
      [
        "Document_ПлатежноеПоручение",
        "Document_ПоступлениеНаРасчетныйСчет",
        "Document_СписаниеСРасчетногоСчета",
      ].sort(),
    );
    expect(f.violations).toEqual([]);
  });

  it("kind=sales и kind=cash дают разные запросы", async () => {
    const f = fake1C();
    const spy = vi.spyOn(f.conn.client, "getCollection");
    await search(f.ctx)({ kind: "sales" });
    const sales = urls(spy);
    spy.mockClear();
    await search(f.ctx)({ kind: "cash" });
    expect(sales.some((u) => u.startsWith("Document_РеализацияТоваровУслуг"))).toBe(true);
    expect(urls(spy).some((u) => u.startsWith("Document_РеализацияТоваровУслуг"))).toBe(false);
    expect(urls(spy).some((u) => u.startsWith("Document_ПриходныйКассовыйОрдер"))).toBe(true);
  });

  it("entitySet и kind вместе — отказ", async () => {
    const f = fake1C();
    const r = await search(f.ctx)({ entitySet: "Document_РеализацияТоваровУслуг", kind: "sales" });
    expect(r.isError).toBe(true);
  });

  it("minAmount по всем видам — складские без суммы пропускаются с причиной, а не отказом", async () => {
    const f = fake1C();
    const r = await search(f.ctx)({ kind: "warehouse", minAmount: 1000 });
    expect(r.isError).toBeFalsy();
    const skipped = (r.structuredContent as { skipped: Array<{ entitySet: string; reason: string }> })
      .skipped;
    expect(skipped.map((s) => s.entitySet)).toContain("Document_ПеремещениеТоваров");
    expect(skipped[0]!.reason).toContain("minAmount");
    expect(f.violations).toEqual([]);
  });

  it("сбой одного вида не валит поиск — вид в failed", async () => {
    const f = fake1C();
    const real = f.conn.client.getCollection.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (path: string) => {
      if (path.startsWith("Document_ПриходныйКассовыйОрдер")) throw new Error("Нет прав");
      return real(path);
    });
    const r = await search(f.ctx)({ kind: "cash" });
    expect(r.isError).toBeFalsy();
    const sc = r.structuredContent as { failed: Array<{ entitySet: string }>; searched: unknown[] };
    expect(sc.failed.map((x) => x.entitySet)).toEqual(["Document_ПриходныйКассовыйОрдер"]);
    expect(sc.searched).toHaveLength(1);
  });
});
