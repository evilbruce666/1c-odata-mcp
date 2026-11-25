import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

// search_documents объявлял counterparty и organization, но не заполнял их: модель видела, чьи документы, только
// открыв каждый. get_document отдавал сырую сущность с GUID (аудит 25.11).
const SALES = "Document_РеализацияТоваровУслуг";
const COUNTERPARTIES = "Catalog_Контрагенты";
const ORGS = "Catalog_Организации";
const EMPTY = "00000000-0000-0000-0000-000000000000";
const guid = (n: number) => `aaaaaaaa-0000-4000-8000-${String(n).padStart(12, "0")}`;
const ORG = guid(9000);
const DOC_REF = guid(8000);

afterEach(() => vi.restoreAllMocks());

interface Opts {
  env?: Record<string, string>;
  docs?: Array<Record<string, unknown>>;
  entity?: Record<string, unknown>;
  failNames?: boolean;
  longName?: number;
}
function setup(o: Opts) {
  const f = fake1C(o.env ?? {});
  const resources: string[] = [];
  vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (path: string) => {
    const set = decodeURIComponent(path.split("?")[0]!);
    resources.push(set);
    if (set === SALES) return { value: o.docs ?? [] } as never;
    if (set === ORGS) return { value: [{ Ref_Key: ORG, Description: "ТЕСТ" }] } as never;
    if (o.failNames) throw new Error("1С не ответила");
    const refs = [...decodeURIComponent(path).matchAll(/guid'([0-9a-f-]{36})'/g)].map((m) => m[1]!);
    return {
      value: refs
        .filter((r) => !r.startsWith("bbbbbbbb"))
        .map((r) => ({ Ref_Key: r, Description: `Имя ${r.slice(-4)}${"я".repeat(o.longName ?? 0)}` })),
    } as never;
  });
  vi.spyOn(f.conn.client, "getEntity").mockImplementation(async () => (o.entity ?? {}) as never);
  const tools = (
    createServer(f.ctx) as unknown as {
      _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  const call = async (name: string, args: object) => {
    const r = await tools[name]!.handler(args, {});
    return {
      r,
      s: r.structuredContent as Record<string, unknown>,
      text: (r.content[0] as { text: string }).text,
    };
  };
  return { call, resources };
}
const search = { entitySet: SALES, limit: 100, offset: 0, postedOnly: false };
const doc = (i: number, extra: Record<string, unknown> = {}) => ({
  Ref_Key: guid(i),
  Number: `0000-${i}`,
  Date: "2026-10-01T10:00:00",
  Posted: true,
  DeletionMark: false,
  СуммаДокумента: 100 + i,
  Контрагент_Key: guid(100 + i),
  Организация_Key: ORG,
  ...extra,
});
const rowsOf = (s: Record<string, unknown>) => s["rows"] as Array<Record<string, unknown>>;

describe("search_documents: названия контрагента и организации", () => {
  it("названия — из справочников, пустая и неизвестная ссылка — без поля", async () => {
    const t = setup({
      docs: [
        doc(1),
        doc(2, { Контрагент_Key: EMPTY }),
        doc(3, { Контрагент_Key: "bbbbbbbb-0000-4000-8000-000000000003" }),
      ],
    });
    const { r, s } = await t.call("read.document.search_documents", search);
    expect(r.isError).toBeFalsy();
    const [a, b, c] = rowsOf(s).sort((x, y) => String(x["number"]).localeCompare(String(y["number"])));
    expect(a).toMatchObject({ counterparty: "Имя 0101", organization: "ТЕСТ" });
    expect(b).not.toHaveProperty("counterparty");
    expect(c).not.toHaveProperty("counterparty");
    expect(c).toMatchObject({ organization: "ТЕСТ" });
    // Внутренние ссылки наружу не выходят.
    expect(JSON.stringify(a)).not.toContain("Контрагент_Key");
  });

  it("45 разных контрагентов — три пачки к справочнику, организации — один запрос", async () => {
    const t = setup({ docs: Array.from({ length: 45 }, (_, i) => doc(i + 1)) });
    await t.call("read.document.search_documents", { ...search, limit: 100 });
    expect(t.resources.filter((x) => x === COUNTERPARTIES)).toHaveLength(3);
    expect(t.resources.filter((x) => x === ORGS)).toHaveLength(1);
  });

  it("названия читаются только для строк страницы", async () => {
    const t = setup({ docs: Array.from({ length: 60 }, (_, i) => doc(i + 1)) });
    await t.call("read.document.search_documents", { ...search, limit: 10 });
    expect(t.resources.filter((x) => x === COUNTERPARTIES)).toHaveLength(1);
  });

  it("справочник не читается — поиск отвечает без названий", async () => {
    const t = setup({ docs: [doc(1)], failNames: true });
    const { r, s } = await t.call("read.document.search_documents", search);
    expect(r.isError).toBeFalsy();
    expect(rowsOf(s)[0]).not.toHaveProperty("counterparty");
  });

  it("длинные названия учтены в размере страницы: ничего не обрезано и не потеряно", async () => {
    const t = setup({
      env: { ODATA_MAX_RESPONSE_CHARS: "10000" },
      docs: Array.from({ length: 60 }, (_, i) => doc(i + 1)),
      longName: 300,
    });
    const seen = new Set<string>();
    let offset = 0;
    for (let i = 0; i < 60; i++) {
      const { s, text } = await t.call("read.document.search_documents", { ...search, limit: 50, offset });
      expect(text.length).toBeLessThanOrEqual(10_000);
      for (const row of rowsOf(s)) {
        seen.add(String(row["number"]));
        expect(row["counterparty"]).toBeTruthy();
      }
      if (s["nextOffset"] === undefined) break;
      offset = s["nextOffset"] as number;
    }
    expect(seen.size).toBe(60);
  });
});

const entity = {
  "odata.metadata": "x",
  Ref_Key: DOC_REF,
  DataVersion: "AAAA",
  Number: "0000-1",
  Posted: false,
  Комментарий: "",
  Договор_Key: EMPTY,
  Контрагент_Key: guid(101),
  Организация_Key: ORG,
  Склад_Key: guid(102),
  Ответственный_Key: guid(103),
  ДатаПлатежа: "0001-01-01T00:00:00",
  СуммаДокумента: 0,
  Услуги: [],
  Товары: [
    { LineNumber: "1", Номенклатура_Key: guid(104), Количество: 2, Содержание: "", СчетУчета_Key: EMPTY },
  ],
  Сделка: guid(105),
  Сделка_Type: "StandardODATA.Catalog_Контрагенты",
  "ДокументОснование@navigationLinkUrl": "x",
};

describe("get_document: compact", () => {
  it("без compact — сущность 1С как есть", async () => {
    const t = setup({ entity });
    const { s } = await t.call("read.document.get_document", { entitySet: SALES, ref: DOC_REF });
    expect({ ...s, database: undefined }).toEqual({ ...entity, database: undefined });
    expect(t.resources).toEqual([]);
  });

  it("compact — без пустых значений и служебных ключей, с названиями рядом со ссылками", async () => {
    const t = setup({ entity });
    const { s } = await t.call("read.document.get_document", {
      entitySet: SALES,
      ref: DOC_REF,
      compact: true,
    });
    for (const k of [
      "odata.metadata",
      "Комментарий",
      "Договор_Key",
      "ДатаПлатежа",
      "Услуги",
      "ДокументОснование@navigationLinkUrl",
    ])
      expect(s, k).not.toHaveProperty([k]);
    // Значения, несущие смысл, остаются.
    expect(s).toMatchObject({ Posted: false, СуммаДокумента: 0, DataVersion: "AAAA", Ref_Key: DOC_REF });
    expect(s).toMatchObject({
      Контрагент_Key: guid(101),
      Контрагент_Название: "Имя 0101",
      Организация_Название: "ТЕСТ",
      Склад_Название: "Имя 0102",
      Сделка_Название: "Имя 0105",
    });
    expect(s).not.toHaveProperty(["Ответственный_Название"]);
    const line = (s["Товары"] as Array<Record<string, unknown>>)[0]!;
    expect(line).toEqual({
      LineNumber: "1",
      Номенклатура_Key: guid(104),
      Номенклатура_Название: "Имя 0104",
      Количество: 2,
    });
  });

  it("compact при отказе справочника — без названий, но не ошибка", async () => {
    const t = setup({ entity, failNames: true });
    const { r, s } = await t.call("read.document.get_document", {
      entitySet: SALES,
      ref: DOC_REF,
      compact: true,
    });
    expect(r.isError).toBeFalsy();
    expect(s).not.toHaveProperty(["Контрагент_Название"]);
    expect(s).toMatchObject({ Организация_Название: "ТЕСТ" });
  });
});
