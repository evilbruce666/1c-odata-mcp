import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

// update_counterparty с ogrn отправлял в PATCH табличную часть из одной строки ОГРН: 1С заменяет табличную часть
// целиком, и остальные дополнительные реквизиты контрагента стирались (найдено 26.11 при правке описаний).
const REF = "dddddddd-0000-4000-8000-000000000001";
const OGRN = "eeeeeeee-0000-4000-8000-00000000000a";
const COLOR = "eeeeeeee-0000-4000-8000-00000000000b";
const SITE = "eeeeeeee-0000-4000-8000-00000000000c";
type Row = Record<string, unknown>;
const kind = (n: number) => `ffffffff-0000-4000-8000-${String(n).padStart(12, "0")}`;
const KINDS: Row[] = [
  { Ref_Key: kind(1), Description: "Телефон", Тип: "Телефон" },
  { Ref_Key: kind(2), Description: "Факс", Тип: "Телефон" },
  { Ref_Key: kind(3), Description: "Фактический адрес", Тип: "Адрес" },
  { Ref_Key: kind(4), Description: "Юридический адрес", Тип: "Адрес" },
  { Ref_Key: kind(5), Description: "Почтовый адрес", Тип: "Адрес" },
];
const contact = (n: number, Тип: string, Представление: string): Row => ({
  LineNumber: String(n),
  Тип,
  Вид_Key: kind(n),
  Представление,
});
const row = (LineNumber: string, Свойство_Key: string, Значение: string, extra: Row = {}): Row => ({
  Ref_Key: REF,
  LineNumber,
  Свойство_Key,
  Значение,
  ТекстоваяСтрока: "",
  Значение_Type: "Edm.String",
  ...extra,
});

afterEach(() => vi.restoreAllMocks());

function setup(existing: Row[], readError?: Error, contacts: Row[] = []) {
  const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
  const published = f.conn.available.bind(f.conn);
  vi.spyOn(f.conn, "available").mockImplementation(
    async () =>
      new Set([...(await published()), "ChartOfCharacteristicTypes_ДополнительныеРеквизитыИСведения"]),
  );
  vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (path: string) => {
    const q = decodeURIComponent(path);
    if (!q.startsWith("Catalog_ВидыКонтактнойИнформации"))
      return { value: [{ Ref_Key: OGRN, Description: "ОГРН" }] } as never;
    if (q.includes("substringof")) return { value: [{ Ref_Key: kind(0) }] } as never;
    return { value: KINDS } as never;
  });
  const read = vi.spyOn(f.conn.client, "getEntity").mockImplementation(async () => {
    if (readError) throw readError;
    return { Ref_Key: REF, ДополнительныеРеквизиты: existing, КонтактнаяИнформация: contacts } as never;
  });
  const patch = vi.spyOn(f.conn.client, "patch").mockImplementation(async () => ({ Ref_Key: REF }) as never);
  const tool = (
    createServer(f.ctx) as unknown as {
      _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools["write.counterparty.update_counterparty"]!;
  const call = async (confirm: boolean, args: object = { ogrn: "1027700000000" }) => {
    const r = await tool.handler({ ref: REF, ...args, confirm }, {});
    return { error: !!r.isError, s: r.structuredContent as Record<string, unknown>, text: JSON.stringify(r) };
  };
  const sent = () => (patch.mock.calls[0]![1] as Record<string, Row[]>)["ДополнительныеРеквизиты"]!;
  return { call, patch, sent, read, tool: tool as unknown as { description: string } };
}

describe("update_counterparty с ogrn: остальные дополнительные реквизиты не стираются", () => {
  const others = [row("1", COLOR, "синий"), row("2", SITE, "", { ТекстоваяСтрока: "длинный текст" })];

  it("ОГРН ещё не было — добавляется строкой, прежние строки как есть", async () => {
    const t = setup(others);
    const r = await t.call(true);
    expect(r.error, r.text).toBe(false);
    expect(t.sent()).toHaveLength(3);
    expect(t.sent().slice(0, 2)).toEqual([
      { ...others[0], LineNumber: 1 },
      { ...others[1], LineNumber: 2 },
    ]);
    expect(t.sent()[2]).toEqual({ LineNumber: 3, Свойство_Key: OGRN, Значение: "1027700000000" });
  });

  it("ОГРН уже был — меняется только его значение, на своём месте", async () => {
    const t = setup([
      others[0]!,
      row("2", OGRN.toUpperCase(), "старый", { ТекстоваяСтрока: "старый" }),
      others[1]!,
    ]);
    await t.call(true);
    expect(t.sent().map((x) => [x["Значение"], x["ТекстоваяСтрока"]])).toEqual([
      ["синий", ""],
      ["1027700000000", "1027700000000"],
      ["", "длинный текст"],
    ]);
    expect(t.sent().map((x) => x["LineNumber"])).toEqual([1, 2, 3]);
  });

  it("предпросмотр показывает все строки и ничего не пишет", async () => {
    const t = setup(others);
    const r = await t.call(false);
    expect((r.s["fields"] as Record<string, Row[]>)["ДополнительныеРеквизиты"]).toHaveLength(3);
    expect(t.patch).not.toHaveBeenCalled();
  });

  it("прежние реквизиты не прочитались — отказ без записи", async () => {
    const t = setup(others, new Error("1С не ответила"));
    const r = await t.call(true);
    expect(r.error).toBe(true);
    expect(t.patch).not.toHaveBeenCalled();
  });
});

// Замена контактов идёт по типу. Поведение прежнее; описание 26.11 обещало замену по виду («остальные виды
// остаются») — тесты на текст описаний и на состав таблицы контактов этого не проверяли (нашло ревью).
describe("update_counterparty: что происходит с контактами — как сказано в описании", () => {
  const contacts = [
    contact(4, "Адрес", "юридический прежний"),
    contact(3, "Адрес", "фактический"),
    contact(5, "Адрес", "почтовый"),
    contact(1, "Телефон", "111"),
    contact(2, "Телефон", "факс"),
  ];
  const result = (s: Record<string, unknown>) =>
    (s["fields"] as Record<string, Row[]>)["КонтактнаяИнформация"]!.map((r) => r["Представление"]);

  it("address убирает все прежние адреса и пишет один юридический, телефоны остаются", async () => {
    const t = setup([], undefined, contacts);
    const r = await t.call(false, { address: "новый" });
    expect(result(r.s)).toEqual(["111", "факс", "новый"]);
    expect((r.s["fields"] as Record<string, Row[]>)["КонтактнаяИнформация"]![2]!["Вид_Key"]).toBe(kind(4));
  });

  it("phone заменяет все строки типа «Телефон», адреса остаются", async () => {
    const t = setup([], undefined, contacts);
    const r = await t.call(false, { phone: "999" });
    expect(result(r.s)).toEqual(["юридический прежний", "фактический", "почтовый", "999"]);
  });

  it("описание говорит о замене по типу и не обещает сохранить другие виды", () => {
    const { description } = setup([]).tool;
    expect(description).toContain("все контакты того же типа");
    expect(description).toContain("включая фактический и почтовый");
    expect(description).not.toContain("остальные виды остаются");
  });

  it("phone вместе с ogrn — контрагент читается один раз", async () => {
    const t = setup([row("1", COLOR, "синий")], undefined, contacts);
    const r = await t.call(true, { phone: "999", ogrn: "1027700000000" });
    expect(r.error, r.text).toBe(false);
    expect(t.sent()).toHaveLength(2);
    expect(t.read).toHaveBeenCalledTimes(1);
  });
});

describe("«непереданное не меняется»: в PATCH только переданные простые реквизиты", () => {
  it.each([
    ["write.counterparty.update_counterparty", { name: "ZZ-TEST" }, { Description: "ZZ-TEST" }],
    ["write.catalog.update_nomenclature", { article: "AB-1" }, { Артикул: "AB-1" }],
  ])("%s", async (name, args, body) => {
    const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const patch = vi
      .spyOn(f.conn.client, "patch")
      .mockImplementation(async () => ({ Ref_Key: REF }) as never);
    const tools = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools;
    await tools[name]!.handler({ ref: REF, ...args, confirm: true }, {});
    expect(patch.mock.calls.map((c) => c[1])).toEqual([body]);
  });
});
