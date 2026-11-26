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
const allKinds = [...KINDS];
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

// Замена контактов шла по типу: новый юридический адрес стирал фактический и почтовый (старое поведение; описание
// 26.11 обещало обратное — нашло ревью, на живой базе у части контрагентов заполнены все три адреса). Почему не
// поймали: слияние контактов не проверял ни один тест.
describe("update_counterparty: заменяется только контакт того же вида", () => {
  const contacts = [
    contact(4, "Адрес", "юридический прежний"),
    contact(3, "Адрес", "фактический"),
    contact(5, "Адрес", "почтовый"),
    contact(1, "Телефон", "111"),
    contact(2, "Телефон", "факс"),
  ];
  const rows = (s: Record<string, unknown>) =>
    (s["fields"] as Record<string, Row[]>)["КонтактнаяИнформация"]!;
  const result = (s: Record<string, unknown>) => rows(s).map((r) => r["Представление"]);
  const notes = (s: Record<string, unknown>) => (s["notes"] as string[] | undefined)?.join(" ") ?? "";

  it("address меняет юридический, фактический и почтовый остаются", async () => {
    const t = setup([], undefined, contacts);
    const r = await t.call(false, { address: "новый" });
    expect(result(r.s)).toEqual(["фактический", "почтовый", "111", "факс", "новый"]);
    expect(rows(r.s)[4]!["Вид_Key"]).toBe(kind(4));
    expect(rows(r.s).map((x) => x["LineNumber"])).toEqual([1, 2, 3, 4, 5]);
    // Прежние строки уходят без изменений, кроме номера.
    expect(rows(r.s)[0]).toEqual({ ...contacts[1], LineNumber: 1 });
    expect(notes(r.s)).toMatch(/других видов того же типа \(Адрес: 2\)/);
  });

  it("phone меняет только свой вид; вид в другом регистре распознаётся", async () => {
    const upper = contacts.map((c) => ({ ...c, Вид_Key: String(c["Вид_Key"]).toUpperCase() }));
    const t = setup([], undefined, upper);
    const r = await t.call(false, { phone: "999" });
    expect(result(r.s)).toEqual(["юридический прежний", "фактический", "почтовый", "факс", "999"]);
  });

  it("две строки того же вида заменяются одной — об этом сказано в заметке", async () => {
    const t = setup([], undefined, [contact(1, "Телефон", "111"), contact(1, "Телефон", "222")]);
    const r = await t.call(false, { phone: "999" });
    expect(result(r.s)).toEqual(["999"]);
    expect(notes(r.s)).toMatch(/того же вида было 2/);
  });

  it("контакта этого вида не было — добавляется, прочие на месте", async () => {
    const t = setup([], undefined, [contact(3, "Адрес", "фактический")]);
    const r = await t.call(false, { address: "новый" });
    expect(result(r.s)).toEqual(["фактический", "новый"]);
  });

  it.each([{ phone: "999" }, { ogrn: "1027700000000" }])(
    "в ответе 1С нет прежней таблицы (%o) — отказ без записи",
    async (args) => {
      const t = setup(undefined as never, undefined, null as never);
      const r = await t.call(true, args);
      expect(r.error).toBe(true);
      expect(r.text).toMatch(/Не удалось прочитать прежние строки/);
      expect(t.patch).not.toHaveBeenCalled();
    },
  );

  it("вид контакта не найден — таблица не отправляется, в отказе сказано почему", async () => {
    const t = setup([], undefined, contacts);
    KINDS.splice(0, KINDS.length);
    try {
      const r = await t.call(true, { phone: "999" });
      expect(r.error).toBe(true);
      expect(r.text).toMatch(/Вид «Телефон» не найден/);
      const named = await t.call(false, { phone: "999", name: "ZZ-TEST" });
      expect(named.s["fields"]).toEqual({ Description: "ZZ-TEST" });
      expect(notes(named.s)).toMatch(/Вид «Телефон» не найден/);
      expect(t.patch).not.toHaveBeenCalled();
    } finally {
      KINDS.push(...allKinds);
    }
  });

  it("описание говорит о замене по виду", () => {
    const { description } = setup([]).tool;
    expect(description).toContain("только контакт того же вида");
    expect(description).toContain("включая фактический и почтовый адреса, остаются");
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

describe("copy_document — только документы", () => {
  // Копия справочника обошла бы проверки дубля ИНН и родителя; тесты копии брали только документы.
  it("справочник — отказ уже в предпросмотре: без него подтверждать нечего", async () => {
    const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const read = vi.spyOn(f.conn.client, "getEntity");
    const tools = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools;
    const r = await tools["write.document.copy_document"]!.handler(
      { entitySet: "Catalog_Контрагенты", ref: REF, confirm: false },
      {},
    );
    expect(r.isError).toBe(true);
    expect(JSON.stringify(r)).toMatch(/только документы/);
    // Образец даже не читался — до записи дело не дошло.
    expect(read).not.toHaveBeenCalled();
  });
});
