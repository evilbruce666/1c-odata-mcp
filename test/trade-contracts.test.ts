import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Договор и соглашение Управления торговлей: партнёр берётся из контрагента, перечисления — значения, которыми
 * заполнены договоры и соглашения живой УТ 11.5 (копия, чтение). Подтверждение — одна запись, та же, что в предпросмотре.
 */

const C = "cccccccc-0000-0000-0000-00000000000c";
const P = "aaaaaaaa-0000-0000-0000-00000000000a";
const K = "bbbbbbbb-0000-0000-0000-00000000000b";

function setup(o: { schema?: FakeSchema; counterparty?: Record<string, unknown> } = {}) {
  const posts: { set: string; body: Record<string, unknown> }[] = [];
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-tc-")),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) =>
        q.base === "Catalog_Контрагенты" && !q.params.get("$filter")
          ? [{ Ref_Key: C, Description: "ZZ-TEST", Партнер_Key: P, DeletionMark: false, ...o.counterparty }]
          : undefined,
    },
  );
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = decodeURIComponent(
      typeof input === "string" ? input : input instanceof URL ? input.href : input.url,
    );
    if ((init?.method ?? "GET").toUpperCase() === "POST")
      posts.push({
        set: /odata\/standard\.odata\/([^?(]+)/.exec(url)?.[1] ?? "",
        body: JSON.parse(String(init?.body)) as Record<string, unknown>,
      });
    return inner(input, init);
  });
  const call = async (name: string, args: Record<string, unknown>) => {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, sc: r.structuredContent ?? {}, text: r.content.map((c) => c.text).join("") };
  };
  return { call, posts, fake };
}

const CONTRACT = "write.catalog.create_contract";
const AGREEMENT = "write.catalog.create_agreement";

async function previewAndConfirm(s: ReturnType<typeof setup>, name: string, args: Record<string, unknown>) {
  const preview = await s.call(name, args);
  expect(preview.isError, preview.text.slice(0, 400)).toBeFalsy();
  const done = await s.call(name, { ...args, confirm: true, operationId: preview.sc["operationId"] });
  expect(done.isError, done.text.slice(0, 400)).toBeFalsy();
  return { preview, done };
}

describe("create_contract в Управлении торговлей", () => {
  it("с поставщиком: тип и хозяйственная операция УТ, партнёр из контрагента; одна запись как в предпросмотре", async () => {
    const s = setup();
    const { preview } = await previewAndConfirm(s, CONTRACT, { counterpartyRef: C, kind: "СПоставщиком" });
    expect(preview.sc["willCreate"]).toBe("Catalog_ДоговорыКонтрагентов");
    const payload = preview.sc["payload"] as Record<string, unknown>;
    expect(payload).toMatchObject({
      Партнер_Key: P,
      Контрагент_Key: C,
      ТипДоговора: "СПоставщиком",
      ХозяйственнаяОперация: "ЗакупкаУПоставщика",
      Статус: "Действует",
      Согласован: true,
      ПорядокРасчетов: "ПоДоговорамКонтрагентов",
      НалогообложениеНДСОпределяетсяВДокументе: true,
      Номер: "б/н",
    });
    expect(String(payload["Дата"])).toMatch(/^\d{4}-\d\d-\d\dT00:00:00$/);
    expect(payload).not.toHaveProperty("ВидДоговора");
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_ДоговорыКонтрагентов"]);
    expect(s.posts[0]!.body).toMatchObject(payload);
    expect(s.fake.violations).toEqual([]);
  });

  it("с покупателем и порядком «по заказам»; поля Бухгалтерии — в notes, не в запись", async () => {
    const s = setup();
    const r = await s.call(CONTRACT, {
      counterpartyRef: C,
      kind: "СПокупателем",
      settlementOrder: "orders",
      headName: "Иванов",
    });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).toMatchObject({
      ТипДоговора: "СПокупателем",
      ХозяйственнаяОперация: "РеализацияКлиенту",
      ПорядокРасчетов: "ПоЗаказам",
    });
    expect(r.sc["payload"]).not.toHaveProperty("РуководительКонтрагента");
    expect(JSON.stringify(r.sc["notes"])).toContain("headName");
  });

  it("«Прочее» и контрагент без партнёра — отказ, записей нет", async () => {
    const s = setup();
    expect((await s.call(CONTRACT, { counterpartyRef: C, kind: "Прочее" })).isError).toBe(true);
    const noPartner = setup({ counterparty: { Партнер_Key: "00000000-0000-0000-0000-000000000000" } });
    const r = await noPartner.call(CONTRACT, { counterpartyRef: C, kind: "СПокупателем" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("без партнёра");
    expect([...s.posts, ...noPartner.posts]).toEqual([]);
  });

  it("подтверждение после полуночи: дата договора — день предпросмотра", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    try {
      vi.setSystemTime(new Date(2026, 10, 20, 23, 59, 58));
      const s = setup();
      const args = { counterpartyRef: C, kind: "СПокупателем" };
      const preview = await s.call(CONTRACT, args);
      vi.setSystemTime(new Date(2026, 10, 21, 0, 0, 5));
      const done = await s.call(CONTRACT, { ...args, confirm: true, operationId: preview.sc["operationId"] });
      expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
      expect(s.posts[0]!.body["Дата"]).toBe("2026-11-20T00:00:00");
    } finally {
      vi.useRealTimers();
    }
  });

  it("в Бухгалтерии — как раньше: вид договора и владелец, без полей УТ", async () => {
    const s = setup({ schema: "bp30" });
    const r = await s.call(CONTRACT, { counterpartyRef: C, kind: "СПокупателем" });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).toMatchObject({ ВидДоговора: "СПокупателем", Owner_Key: C });
    expect(r.sc["payload"]).not.toHaveProperty("ТипДоговора");
  });
});

describe("create_agreement", () => {
  it("с клиентом и видом цен: цены по виду в шапке; одна запись", async () => {
    const s = setup();
    const { preview } = await previewAndConfirm(s, AGREEMENT, {
      side: "client",
      counterpartyRef: C,
      name: "ZZ-TEST соглашение",
      priceKindRef: K,
    });
    expect(preview.sc["willCreate"]).toBe("Catalog_СоглашенияСКлиентами");
    expect(preview.sc["payload"]).toMatchObject({
      Партнер_Key: P,
      Контрагент_Key: C,
      ХозяйственнаяОперация: "РеализацияКлиенту",
      ВидЦен_Key: K,
      ВариантРасчетаЦен: "ВидЦенВШапке",
      Статус: "Действует",
      Согласован: true,
      Типовое: false,
      ИспользуютсяДоговорыКонтрагентов: true,
      ЦенаВключаетНДС: false,
    });
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_СоглашенияСКлиентами"]);
    expect(s.fake.violations).toEqual([]);
  });

  it("с клиентом без вида цен — «без цен»; с поставщиком — дата начала и закупка", async () => {
    const s = setup();
    const client = await s.call(AGREEMENT, { side: "client", counterpartyRef: C, name: "ZZ-TEST" });
    expect(client.sc["payload"]).toMatchObject({ ВариантРасчетаЦен: "БезЦен" });
    expect(client.sc["payload"]).not.toHaveProperty("ВидЦен_Key");
    const supplier = await s.call(AGREEMENT, {
      side: "supplier",
      counterpartyRef: C,
      name: "ZZ-TEST",
      pricesIncludeVat: true,
    });
    expect(supplier.isError, supplier.text.slice(0, 300)).toBeFalsy();
    expect(supplier.sc["willCreate"]).toBe("Catalog_СоглашенияСПоставщиками");
    expect(supplier.sc["payload"]).toMatchObject({
      ХозяйственнаяОперация: "ЗакупкаУПоставщика",
      ЦенаВключаетНДС: true,
      СпособДоставки: "СиламиПоставщикаДоНашегоСклада",
    });
    expect(String((supplier.sc["payload"] as Record<string, unknown>)["ДатаНачалаДействия"])).toMatch(
      /T00:00:00$/,
    );
    expect(s.fake.violations).toEqual([]);
  });

  it("вид цен у соглашения с поставщиком — отказ", async () => {
    const r = await setup().call(AGREEMENT, {
      side: "supplier",
      counterpartyRef: C,
      name: "ZZ",
      priceKindRef: K,
    });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("priceKindRef");
  });
});
