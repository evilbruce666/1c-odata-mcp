import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Номенклатура Управления торговлей: вид обязателен, свойства вида переносятся в позицию (так их переносит 1С при
 * создании в форме — на живой базе совпадают у всех позиций), единица по умолчанию «шт», ставка НДС — основная
 * действующая на день предпросмотра. Подтверждение — одна запись, та же, что в предпросмотре.
 */

const g = (n: number) =>
  `00000000-0000-0000-0000-${String(n).padStart(12, "0")}`.replace(/^0{8}/, "aaaaaaaa");
const GOODS = g(1);
const SERVICES = g(2);
const FOLDER = g(3);
const SERIES_OWNER = g(4);
const PCS = g(10);
const KG = g(11);
/** Вторая «шт» — как на живой УТ: код «796 » заведён позже, номенклатура им не пользуется. */
const PCS2 = g(12);
const VAT20 = g(20);
const VAT22 = g(22);
const NOVAT = g(29);

const TABLES: Record<string, Record<string, unknown>[]> = {
  Catalog_ВидыНоменклатуры: [
    {
      Ref_Key: GOODS,
      Description: "Товары",
      IsFolder: false,
      ТипНоменклатуры: "Товар",
      ИспользованиеХарактеристик: "НеИспользовать",
      ВладелецСерий_Key: SERIES_OWNER,
      ВариантОформленияПродажи: "РеализацияТоваровУслуг",
      ИспользоватьУпаковки: false,
    },
    { Ref_Key: SERVICES, Description: "Услуги", IsFolder: false, ТипНоменклатуры: "Услуга" },
    { Ref_Key: FOLDER, Description: "Группа видов", IsFolder: true },
    { Ref_Key: g(5), Description: "Дубль", IsFolder: false, ТипНоменклатуры: "Товар" },
    { Ref_Key: g(6), Description: "Дубль", IsFolder: false, ТипНоменклатуры: "Товар" },
  ],
  Catalog_УпаковкиЕдиницыИзмерения: [
    { Ref_Key: PCS, Description: "шт", Code: "0796" },
    { Ref_Key: KG, Description: "кг", Code: "0166" },
    { Ref_Key: PCS2, Description: "шт", Code: "796 " },
  ],
  Catalog_СтавкиНДС: [
    {
      Ref_Key: VAT20,
      Description: "20%",
      Основная: true,
      РасчетнаяСтавка: false,
      НачалоПериода: "2019-01-01T00:00:00",
      КонецПериода: "2025-12-31T00:00:00",
    },
    {
      Ref_Key: VAT22,
      Description: "22%",
      Основная: true,
      РасчетнаяСтавка: false,
      НачалоПериода: "2026-01-01T00:00:00",
      КонецПериода: "0001-01-01T00:00:00",
    },
    { Ref_Key: NOVAT, Description: "Без НДС", Основная: false, РасчетнаяСтавка: false },
  ],
};

/** Отбор как в 1С для запросов коннектора: по ссылке, по наименованию или коду, по признаку «Основная». */
function select(rows: Record<string, unknown>[], filter: string) {
  const ref = /Ref_Key eq guid'([0-9a-f-]{36})'/i.exec(filter)?.[1];
  if (ref) return rows.filter((r) => r["Ref_Key"] === ref);
  const values = [...filter.matchAll(/(?:Description|Code) eq '((?:[^']|'')*)'/g)].map((m) =>
    m[1]!.replace(/''/g, "'"),
  );
  return rows.filter(
    (r) =>
      (values.length === 0 ||
        values.includes(String(r["Description"])) ||
        // 1С сравнивает строки без хвостовых пробелов: «796» находит код «796 ».
        values.includes(String(r["Code"] ?? "").trimEnd())) &&
      (!filter.includes("Основная eq true") || r["Основная"] === true),
  );
}

function setup(schema: FakeSchema = "ut11", usage: Record<string, number> = { [PCS]: 14998, [PCS2]: 2 }) {
  const posts: { set: string; body: Record<string, unknown> }[] = [];
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-tn-")),
    },
    {
      schema,
      data: (q: FakeQuery) => {
        const filter = q.params.get("$filter") ?? "";
        const unit = /ЕдиницаИзмерения_Key eq guid'([0-9a-f-]{36})'/.exec(filter)?.[1];
        if (q.base === "Catalog_Номенклатура" && unit)
          return Array.from({ length: usage[unit] ?? 0 }, (_, i) => ({ Ref_Key: g(1000 + i) }));
        const rows = TABLES[q.base];
        return rows ? select(rows, q.params.get("$filter") ?? "") : undefined;
      },
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
  const call = async (args: Record<string, unknown>, name = "write.catalog.create_nomenclature") => {
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

const at = (y: number, m: number, d: number) => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date(y, m - 1, d, 12, 0, 0));
};
afterEach(() => vi.useRealTimers());

describe("create_nomenclature в Управлении торговлей", () => {
  it("без вида — отказ с объяснением, записей нет", async () => {
    const s = setup();
    const r = await s.call({ name: "ZZ-TEST товар" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("kind");
    expect(s.posts).toEqual([]);
  });

  it("свойства вида — в позицию; «шт» и основная ставка на день (2026 — 22%); одна запись как в предпросмотре", async () => {
    at(2026, 11, 20);
    const s = setup();
    const args = { name: "ZZ-TEST товар", kind: "Товары", article: "ZZ-1" };
    const preview = await s.call(args);
    expect(preview.isError, preview.text.slice(0, 400)).toBeFalsy();
    const payload = preview.sc["payload"] as Record<string, unknown>;
    expect(payload).toMatchObject({
      Description: "ZZ-TEST товар",
      Артикул: "ZZ-1",
      ВидНоменклатуры_Key: GOODS,
      ТипНоменклатуры: "Товар",
      ИспользованиеХарактеристик: "НеИспользовать",
      ВладелецСерий_Key: SERIES_OWNER,
      ВариантОформленияПродажи: "РеализацияТоваровУслуг",
      ЕдиницаИзмерения_Key: PCS,
      ЕдиницаДляОтчетов_Key: PCS,
      КоэффициентЕдиницыДляОтчетов: 1,
      СтавкаНДС_Key: VAT22,
    });
    // Пустые и «ложные» свойства вида не пишутся: 1С оставит значения по умолчанию.
    expect(payload).not.toHaveProperty("ИспользоватьУпаковки");
    expect(JSON.stringify(preview.sc["notes"])).toContain("основная действующая");
    const done = await s.call({ ...args, confirm: true, operationId: preview.sc["operationId"] });
    expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_Номенклатура"]);
    expect(s.posts[0]!.body).toMatchObject(payload);
    expect(s.fake.violations).toEqual([]);
  });

  it("в 2025 основная действующая — 20%; явные единица и ставка — по названию", async () => {
    at(2025, 6, 1);
    const s = setup();
    const r = await s.call({ name: "ZZ-TEST", kind: "Товары" });
    expect(r.sc["payload"]).toMatchObject({ СтавкаНДС_Key: VAT20 });
    const explicit = await s.call({ name: "ZZ-TEST", kind: "Товары", unit: "кг", vatRate: "Без НДС" });
    expect(explicit.sc["payload"]).toMatchObject({ ЕдиницаИзмерения_Key: KG, СтавкаНДС_Key: NOVAT });
  });

  it("единица по коду ОКЕИ без ведущего нуля находится", async () => {
    const r = await setup().call({ name: "ZZ-TEST", kind: "Товары", unit: "166" });
    expect(r.sc["payload"]).toMatchObject({ ЕдиницаИзмерения_Key: KG });
  });

  it.each([
    [{ kind: "Группа видов" }, "группа видов"],
    [{ kind: "Дубль" }, "несколько"],
    [{ kind: "Нет такого" }, "не найден"],
    [{ kind: "Товары", isService: true }, "услуг"],
    [{ kind: "Товары", unit: "литр" }, "не найден"],
  ])("%j — отказ", async (extra, text) => {
    const s = setup();
    const r = await s.call({ name: "ZZ-TEST", ...extra });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(text);
    expect(s.posts).toEqual([]);
  });

  it("услуга — вид услуг с isService", async () => {
    const r = await setup().call({ name: "ZZ-TEST услуга", kind: SERVICES, isService: true });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).toMatchObject({ ВидНоменклатуры_Key: SERVICES, ТипНоменклатуры: "Услуга" });
  });

  it("папка номенклатуры в УТ проходит защиту; папка контрагентов — отказ (иерархии нет)", async () => {
    const s = setup();
    const folder = await s.call(
      { entitySet: "Catalog_Номенклатура", name: "ZZ-TEST папка" },
      "write.entity.create_folder",
    );
    expect(folder.isError, folder.text.slice(0, 300)).toBeFalsy();
    expect(folder.sc["payload"]).toMatchObject({ IsFolder: true, Description: "ZZ-TEST папка" });
    const cp = await s.call(
      { entitySet: "Catalog_Контрагенты", name: "ZZ-TEST" },
      "write.entity.create_folder",
    );
    expect(cp.isError).toBe(true);
    expect(cp.text).toContain("только для Бухгалтерии");
    expect(s.posts).toEqual([]);
  });

  it("update_nomenclature в УТ проходит защиту", async () => {
    const r = await setup().call({ ref: GOODS, name: "ZZ-TEST" }, "write.catalog.update_nomenclature");
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["dryRun"]).toBe(true);
  });

  it("в Бухгалтерии — как раньше: без вида и полей УТ", async () => {
    const r = await setup("bp30").call({ name: "ZZ-TEST" });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).not.toHaveProperty("ВидНоменклатуры_Key");
    expect(r.sc["payload"]).not.toHaveProperty("ЕдиницаДляОтчетов_Key");
  });

  it("две «шт» в базе — взята самая используемая (вдвое и больше), с заметкой; без перевеса — отказ", async () => {
    const r = await setup().call({ name: "ZZ-TEST", kind: "Товары" });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).toMatchObject({ ЕдиницаИзмерения_Key: PCS });
    expect(JSON.stringify(r.sc["notes"])).toContain("самая используемая");
    const both = await setup("ut11", { [PCS]: 10, [PCS2]: 6 }).call({ name: "ZZ-TEST", kind: "Товары" });
    expect(both.isError).toBe(true);
    expect(both.text).toContain(PCS2);
    const explicit = await setup().call({ name: "ZZ-TEST", kind: "Товары", unit: PCS2 });
    expect(explicit.sc["payload"]).toMatchObject({ ЕдиницаИзмерения_Key: PCS2 });
  });

  it("изменение номенклатуры: 1С отвечает 500 (как на живой УТ) — понятный отказ; у контрагента — как был", async () => {
    const s = setup();
    const inner = globalThis.fetch;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      if ((init?.method ?? "GET").toUpperCase() === "PATCH")
        return new Response(
          JSON.stringify({
            "odata.error": {
              code: "-1",
              message: { lang: "ru", value: 'Не удалось записать "X (Номенклатура)"!' },
            },
          }),
          { status: 500, headers: { "content-type": "application/json" } },
        );
      // Папка назначения — настоящая группа, как в 1С (move_to_folder это проверяет).
      if (
        decodeURIComponent(String(input instanceof Request ? input.url : input)).includes(`(guid'${FOLDER}')`)
      )
        return Response.json({ Ref_Key: FOLDER, IsFolder: true, DeletionMark: false, Parent_Key: "" });
      return inner(input, init);
    });
    const MARK = "write.entity.mark_for_deletion";
    const r = await s.call(
      { entitySet: "Catalog_Номенклатура", ref: GOODS, mark: true, confirm: true },
      MARK,
    );
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Измените позицию в 1С");
    for (const [tool, args] of [
      ["write.catalog.update_nomenclature", { ref: GOODS, article: "ZZ-2" }],
      ["write.entity.move_to_folder", { entitySet: "Catalog_Номенклатура", ref: GOODS, folder: FOLDER }],
    ] as const) {
      const u = await s.call({ ...args, confirm: true }, tool);
      expect(u.isError, tool).toBe(true);
      expect(u.text, tool).toContain("Измените позицию в 1С");
    }
    expect(r.text).toContain("Не удалось записать");
    const other = await s.call(
      { entitySet: "Catalog_Контрагенты", ref: GOODS, mark: true, confirm: true },
      MARK,
    );
    expect(other.isError).toBe(true);
    expect(other.text).not.toContain("Измените позицию в 1С");
  });

  it.each([
    ["504 в УТ — исход неизвестен", "ut11", 504, "Gateway Timeout"],
    ["500 другого текста в УТ", "ut11", 500, "Объект заблокирован"],
    ["502 с тем же текстом в УТ", "ut11", 502, 'Не удалось записать "X"!'],
    ["500 «Не удалось записать» в Бухгалтерии", "bp30", 500, 'Не удалось записать "X"!'],
  ] as const)("%s — без объяснения про УТ (ревью)", async (_, schema, status, text) => {
    const s = setup(schema);
    const inner = globalThis.fetch;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      if ((init?.method ?? "GET").toUpperCase() === "PATCH")
        return new Response(
          JSON.stringify({ "odata.error": { code: "-1", message: { lang: "ru", value: text } } }),
          {
            status,
            headers: { "content-type": "application/json" },
          },
        );
      return inner(input, init);
    });
    const r = await s.call(
      { entitySet: "Catalog_Номенклатура", ref: GOODS, mark: true, confirm: true },
      "write.entity.mark_for_deletion",
    );
    expect(r.isError).toBe(true);
    expect(r.text).not.toContain("Измените позицию в 1С");
  });

  it("поля вида — только опубликованные: нет ВладелецСерий_Key у вида в $metadata — не в payload", async () => {
    const s = setup();
    const meta = await s.fake.conn.getMetadata();
    const kinds = meta.entities.get("Catalog_ВидыНоменклатуры")!;
    vi.spyOn(s.fake.conn, "getMetadata").mockResolvedValue({
      ...meta,
      entities: new Map([
        ...meta.entities,
        [
          "Catalog_ВидыНоменклатуры",
          { ...kinds, properties: kinds.properties.filter((p) => p.name !== "ВладелецСерий_Key") },
        ],
      ]),
    });
    const r = await s.call({ name: "ZZ-TEST", kind: "Товары" });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).not.toHaveProperty("ВладелецСерий_Key");
    expect(r.sc["payload"]).toMatchObject({ ТипНоменклатуры: "Товар" });
  });

  it("справочник «Ставки НДС» не опубликован — понятный отказ, записей нет", async () => {
    const s = setup();
    const available = await s.fake.conn.available();
    vi.spyOn(s.fake.conn, "available").mockResolvedValue(
      new Set([...available].filter((x) => x !== "Catalog_СтавкиНДС")),
    );
    const r = await s.call({ name: "ZZ-TEST", kind: "Товары" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Ставки НДС");
    expect(s.posts).toEqual([]);
  });
});
