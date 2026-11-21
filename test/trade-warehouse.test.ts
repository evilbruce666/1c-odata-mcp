import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Складские документы Управления торговлей: перемещение, оприходование излишков, списание недостач, пересчёт. Состав
 * полей — как у документа, который приняла и провела копия рабочей УТ 11.5 (21.11.2026): вид запасов и ключи
 * аналитики 1С заполняет сама. Проведение проверяется по регистру «Товары на складах»: поддельная 1С проводит
 * документ, как настоящая, — по строкам и статусу, а тест портит движения, чтобы проверка их поймала.
 */

const g = (p: string, n: number) => `${p.repeat(8)}-0000-0000-0000-${String(n).padStart(12, "0")}`;
const WH_A = g("a", 1);
const WH_B = g("a", 2);
const ORG = g("b", 1);
const ITEM = g("c", 1);
const ITEM_SERIES = g("c", 2);
const KIND_PLAIN = g("d", 1);
const KIND_SERIES = g("d", 2);
const PRICE_KIND = g("e", 1);
const INCOME = g("e", 2);
const EXPENSE = g("e", 3);
const RUB = g("e", 4);
const SERIES = g("f", 1);
const EMPTY = "00000000-0000-0000-0000-000000000000";

const TABLES: Record<string, Record<string, unknown>[]> = {
  Catalog_Склады: [
    { Ref_Key: WH_A, Description: "Основной", Code: "1" },
    { Ref_Key: WH_B, Description: "Брак", Code: "2" },
  ],
  Catalog_Организации: [{ Ref_Key: ORG, Description: "ТЕСТ", Code: "1" }],
  Catalog_Номенклатура: [
    { Ref_Key: ITEM, Description: "Товар", ВидНоменклатуры_Key: KIND_PLAIN },
    { Ref_Key: ITEM_SERIES, Description: "Телефон", ВидНоменклатуры_Key: KIND_SERIES },
  ],
  Catalog_ВидыНоменклатуры: [
    { Ref_Key: KIND_PLAIN, Description: "Товары", ИспользоватьСерии: false },
    { Ref_Key: KIND_SERIES, Description: "Серийные", ИспользоватьСерии: true },
  ],
  Catalog_ВидыЦен: [{ Ref_Key: PRICE_KIND, Description: "Закупочные цены", Code: "1" }],
  ChartOfCharacteristicTypes_СтатьиДоходов: [{ Ref_Key: INCOME, Description: "Излишки", Code: "1" }],
  ChartOfCharacteristicTypes_СтатьиРасходов: [{ Ref_Key: EXPENSE, Description: "Списание брака", Code: "1" }],
  Catalog_Валюты: [{ Ref_Key: RUB, Description: "RUB", Code: "643" }],
};

/** Отбор как в 1С: по Ref_Key или по точному Code/Description; без отбора — все строки. */
function pick(rows: Record<string, unknown>[], filter: string): Record<string, unknown>[] {
  const refs = [...filter.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());
  if (refs.length) return rows.filter((r) => refs.includes(String(r["Ref_Key"]).toLowerCase()));
  const s = /(Code|Description) eq '((?:[^']|'')*)'/.exec(filter);
  if (s) return rows.filter((r) => String(r[s[1]!]) === s[2]!.replace(/''/g, "'"));
  return rows;
}

const DOCS = [
  "Document_ПеремещениеТоваров",
  "Document_ОприходованиеИзлишковТоваров",
  "Document_СписаниеНедостачТоваров",
  "Document_ПересчетТоваров",
];

/** Как проводит 1С: оприходование — приход, списание — расход, перемещение — по статусу; пересчёт — ничего. */
function movesOf(set: string, doc: Record<string, unknown>): Record<string, unknown>[] {
  const rows = (doc["Товары"] as Record<string, unknown>[]) ?? [];
  const sides: Array<[string, unknown]> =
    set === "Document_ОприходованиеИзлишковТоваров"
      ? [["Receipt", doc["Склад_Key"]]]
      : set === "Document_СписаниеНедостачТоваров"
        ? [["Expense", doc["Склад_Key"]]]
        : set === "Document_ПеремещениеТоваров"
          ? [
              ["Expense", doc["СкладОтправитель_Key"]],
              ...(doc["Статус"] === "Принято" ? ([["Receipt", doc["СкладПолучатель_Key"]]] as const) : []),
            ]
          : [];
  return sides.flatMap(([RecordType, Склад_Key]) =>
    rows.map((r) => ({
      RecordType,
      Склад_Key,
      Номенклатура_Key: r["Номенклатура_Key"],
      Характеристика_Key: EMPTY,
      Серия_Key: EMPTY,
      ВНаличии: r["Количество"],
      Active: true,
    })),
  );
}

function setup(
  o: {
    schema?: FakeSchema;
    /** Порча движений: как будто 1С провела иначе, чем в строках. */
    spoil?: (moves: Record<string, unknown>[]) => Record<string, unknown>[];
    /** После отмены проведения движения остались. */
    keepOnUnpost?: boolean;
    /** Как 1С без статуса указания серий: таблица «Серии» при записи очищается. */
    dropSeriesWithoutStatus?: boolean;
    /** Последнее проведённое перемещение организации: откуда берётся налогообложение НДС. */
    lastTransfer?: Record<string, unknown>;
    /** Действующие цены вида «Закупочные цены» по позициям: срез последних регистра цен. */
    prices?: Record<string, number>;
    /** Флаги ордерной схемы складов: по Ref_Key склада. */
    orderScheme?: Record<string, Record<string, unknown>>;
  } = {},
) {
  const posts: { set: string; body: Record<string, unknown> }[] = [];
  const actions: string[] = [];
  const docs = new Map<
    string,
    {
      set: string;
      body: Record<string, unknown>;
      posted: boolean;
      moves: Record<string, unknown>[];
      version: number;
    }
  >();
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-tw-")),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) => {
        // Механизм цен — прежний: последние документы цен заполняют «Товары», а не «Товары2_5».
        if (q.base === "Document_УстановкаЦенНоменклатуры_Товары2_5") return [];
        if (
          q.base === "Document_УстановкаЦенНоменклатуры" ||
          q.base === "Document_УстановкаЦенНоменклатуры_Товары"
        )
          return [{ Ref_Key: g("7", 1) }];
        if (q.base === "InformationRegister_ЦеныНоменклатуры_RecordType" && q.fn === "SliceLast")
          return Object.entries(o.prices ?? { [ITEM]: 12.5 })
            .filter(([item]) => (q.args["Condition"] ?? "").includes(item))
            .map(([item, Цена]) => ({
              Period: "2026-11-01T00:00:00",
              Номенклатура_Key: item,
              ВидЦены_Key: PRICE_KIND,
              Характеристика_Key: EMPTY,
              Цена,
            }));
        if (q.base === "Catalog_Склады" && o.orderScheme)
          return pick(
            TABLES["Catalog_Склады"]!.map((w) => ({ ...w, ...o.orderScheme![String(w["Ref_Key"])] })),
            q.params.get("$filter") ?? "",
          );
        if (
          q.base === "Document_ПеремещениеТоваров" &&
          !q.fn &&
          (q.params.get("$filter") ?? "").includes("Posted eq true")
        )
          return o.lastTransfer ? [o.lastTransfer] : [];
        const rows = TABLES[q.base];
        return rows && !q.fn ? pick(rows, q.params.get("$filter") ?? "") : undefined;
      },
    },
  );
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    const path = decodeURIComponent(url.replace(/^https?:\/\/[^/]+\/[^?]*?standard\.odata\//, ""));
    const method = (init?.method ?? "GET").toUpperCase();
    const json = (status: number, body: unknown) =>
      new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
    const set = /^[^(?/]+/.exec(path)?.[0] ?? "";
    const ref = /\(guid'([0-9a-f-]{36})'\)/i.exec(path)?.[1]?.toLowerCase();
    if (DOCS.includes(set)) {
      if (method === "POST" && !ref) {
        const res = await inner(input, init); // поддельная 1С проверяет поля по схеме
        if (!res.ok) return res;
        const body = JSON.parse(String(init?.body)) as Record<string, unknown>;
        posts.push({ set, body });
        const rows = (body["Товары"] as Record<string, unknown>[] | undefined) ?? [];
        if (o.dropSeriesWithoutStatus && rows.every((r) => !r["СтатусУказанияСерий"])) body["Серии"] = [];
        docs.set(String(body["Ref_Key"]).toLowerCase(), { set, body, posted: false, moves: [], version: 1 });
        return json(201, { ...body, Number: "0000-000001", DataVersion: "v1" });
      }
      const d = ref ? docs.get(ref) : undefined;
      if (d && /\/(Post|Unpost)\?/.test(path)) {
        d.posted = path.includes("/Post?");
        actions.push(`${set} ${d.posted ? "Post" : "Unpost"}`);
        d.moves = d.posted ? (o.spoil ?? ((m) => m))(movesOf(set, d.body)) : o.keepOnUnpost ? d.moves : [];
        d.version++;
        return json(200, {});
      }
      if (d && method === "GET")
        return json(200, {
          ...d.body,
          Ref_Key: ref,
          Posted: d.posted,
          DeletionMark: false,
          DataVersion: `v${d.version}`,
        });
      if (d && method === "PATCH")
        return json(200, { ...d.body, ...(JSON.parse(String(init?.body)) as object) });
    }
    if (path.startsWith("AccumulationRegister_ТоварыНаСкладах(Recorder=")) {
      const rec = /Recorder='([0-9a-f-]{36})'/i.exec(path)![1]!.toLowerCase();
      const moves = docs.get(rec)?.moves ?? [];
      // Как живая УТ 11.5: у документа без движений набора нет — 404.
      if (!moves.length)
        return json(404, {
          "odata.error": { code: "9", message: { lang: "ru", value: "Экземпляр сущности не найден" } },
        });
      return json(200, { Recorder: rec, RecordSet: moves });
    }
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
  /** Предпросмотр, затем подтверждение с тем же operationId: созданный документ. */
  const create = async (name: string, args: Record<string, unknown>) => {
    const preview = await call(name, args);
    expect(preview.isError, preview.text.slice(0, 400)).toBeFalsy();
    const done = await call(name, { ...args, confirm: true, operationId: preview.sc["operationId"] });
    expect(done.isError, done.text.slice(0, 400)).toBeFalsy();
    // Ссылку назначает коннектор при предпросмотре; в 1С уходит она же.
    return { preview, done, ref: String(posts.at(-1)!.body["Ref_Key"]) };
  };
  const post = async (entitySet: string, ref: string, postIt = true) => {
    const args = { entitySet, ref, post: postIt };
    await call("write.document.post_document", args);
    return call("write.document.post_document", { ...args, confirm: true });
  };
  return { call, create, post, posts, actions, docs, fake };
}

const T = "write.warehouse.create_transfer";
const S = "write.warehouse.create_surplus";
const W = "write.warehouse.create_writeoff";
const I = "write.warehouse.create_inventory";
const line = (quantity = 2, extra: Record<string, unknown> = {}) => ({
  nomenclatureRef: ITEM,
  quantity,
  ...extra,
});
const payloadOf = (r: { sc: Record<string, unknown> }) => r.sc["payload"] as Record<string, unknown>;

describe("складские документы в Управлении торговлей: создание", () => {
  it("перемещение: статус «Принято» по умолчанию, поля УТ, без счетов Бухгалтерии; подтверждение — один POST как в предпросмотре", async () => {
    const s = setup();
    const { preview, ref } = await s.create(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      lines: [line(2)],
      comment: "ZZ-TEST",
    });
    expect(payloadOf(preview)).toMatchObject({
      Организация_Key: ORG,
      СкладОтправитель_Key: WH_A,
      СкладПолучатель_Key: WH_B,
      ХозяйственнаяОперация: "ПеремещениеТоваров",
      Статус: "Принято",
      Posted: false,
      Комментарий: "ZZ-TEST",
      // Количество упаковок — как у всех перемещений копии рабочей базы: без упаковки равно количеству.
      Товары: [{ LineNumber: 1, Номенклатура_Key: ITEM, Количество: 2, КоличествоУпаковок: 2 }],
    });
    expect(JSON.stringify(payloadOf(preview))).not.toContain("СчетУчета");
    // Проведённых перемещений у организации нет — налогообложение пустое, с заметкой.
    expect(payloadOf(preview)).not.toHaveProperty("ПеремещениеПодДеятельность");
    expect(JSON.stringify(preview.sc["notes"])).toContain("нет проведённых перемещений");
    expect(s.posts).toHaveLength(1);
    expect(s.posts[0]!.body).toMatchObject({ ...payloadOf(preview), Ref_Key: ref });
    expect(s.fake.violations).toEqual([]);
  });

  it("перемещение: «Отгружено» — параметром; серия — в табличной части «Серии»; цена — заметка, не поле", async () => {
    const s = setup();
    const r = await s.call(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      status: "Отгружено",
      lines: [line(1, { nomenclatureRef: ITEM_SERIES, seriesRef: SERIES, price: 5 })],
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(payloadOf(r)).toMatchObject({
      Статус: "Отгружено",
      Серии: [{ LineNumber: 1, Номенклатура_Key: ITEM_SERIES, Серия_Key: SERIES, Количество: 1 }],
      // Статус «серии в таблице» — как у перемещений с сериями в базе: без него 1С очищает «Серии» при записи.
      Товары: [
        { СтатусУказанияСерий: 2, СтатусУказанияСерийОтправитель: 2, СтатусУказанияСерийПолучатель: 2 },
      ],
    });
    expect(JSON.stringify(payloadOf(r)["Товары"])).not.toContain("Цена");
    expect(JSON.stringify(r.sc["notes"])).toContain("Цена в перемещении");
  });

  it("перемещение: налогообложение НДС — из последнего проведённого перемещения организации, с заметкой", async () => {
    for (const value of ["ПродажаОблагаетсяНДС", "ЛьготноеНалогообложениеНДСНаУСН"]) {
      const s = setup({
        lastTransfer: {
          Number: "0000-000007",
          Date: "2026-11-01T10:00:00",
          ПеремещениеПодДеятельность: value,
        },
      });
      const r = await s.call(T, {
        organization: "ТЕСТ",
        fromWarehouse: "Основной",
        toWarehouse: "Брак",
        lines: [line(1)],
      });
      expect(r.isError, r.text).toBeFalsy();
      expect(payloadOf(r)["ПеремещениеПодДеятельность"]).toBe(value);
      expect(JSON.stringify(r.sc["notes"])).toContain("№ 0000-000007 от 2026-11-01");
      // Отбор — по организации документа и только проведённые.
      expect(s.fake.violations).toEqual([]);
    }
  });

  it("перемещение: 1С не сохранила серии — заметка сразу после создания; со статусом серии сохраняются", async () => {
    const s = setup({ dropSeriesWithoutStatus: true });
    const args = {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      lines: [line(1, { nomenclatureRef: ITEM_SERIES, seriesRef: SERIES })],
    };
    const { done } = await s.create(T, args);
    expect(JSON.stringify(done.sc["notes"] ?? [])).not.toContain("сохранила серий");
    // Без статуса (как раньше) — потеря видна в ответе.
    s.posts.length = 0;
    const plain = setup({ dropSeriesWithoutStatus: true });
    const inner = globalThis.fetch;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      if (init?.body && String(init.body).includes("СтатусУказанияСерий")) {
        const b = JSON.parse(String(init.body)) as Record<string, unknown>;
        for (const r of b["Товары"] as Record<string, unknown>[]) delete r["СтатусУказанияСерий"];
        return inner(input, { ...init, body: JSON.stringify(b) });
      }
      return inner(input, init);
    });
    const lost = await plain.create(T, args);
    expect(JSON.stringify(lost.done.sc["notes"])).toContain("сохранила серий 0 из 1");
  });

  it("оприходование и списание с серией — отказ без записи: 1С сбросила бы серию молча", async () => {
    const s = setup();
    const lines = [line(1, { nomenclatureRef: ITEM_SERIES, seriesRef: SERIES })];
    const sur = await s.call(S, {
      organization: "ТЕСТ",
      warehouse: "Основной",
      priceKind: PRICE_KIND,
      incomeItem: "1",
      lines,
    });
    const wo = await s.call(W, {
      organization: "ТЕСТ",
      warehouse: "Брак",
      expenseItem: "Списание брака",
      lines,
    });
    for (const r of [sur, wo]) {
      expect(r.isError).toBe(true);
      expect(r.text).toContain("пока не поддержаны");
    }
    expect(s.posts).toEqual([]);
    // Без серии — проходит, но с предупреждением: серии здесь оформляются в 1С.
    const plain = await s.call(W, {
      organization: "ТЕСТ",
      warehouse: "Брак",
      expenseItem: "Списание брака",
      lines: [line(1, { nomenclatureRef: ITEM_SERIES })],
    });
    expect(plain.isError, plain.text).toBeFalsy();
    expect(JSON.stringify(plain.sc["notes"])).toContain("оформите документ в 1С");
  });

  it("перемещение на тот же склад и без склада при нескольких — отказ, записей нет", async () => {
    const s = setup();
    const same = await s.call(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: WH_A,
      lines: [line()],
    });
    expect(same.isError).toBe(true);
    expect(same.text).toContain("совпадают");
    const none = await s.call(S, {
      organization: "ТЕСТ",
      priceKind: "Закупочные цены",
      incomeItem: "Излишки",
      lines: [line()],
    });
    expect(none.isError).toBe(true);
    expect(none.text).toContain("укажите склад");
    expect(s.posts).toEqual([]);
  });

  it("оприходование: без вида цены или статьи доходов — отказ; с ними — вид цены, статья, рубль, цена и сумма", async () => {
    const s = setup();
    const base = { organization: "ТЕСТ", warehouse: "Основной", lines: [line(3, { price: 10.5 })] };
    for (const missing of [{ incomeItem: "Излишки" }, { priceKind: "Закупочные цены" }]) {
      const r = await s.call(S, { ...base, ...missing });
      expect(r.isError).toBe(true);
      expect(r.text).toContain("уточните у пользователя");
    }
    const { preview } = await s.create(S, { ...base, priceKind: "Закупочные цены", incomeItem: "Излишки" });
    expect(payloadOf(preview)).toMatchObject({
      Склад_Key: WH_A,
      ВидЦены_Key: PRICE_KIND,
      СтатьяДоходов_Key: INCOME,
      Валюта_Key: RUB,
      Товары: [{ Номенклатура_Key: ITEM, Количество: 3, Цена: 10.5, Сумма: 31.5 }],
    });
    expect(s.fake.violations).toEqual([]);
  });

  it("оприходование без цены в строке — цена по виду цены на дату документа, с заметкой; указанная — как есть", async () => {
    const s = setup({ prices: { [ITEM]: 12.5, [ITEM_SERIES]: 99 } });
    const r = await s.call(S, {
      organization: "ТЕСТ",
      warehouse: "Основной",
      priceKind: "Закупочные цены",
      incomeItem: "Излишки",
      date: "2026-11-20",
      lines: [line(2), line(1, { nomenclatureRef: ITEM_SERIES, price: 7 })],
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(payloadOf(r)["Товары"]).toMatchObject([
      { Номенклатура_Key: ITEM, Цена: 12.5, Сумма: 25 },
      { Номенклатура_Key: ITEM_SERIES, Цена: 7, Сумма: 7 },
    ]);
    expect(JSON.stringify(r.sc["notes"])).toContain(
      "подставлены по виду цены «Закупочные цены» на 2026-11-20",
    );
  });

  it("оприходование: у позиции нет цены этого вида — отказ «укажите price», записей нет", async () => {
    const s = setup({ prices: {} });
    const r = await s.call(S, {
      organization: "ТЕСТ",
      warehouse: "Основной",
      priceKind: "Закупочные цены",
      incomeItem: "Излишки",
      lines: [line(1)],
    });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Укажите price");
    expect(s.posts).toEqual([]);
  });

  it("пересчёт: organization не используется — заметка, а не молчание", async () => {
    const s = setup();
    const lines = [{ nomenclatureRef: ITEM, factQuantity: 1, accountingQuantity: 1 }];
    const withOrg = await s.call(I, { organization: "ТЕСТ", warehouse: "Основной", lines });
    expect(JSON.stringify(withOrg.sc["notes"])).toContain("organization не используется");
    const without = await s.call(I, { warehouse: "Основной", lines });
    expect(JSON.stringify(without.sc["notes"] ?? [])).not.toContain("organization");
  });

  it("оприходование: вид цены по ссылке и пересчёт-основание", async () => {
    const s = setup();
    const r = await s.call(S, {
      organization: "ТЕСТ",
      warehouse: "Основной",
      priceKind: PRICE_KIND,
      incomeItem: "1",
      inventoryRef: g("9", 9),
      lines: [line(1)],
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(payloadOf(r)).toMatchObject({
      ВидЦены_Key: PRICE_KIND,
      СтатьяДоходов_Key: INCOME,
      ПересчетТоваров_Key: g("9", 9),
    });
  });

  it("списание: без статьи расходов — отказ; статья — составное поле с типом; цены нет", async () => {
    const s = setup();
    const base = { organization: "ТЕСТ", warehouse: "Брак", lines: [line(1, { price: 7 })] };
    const without = await s.call(W, base);
    expect(without.isError).toBe(true);
    expect(without.text).toContain("expenseItem");
    const { preview } = await s.create(W, { ...base, expenseItem: "Списание брака" });
    expect(payloadOf(preview)).toMatchObject({
      Склад_Key: WH_B,
      СтатьяРасходов: EXPENSE,
      СтатьяРасходов_Type: "StandardODATA.ChartOfCharacteristicTypes_СтатьиРасходов",
      Товары: [{ Номенклатура_Key: ITEM, Количество: 1 }],
    });
    expect(JSON.stringify(payloadOf(preview)["Товары"])).not.toContain("Цена");
    expect(s.fake.violations).toEqual([]);
  });

  it("пересчёт: учёт — в «Количество», факт — в «КоличествоФакт», статус «Выполнено», без организации", async () => {
    const s = setup();
    const { preview } = await s.create(I, {
      organization: "ТЕСТ",
      warehouse: "Основной",
      lines: [{ nomenclatureRef: ITEM, factQuantity: 5, accountingQuantity: 3, price: 2 }],
    });
    const p = payloadOf(preview);
    expect(p).toMatchObject({
      Склад_Key: WH_A,
      Статус: "Выполнено",
      УчетныеДанныеЗаполнены: true,
      Товары: [
        {
          Количество: 3,
          КоличествоФакт: 5,
          КоличествоУпаковок: 3,
          КоличествоУпаковокФакт: 5,
          Цена: 2,
          Сумма: 6,
          СуммаФакт: 10,
        },
      ],
    });
    expect(p).not.toHaveProperty("Организация_Key");
    expect(s.posts[0]!.set).toBe("Document_ПересчетТоваров");
    expect(s.fake.violations).toEqual([]);
  });

  it("вид номенклатуры ведёт серии, а серии нет — предупреждение в предпросмотре, не отказ; с серией и у обычного товара — без него", async () => {
    const s = setup();
    const args = { organization: "ТЕСТ", fromWarehouse: "Основной", toWarehouse: "Брак" };
    const warn = await s.call(T, { ...args, lines: [line(1, { nomenclatureRef: ITEM_SERIES })] });
    expect(warn.isError, warn.text).toBeFalsy();
    expect(JSON.stringify(warn.sc["notes"])).toContain("«Телефон»");
    for (const lines of [[line(1, { nomenclatureRef: ITEM_SERIES, seriesRef: SERIES })], [line(1)]]) {
      const r = await s.call(T, { ...args, lines });
      expect(r.isError, r.text).toBeFalsy();
      expect(JSON.stringify(r.sc["notes"] ?? [])).not.toContain("ведёт серии");
    }
  });
});

describe("складские документы в Управлении торговлей: проведение по регистру «Товары на складах»", () => {
  const surplusArgs = {
    organization: "ТЕСТ",
    warehouse: "Основной",
    priceKind: "Закупочные цены",
    incomeItem: "Излишки",
    lines: [line(2), line(1)],
  };

  it("оприходование: движения совпали — verified; отмена проведения — движений нет, тоже verified; проводки Бухгалтерии не читаются", async () => {
    const s = setup();
    const { ref } = await s.create(S, surplusArgs);
    const posted = await s.post("Document_ОприходованиеИзлишковТоваров", ref);
    expect(posted.isError, posted.text).toBeFalsy();
    expect(posted.sc).toMatchObject({
      done: true,
      verified: true,
      movementsVerified: true,
      movements: { expected: 1, found: 2 },
    });
    const unposted = await s.post("Document_ОприходованиеИзлишковТоваров", ref, false);
    expect(unposted.sc).toMatchObject({
      done: true,
      verified: true,
      movementsVerified: true,
      movements: { found: 0 },
    });
    expect(s.fake.requests.some((q) => q.resource.startsWith("AccountingRegister_"))).toBe(false);
  });

  it("количество в регистре другое — verified=false с расхождением и советом отменить проведение", async () => {
    const s = setup({ spoil: (m) => m.map((x) => ({ ...x, ВНаличии: Number(x["ВНаличии"]) + 1 })) });
    const { ref } = await s.create(S, surplusArgs);
    const posted = await s.post("Document_ОприходованиеИзлишковТоваров", ref);
    expect(posted.sc).toMatchObject({ done: true, verified: false, movementsVerified: false });
    expect(JSON.stringify(posted.sc["movements"])).toContain("в регистре 5, в документе 3");
    expect(String(posted.sc["note"])).toContain("post=false");
  });

  it("лишний склад в движениях и остаток после отмены — расхождение", async () => {
    const extra = setup({ spoil: (m) => [...m, { ...m[0]!, Склад_Key: WH_B }] });
    const a = await extra.create(S, surplusArgs);
    expect((await extra.post("Document_ОприходованиеИзлишковТоваров", a.ref)).sc).toMatchObject({
      verified: false,
    });
    const kept = setup({ keepOnUnpost: true });
    const b = await kept.create(S, surplusArgs);
    await kept.post("Document_ОприходованиеИзлишковТоваров", b.ref);
    const unposted = await kept.post("Document_ОприходованиеИзлишковТоваров", b.ref, false);
    expect(unposted.sc).toMatchObject({ verified: false, movementsVerified: false });
    expect(String(unposted.sc["note"])).toContain("остались движения");
  });

  it("перемещение «Принято» — расход и приход; «Отгружено» — только расход, приход на получателя — расхождение", async () => {
    const s = setup();
    const accepted = await s.create(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      lines: [line(4)],
    });
    expect((await s.post("Document_ПеремещениеТоваров", accepted.ref)).sc).toMatchObject({
      verified: true,
      movementsVerified: true,
      movements: { expected: 2, found: 2 },
    });
    const shipped = await s.create(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      status: "Отгружено",
      lines: [line(4)],
    });
    expect((await s.post("Document_ПеремещениеТоваров", shipped.ref)).sc).toMatchObject({
      movementsVerified: true,
      movements: { expected: 1, found: 1 },
    });
    // 1С провела «Отгружено» с приходом — проверка ждёт по статусу, а не по тому, что пришло.
    const wrong = setup({
      spoil: (m) => [...m, { ...m[0]!, RecordType: "Receipt", Склад_Key: WH_B }],
    });
    const c = await wrong.create(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      status: "Отгружено",
      lines: [line(4)],
    });
    expect((await wrong.post("Document_ПеремещениеТоваров", c.ref)).sc).toMatchObject({ verified: false });
  });

  it("списание — расход; направление перепутано — расхождение", async () => {
    const s = setup();
    const w = await s.create(W, {
      organization: "ТЕСТ",
      warehouse: "Брак",
      expenseItem: "Списание брака",
      lines: [line(1)],
    });
    expect((await s.post("Document_СписаниеНедостачТоваров", w.ref)).sc).toMatchObject({
      verified: true,
      movementsVerified: true,
    });
    const flip = setup({ spoil: (m) => m.map((x) => ({ ...x, RecordType: "Receipt" })) });
    const f = await flip.create(W, {
      organization: "ТЕСТ",
      warehouse: "Брак",
      expenseItem: "Списание брака",
      lines: [line(1)],
    });
    expect((await flip.post("Document_СписаниеНедостачТоваров", f.ref)).sc).toMatchObject({
      verified: false,
    });
  });

  it("склад с ордерной схемой на дату документа — сверка не выполняется, пояснение вместо ложного расхождения", async () => {
    const flag = "ИспользоватьОрдернуюСхемуПриОтраженииИзлишковНедостач";
    const since = "ДатаНачалаОрдернойСхемыПриОтраженииИзлишковНедостач";
    // Движения как при ордерной схеме: «в наличии» документ не меняет.
    const spoil = (m: Record<string, unknown>[]) =>
      m.map((x) => ({ ...x, ВНаличии: 0, КОтгрузке: x["ВНаличии"] }));
    const ordered = setup({
      spoil,
      orderScheme: { [WH_A]: { [flag]: true, [since]: "2026-01-01T00:00:00" } },
    });
    const a = await ordered.create(S, { ...surplusArgs, date: "2026-11-20" });
    const r = await ordered.post("Document_ОприходованиеИзлишковТоваров", a.ref);
    expect(r.sc).toMatchObject({ done: true, verified: true });
    expect(String(r.sc["movementsNote"])).toContain("ордерной схемой («Основной»)");
    expect(r.sc).not.toHaveProperty("movementsVerified");
    // Ордерная схема с более поздней даты — на дату документа её нет: сверка идёт и видит расхождение.
    const later = setup({ spoil, orderScheme: { [WH_A]: { [flag]: true, [since]: "2026-12-01T00:00:00" } } });
    const b = await later.create(S, { ...surplusArgs, date: "2026-11-20" });
    expect((await later.post("Document_ОприходованиеИзлишковТоваров", b.ref)).sc).toMatchObject({
      verified: false,
    });
    // Перемещение: ордерная схема при поступлении у получателя — тоже без сверки.
    const tr = setup({ orderScheme: { [WH_B]: { ИспользоватьОрдернуюСхемуПриПоступлении: true } } });
    const c = await tr.create(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      lines: [line(1)],
    });
    expect(String((await tr.post("Document_ПеремещениеТоваров", c.ref)).sc["movementsNote"])).toContain(
      "«Брак»",
    );
  });

  it("пересчёт: проведение проходит защиту, движений не ждёт — пояснение, без проверки по регистру", async () => {
    const s = setup();
    const p = await s.create(I, {
      warehouse: "Основной",
      lines: [{ nomenclatureRef: ITEM, factQuantity: 1, accountingQuantity: 1 }],
    });
    const r = await s.post("Document_ПересчетТоваров", p.ref);
    expect(r.isError, r.text).toBeFalsy();
    expect(r.sc).toMatchObject({ done: true, verified: true });
    expect(String(r.sc["movementsNote"])).toContain("не делает движений");
    expect(r.sc).not.toHaveProperty("movementsVerified");
  });

  it("пометка на удаление складского документа УТ проходит защиту; проведение другого документа УТ — по-прежнему отказ", async () => {
    const s = setup();
    const w = await s.create(W, {
      organization: "ТЕСТ",
      warehouse: "Брак",
      expenseItem: "Списание брака",
      lines: [line(1)],
    });
    const mark = await s.call("write.entity.mark_for_deletion", {
      entitySet: "Document_СписаниеНедостачТоваров",
      ref: w.ref,
    });
    expect(mark.isError, mark.text).toBeFalsy();
    const other = await s.call("write.document.post_document", {
      entitySet: "Document_РеализацияТоваровУслуг",
      ref: w.ref,
      post: true,
    });
    expect(other.isError).toBe(true);
    expect(other.text).toContain("только для Бухгалтерии");
  });
});

describe("складские инструменты в Бухгалтерии — как раньше", () => {
  it("КА и ERP (есть «Хозрасчётный» и документы как в УТ): сборщик и сверка УТ не включаются — как раньше", async () => {
    const s = setup();
    const available = await s.fake.conn.available();
    vi.spyOn(s.fake.conn, "available").mockResolvedValue(
      new Set([...available, "AccountingRegister_Хозрасчетный"]),
    );
    const r = await s.call(T, {
      organization: "ТЕСТ",
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      lines: [line(1)],
    });
    expect(JSON.stringify(r.sc["payload"] ?? {})).not.toContain("ХозяйственнаяОперация");
    const posted = await s.call("write.document.post_document", {
      entitySet: "Document_ПеремещениеТоваров",
      ref: g("9", 1),
      post: true,
      confirm: true,
    });
    expect(posted.sc).not.toHaveProperty("movementsVerified");
    expect(posted.sc).not.toHaveProperty("movementsNote");
    expect(s.fake.requests.some((q) => q.resource === "AccumulationRegister_ТоварыНаСкладах")).toBe(false);
  });

  it("серия и статус — только в УТ: в Бухгалтерии отказ без записи", async () => {
    const s = setup({ schema: "bp30" });
    const series = await s.call(W, { warehouse: "Основной", lines: [line(1, { seriesRef: SERIES })] });
    expect(series.isError).toBe(true);
    expect(series.text).toContain("только в Управлении торговлей");
    const status = await s.call(T, {
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      status: "Отгружено",
      lines: [line(1)],
    });
    expect(status.isError).toBe(true);
    expect(status.text).toContain("только в Управлении торговлей");
    expect(s.posts).toEqual([]);
  });

  it("перемещение в Бухгалтерии — сборщик Бухгалтерии: счета учёта в строках, без полей УТ", async () => {
    const s = setup({ schema: "bp30" });
    const r = await s.call(T, {
      fromWarehouse: "Основной",
      toWarehouse: "Брак",
      lines: [line(1, { price: 3 })],
    });
    expect(r.isError, r.text).toBeFalsy();
    const p = payloadOf(r);
    expect(p).not.toHaveProperty("ХозяйственнаяОперация");
    expect(p).not.toHaveProperty("Статус");
    expect((p["Товары"] as Record<string, unknown>[])[0]).toMatchObject({ Цена: 3 });
  });
});
