import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Установка и чтение цен Управления торговлей на синтетической базе. Метрика приёмки: движения документа и срез цен
 * после проведения равны заданным до копейки, другие позиции и виды цен не тронуты; отмена проведения возвращает
 * прежний срез; повтор с тем же operationId не создаёт второй документ; расхождение движений — verified=false.
 * Документы, регистр и срез последних — модель ниже (поверх поддельной 1С), справочники — поддельная 1С. Имена
 * выдуманы.
 */

afterEach(() => {
  vi.unstubAllGlobals();
});

const uuid = (k: string, i: number) => `${k}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const EMPTY = "00000000-0000-0000-0000-000000000000";
const I = (i: number) => uuid("a", i);
const CUR = uuid("c", 1);
const K_BASE = uuid("e", 1);
const K_MARKUP = uuid("e", 2);
const K_OFF = uuid("e", 3);
const G = uuid("d", 1);
const DOC_TYPE = "StandardODATA.Document_УстановкаЦенНоменклатуры";

const KINDS = [
  {
    Ref_Key: K_BASE,
    Description: "Закупочная",
    ВалютаЦены_Key: CUR,
    СпособЗаданияЦены: "ЗаполнятьПоДаннымИБ",
    Статус: "Действует",
  },
  {
    Ref_Key: K_MARKUP,
    Description: "Розничная",
    ВалютаЦены_Key: CUR,
    СпособЗаданияЦены: "НаценкаНаДругойВидЦен",
    Статус: "Действует",
    БазовыйВидЦены_Key: K_BASE,
  },
  {
    Ref_Key: K_OFF,
    Description: "Старая",
    ВалютаЦены_Key: CUR,
    СпособЗаданияЦены: "Вручную",
    Статус: "НеДействует",
  },
].map((k) => ({
  DeletionMark: false,
  IsFolder: false,
  БазовыйВидЦены_Key: EMPTY,
  ВлияющиеВидыЦен: [],
  ...k,
}));

// 3-я позиция и дальше — без цен; 100 позиций нужны для проверки размера предпросмотра.
const ITEMS = Array.from({ length: 110 }, (_, n) => n + 1).map((i) => ({
  Ref_Key: I(i),
  Description: `Позиция ${i}`,
  Артикул: `A-${i}`,
  Parent_Key: G,
  IsFolder: false,
  DeletionMark: false,
}));

interface Rec {
  Active?: boolean;
  Period: string;
  Номенклатура_Key: string;
  Характеристика_Key: string;
  ВидЦены_Key: string;
  Цена: number;
  Recorder?: string;
}
const rec = (item: string, kind: string, price: number, period: string): Rec => ({
  Period: period,
  Номенклатура_Key: item,
  Характеристика_Key: EMPTY,
  ВидЦены_Key: kind,
  Цена: price,
});

interface Options {
  /** Проведение пишет в регистр не ту цену (порча для проверки сверки). */
  corruptPost?: boolean;
  /** Проведение пишет ещё одну, лишнюю запись. */
  extraMovement?: boolean;
  /** Отмена проведения оставляет записи в регистре. */
  keepAfterUnpost?: boolean;
  /** Чтение набора записей регистра падает. */
  failMovements?: boolean;
  /** В наборе записей есть неактивная запись (Active=false) — 1С её не учитывает. */
  inactiveMovement?: boolean;
  /** В прежнем регистре цен нет записей (для выбора механизма). */
  legacyEmpty?: boolean;
  /** Табличная часть последних проведённых документов: legacy или 2.5. */
  recentRows?: "Товары" | "Товары2_5" | "both";
  schema?: FakeSchema;
  /** Справочник номенклатуры вместо ITEMS (масштаб: длинные названия, тысяча позиций). */
  items?: typeof ITEMS;
  /** Записи регистра цен до проверки вместо стандартных. */
  base?: Rec[];
  /** Предел текста ответа (ODATA_MAX_RESPONSE_CHARS). */
  maxResponseChars?: number;
  /** Предел строк чтения (ODATA_ANALYTICS_MAX_ROWS). */
  analyticsMaxRows?: number;
}

function setup(o: Options = {}) {
  const bp = o.schema === "bp30";
  const items = o.items ?? ITEMS;
  const docs = new Map<string, { body: Record<string, unknown>; posted: boolean; version: number }>();
  const base: Rec[] = o.base ?? [
    rec(I(1), K_BASE, 100, "2026-11-01T00:00:01"),
    rec(I(1), K_MARKUP, 150, "2026-11-01T00:00:01"),
    rec(I(2), K_BASE, 200, "2026-11-01T00:00:01"),
    rec(I(2), K_BASE, 210, "2026-12-01T00:00:00"), // установлена на будущее
  ];
  const posts: string[] = [];
  const unposted = new Set<string>();
  const movementsOf = (ref: string): Rec[] => {
    const d = docs.get(ref);
    if (!d || (!d.posted && !(o.keepAfterUnpost && unposted.has(ref)))) return [];
    const rows = ((d.body["Товары"] as Record<string, unknown>[]) ?? []).map((r) => ({
      Period: String(d.body["Date"]),
      Номенклатура_Key: String(r["Номенклатура_Key"]),
      Характеристика_Key: String(r["Характеристика_Key"] ?? EMPTY),
      // Бухгалтерия: тип цен — в шапке документа.
      ВидЦены_Key: String(bp ? d.body["ТипЦен_Key"] : r["ВидЦены_Key"]),
      Цена: Number(r["Цена"]) + (o.corruptPost ? 1 : 0),
      Recorder: ref,
    }));
    return [
      ...rows,
      ...(o.extraMovement ? [{ ...rows[0]!, Номенклатура_Key: I(3) }] : []),
      ...(o.inactiveMovement && rows.length ? [{ ...rows[0]!, Номенклатура_Key: I(4), Active: false }] : []),
    ];
  };
  const records = () => [...base, ...[...docs.keys()].flatMap(movementsOf)];
  // Запись регистра в ответе 1С: в Бухгалтерии — тип цен и без характеристики.
  const out = (r: Rec) => {
    if (!bp) return r;
    const { ВидЦены_Key, Характеристика_Key: _c, ...rest } = r;
    return { ...rest, ТипЦен_Key: ВидЦены_Key };
  };
  const guids = (text: string) =>
    // В Condition среза апострофы удвоены: guid''…''.
    [...text.matchAll(/guid'{1,2}([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());

  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-pricing-")),
      ...(o.maxResponseChars ? { ODATA_MAX_RESPONSE_CHARS: String(o.maxResponseChars) } : {}),
      ...(o.analyticsMaxRows ? { ODATA_ANALYTICS_MAX_ROWS: String(o.analyticsMaxRows) } : {}),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) => {
        const filter = q.params.get("$filter") ?? "";
        if (q.base === "Catalog_ВидыЦен" || q.base === "Catalog_ТипыЦенНоменклатуры") return KINDS;
        if (q.base === "Catalog_Номенклатура") {
          if (filter.includes("IsFolder eq true"))
            return [{ Ref_Key: G, Description: "Группа", Parent_Key: EMPTY, IsFolder: true }];
          const refs = new Set(guids(filter));
          // Чтение по ссылкам находит и папку: её отличает только IsFolder.
          if (filter.includes("Ref_Key eq"))
            return [
              ...items,
              { Ref_Key: G, Description: "Группа", Parent_Key: EMPTY, IsFolder: true },
            ].filter((i) => refs.has(i.Ref_Key));
          if (filter.includes("Parent_Key eq")) return items.filter((i) => refs.has(i.Parent_Key));
          return items;
        }
        return undefined;
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
    if (path.startsWith("Document_УстановкаЦенНоменклатуры")) {
      const ref = /\(guid'([0-9a-f-]{36})'\)/i.exec(path)?.[1]?.toLowerCase();
      if (method === "POST" && !ref) {
        const body = JSON.parse(String(init?.body)) as Record<string, unknown>;
        posts.push(String(body["Ref_Key"]));
        docs.set(String(body["Ref_Key"]).toLowerCase(), { body, posted: false, version: 1 });
        return json(201, { ...body, Posted: false, DataVersion: "v1", Number: "0000-000001" });
      }
      if (ref && /\/(Post|Unpost)\?/.test(path)) {
        const d = docs.get(ref)!;
        d.posted = path.includes("/Post?");
        if (!d.posted) unposted.add(ref);
        d.version++;
        return json(200, {});
      }
      if (ref && method === "GET") {
        const d = docs.get(ref);
        if (!d)
          return json(404, { "odata.error": { code: "9", message: { lang: "ru", value: "Не найден" } } });
        return json(200, {
          ...d.body,
          Ref_Key: ref,
          Posted: d.posted,
          DeletionMark: false,
          DataVersion: `v${d.version}`,
          Товары2_5: [],
        });
      }
      // Механизм цен: ссылки последних проведённых документов, затем строка из набора табличной части.
      const rows = o.recentRows ?? "Товары";
      if (path.startsWith("Document_УстановкаЦенНоменклатуры_Товары2_5?"))
        return json(200, { value: rows !== "Товары" ? [{ Ref_Key: uuid("f", 1) }] : [] });
      if (path.startsWith("Document_УстановкаЦенНоменклатуры_Товары?"))
        return json(200, { value: rows !== "Товары2_5" ? [{ Ref_Key: uuid("f", 1) }] : [] });
      if (method === "GET") return json(200, { value: [{ Ref_Key: uuid("f", 1) }] });
    }
    if (path.startsWith("InformationRegister_ЦеныНоменклатуры(Recorder=")) {
      const ref = /Recorder='([0-9a-f-]{36})'/i.exec(path)![1]!.toLowerCase();
      if (o.failMovements)
        return json(500, { "odata.error": { code: "-1", message: { lang: "ru", value: "Сбой" } } });
      const set = movementsOf(ref);
      // Как живая УТ 11.5: у документа без движений набора нет — 404, а не пустой набор.
      if (!set.length)
        return json(404, {
          "odata.error": { code: "9", message: { lang: "ru", value: "Экземпляр сущности не найден" } },
        });
      return json(200, { Recorder: ref, Recorder_Type: DOC_TYPE, RecordSet: set.map(out) });
    }
    if (path.startsWith("InformationRegister_ЦеныНоменклатуры25"))
      return json(200, {
        value: o.legacyEmpty && path.includes("$top=1") ? [{ Period: "2026-11-01T00:00:00" }] : [],
      });
    if (
      o.legacyEmpty &&
      path.startsWith("InformationRegister_ЦеныНоменклатуры_RecordType?") &&
      path.includes("$top=1")
    )
      return json(200, { value: [] });
    if (path.startsWith("InformationRegister_ЦеныНоменклатуры_RecordType")) {
      const all = guids(path);
      const kinds = new Set(all.filter((g) => KINDS.some((k) => k.Ref_Key === g)));
      const itemRefs = new Set(all.filter((g) => items.some((i) => i.Ref_Key === g)));
      const pick = (r: Rec) =>
        (!kinds.size || kinds.has(r.ВидЦены_Key)) && (!itemRefs.size || itemRefs.has(r.Номенклатура_Key));
      if (path.includes("/SliceLast(")) {
        const at = /Period=datetime'([^']+)'/.exec(path)?.[1];
        const last = new Map<string, Rec>();
        for (const r of records()
          .filter((x) => pick(x) && (!at || x.Period <= at))
          .sort((a, b) => a.Period.localeCompare(b.Period)))
          last.set(`${r.Номенклатура_Key}|${r.Характеристика_Key}|${r.ВидЦены_Key}`, r);
        const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
        return json(200, { value: skip ? [] : [...last.values()].map(out) });
      }
      const after = /Period ge datetime'([^']+)'/.exec(path)?.[1] ?? "";
      const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
      return json(200, {
        value: skip
          ? []
          : records()
              .filter((r) => pick(r) && r.Period >= after)
              .map(out),
      });
    }
    return inner(input, init);
  });

  const [a, b] = InMemoryTransport.createLinkedPair();
  void createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  const ready = client.connect(b);
  const call = async (name: string, args: Record<string, unknown>) => {
    await ready;
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, docs, posts, fake };
}

const SET = "write.pricing.create_price_setting";
const AT = "2026-11-20T12:00:00";
const lines = [
  { itemRef: I(1), price: 120 },
  { itemRef: I(2), price: 220.5 },
];

async function createAndPost(s: ReturnType<typeof setup>, extra: Record<string, unknown> = {}) {
  const preview = await s.call(SET, { priceKind: "Закупочная", lines, date: AT, ...extra });
  expect(preview.isError, preview.text.slice(0, 300)).toBeFalsy();
  const created = await s.call(SET, {
    priceKind: "Закупочная",
    lines,
    date: AT,
    ...extra,
    confirm: true,
    operationId: preview.structuredContent["operationId"],
  });
  expect(created.isError, created.text.slice(0, 300)).toBeFalsy();
  const ref = String(created.structuredContent["ref"]);
  const posted = await s.call("write.document.post_document", {
    entitySet: "Document_УстановкаЦенНоменклатуры",
    ref,
    post: true,
    confirm: true,
  });
  return { preview, created, ref, posted };
}

describe("установка цен: предпросмотр", () => {
  it("было → станет, зависимые виды, более поздняя цена; в 1С ничего не отправлено", async () => {
    const s = setup();
    const r = await s.call(SET, { priceKind: "Закупочная", lines, date: AT });
    expect(r.isError, r.text).toBeFalsy();
    const sc = r.structuredContent;
    expect(sc["dryRun"]).toBe(true);
    expect(sc["mechanism"]).toBe("legacy");
    const prices = sc["prices"] as Record<string, unknown>[];
    expect(prices[0]).toMatchObject({ ref: I(1), current: 100, price: 120, change: 20 });
    expect(prices[1]).toMatchObject({
      ref: I(2),
      current: 200,
      price: 220.5,
      laterPriceFrom: "2026-12-01T00:00:00",
    });
    const notes = (sc["notes"] as string[]).join(" ");
    expect(notes).toContain("Розничная"); // зависит от закупочной — не пересчитается
    expect(notes).toContain("более поздней датой");
    expect(s.posts).toEqual([]);
    const payload = sc["payload"] as Record<string, unknown>;
    expect(payload).toMatchObject({ Date: AT, Статус: "Согласован", Согласован: true });
    expect(payload["ВидыЦен"]).toEqual([{ LineNumber: 1, ВидЦены_Key: K_BASE }]);
    // Строки — только в prices: вдвоём с payload они не помещались в норму ответа (ревью).
    expect(payload["Товары"]).toBe("2 строк — см. prices");
  });

  it.each([
    ["повтор позиции", { lines: [lines[0], lines[0]] }, "повторяется"],
    ["три знака после запятой", { lines: [{ itemRef: I(1), price: 1.005 }] }, "двух знаков"],
    ["цена на тот же момент уже есть", { date: "2026-11-01T00:00:01" }, "уже есть цена этого вида"],
    ["вид цен не действует", { priceKind: "Старая" }, "не действует"],
    ["позиции нет в справочнике", { lines: [{ itemRef: uuid("a", 999), price: 1 }] }, "Нет в справочнике"],
    ["вид цен не найден", { priceKind: "Оптовая" }, "не найден"],
  ] as const)("отказ: %s", async (_name, patch, text) => {
    const s = setup();
    const r = await s.call(SET, { priceKind: "Закупочная", lines, date: AT, ...patch });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(text);
    expect(s.posts).toEqual([]);
  });
});

describe("установка цен: создание, проведение, сверка движений", () => {
  it("создаёт непроведённый документ, проведение даёт движения как в строках, срез — новые цены", async () => {
    const s = setup();
    const { created, ref, posted } = await createAndPost(s);
    expect(created.structuredContent).toMatchObject({ created: true, posted: false });
    // В 1С ушла строка, заполненная как в живых документах.
    expect((s.docs.get(ref.toLowerCase())!.body["Товары"] as unknown[])[0]).toMatchObject({
      Номенклатура_Key: I(1),
      Характеристика_Key: EMPTY,
      Упаковка_Key: EMPTY,
      ВидЦены_Key: K_BASE,
      Цена: 120,
      ЦенаИзмененаВручную: true,
      Валюта_Key: CUR,
    });
    expect(s.posts).toHaveLength(1);
    expect(posted.isError, posted.text).toBeFalsy();
    expect(posted.structuredContent).toMatchObject({
      done: true,
      verified: true,
      movementsVerified: true,
      movements: { mechanism: "legacy", expected: 2, found: 2 },
    });
    expect(posted.structuredContent).not.toHaveProperty("postings");
    const now = await s.call("read.analytics.get_prices", { priceKinds: ["Закупочная"], asOf: "2026-11-21" });
    const price = (r: string) =>
      (now.structuredContent["items"] as Record<string, unknown>[]).find((x) => x["ref"] === r);
    expect(price(I(1))).toMatchObject({ price: 120, since: AT, recorderRef: ref });
    expect(price(I(2))).toMatchObject({ price: 220.5 });
    // Другие виды цен и позиции не тронуты.
    const markup = await s.call("read.analytics.get_prices", {
      priceKinds: ["Розничная"],
      asOf: "2026-11-21",
    });
    expect(markup.structuredContent["items"]).toEqual([expect.objectContaining({ ref: I(1), price: 150 })]);
  });

  it("отмена проведения: движений нет, срез — прежние цены; потом пометка удаления проходит", async () => {
    const s = setup();
    const { ref } = await createAndPost(s);
    const markPosted = await s.call("write.entity.mark_for_deletion", {
      entitySet: "Document_УстановкаЦенНоменклатуры",
      ref,
      confirm: true,
    });
    expect(markPosted.isError).toBe(true);
    expect(markPosted.text).toContain("Сначала отмените проведение");
    const unposted = await s.call("write.document.post_document", {
      entitySet: "Document_УстановкаЦенНоменклатуры",
      ref,
      post: false,
      confirm: true,
    });
    expect(unposted.structuredContent).toMatchObject({
      done: true,
      movementsVerified: true,
      movements: { found: 0 },
    });
    const back = await s.call("read.analytics.get_prices", {
      priceKinds: ["Закупочная"],
      asOf: "2026-11-21",
    });
    expect(
      (back.structuredContent["items"] as Record<string, unknown>[]).find((x) => x["ref"] === I(1)),
    ).toMatchObject({
      price: 100,
    });
    const mark = await s.call("write.entity.mark_for_deletion", {
      entitySet: "Document_УстановкаЦенНоменклатуры",
      ref,
      confirm: true,
    });
    expect(mark.isError, mark.text).toBeFalsy();
  });

  it("движения не совпали со строками — done, но verified=false, расхождения и следующий шаг", async () => {
    const s = setup({ corruptPost: true });
    const { posted } = await createAndPost(s);
    expect(posted.structuredContent).toMatchObject({ done: true, verified: false, movementsVerified: false });
    expect((posted.structuredContent["movements"] as { mismatches: string[] }).mismatches).toHaveLength(2);
    expect(String(posted.structuredContent["note"])).toContain("post=false");
  });

  it("повтор подтверждения с тем же operationId не создаёт второй документ", async () => {
    const s = setup();
    const preview = await s.call(SET, { priceKind: "Закупочная", lines, date: AT });
    const args = {
      priceKind: "Закупочная",
      lines,
      date: AT,
      confirm: true,
      operationId: preview.structuredContent["operationId"],
    };
    const first = await s.call(SET, args);
    const again = await s.call(SET, args);
    expect(first.isError, first.text).toBeFalsy();
    expect(again.structuredContent).toMatchObject({ replayed: true, ref: first.structuredContent["ref"] });
    expect(s.posts).toHaveLength(1);
  });
});

describe("механизм цен и чтение", () => {
  it("последние документы заполняют Товары2_5 — механизм 2.5, с пометкой «не проверено»", async () => {
    const s = setup({ recentRows: "Товары2_5" });
    const r = await s.call(SET, { priceKind: "Закупочная", lines, date: AT });
    expect(r.structuredContent["mechanism"]).toBe("2.5");
    expect((r.structuredContent["notes"] as string[]).join(" ")).toContain("не проверено");
    const payload = r.structuredContent["payload"] as Record<string, unknown>;
    expect(payload["Товары2_5"]).toBe("2 строк — см. prices");
    expect(payload).not.toHaveProperty("Товары");
  });

  it("заполнены обе табличные части, записи есть только в прежнем регистре — legacy", async () => {
    const s = setup({ recentRows: "both" });
    const r = await s.call(SET, { priceKind: "Закупочная", lines, date: AT });
    expect(r.structuredContent["mechanism"]).toBe("legacy");
  });

  it("цены на момент: до установки — прежние, позиции группы, отбор по виду цен", async () => {
    const s = setup();
    const r = await s.call("read.analytics.get_prices", { asOf: "2026-11-15", group: "Группа" });
    expect(r.isError, r.text).toBeFalsy();
    const items = r.structuredContent["items"] as Record<string, unknown>[];
    // Не действующий вид цен в ответ без явного запроса не входит; позиция 3 без цен — тоже.
    expect(items.map((x) => `${x["ref"]}|${x["priceKind"]}|${x["price"]}`).sort()).toEqual(
      [`${I(1)}|Закупочная|100`, `${I(1)}|Розничная|150`, `${I(2)}|Закупочная|200`].sort(),
    );
    expect(r.structuredContent["count"]).toBe(3);
  });

  it("Бухгалтерия: цены по типам цен — тот же инструмент, тип цен в шапке, проведение сверяется", async () => {
    const s = setup({ schema: "bp30" });
    const { preview, created, ref, posted } = await createAndPost(s);
    expect(preview.structuredContent["mechanism"]).toBe("bp");
    const payload = preview.structuredContent["payload"] as Record<string, unknown>;
    expect(payload).toMatchObject({ Date: AT, ТипЦен_Key: K_BASE });
    expect(payload).not.toHaveProperty("ВидыЦен");
    expect(payload).not.toHaveProperty("Статус");
    expect(created.isError, created.text).toBeFalsy();
    const row = (s.docs.get(ref.toLowerCase())!.body["Товары"] as Record<string, unknown>[])[0]!;
    expect(Object.keys(row).sort()).toEqual(["LineNumber", "Валюта_Key", "Номенклатура_Key", "Цена"]);
    expect(posted.structuredContent).toMatchObject({ done: true, verified: true, movementsVerified: true });
    const prices = await s.call("read.analytics.get_prices", {
      priceKinds: ["Закупочная"],
      asOf: "2026-11-21",
    });
    expect(prices.structuredContent["mechanism"]).toBe("bp");
    expect(
      (prices.structuredContent["items"] as Record<string, unknown>[]).find((x) => x["ref"] === I(1)),
    ).toMatchObject({
      price: 120,
    });
  });

  it("Бухгалтерия: характеристики в строках — отказ, их там нет", async () => {
    const s = setup({ schema: "bp30" });
    const r = await s.call(SET, {
      priceKind: "Закупочная",
      lines: [{ itemRef: I(1), price: 1, characteristicRef: uuid("c", 9) }],
      date: AT,
    });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("без характеристик");
  });
});

describe("находки ревью и живой проверки", () => {
  it("документ без движений: 1С отвечает 404 — это пустой набор, после отмены проведения сверка проходит", async () => {
    const s = setup();
    const { ref } = await createAndPost(s);
    const unposted = await s.call("write.document.post_document", {
      entitySet: "Document_УстановкаЦенНоменклатуры",
      ref,
      post: false,
      confirm: true,
    });
    expect(unposted.structuredContent).toMatchObject({
      verified: true,
      movementsVerified: true,
      movements: { found: 0 },
    });
  });

  it.each([
    ["лишняя запись в регистре", { extraMovement: true }, true, "лишних записей"],
    [
      "записи остались после отмены проведения",
      { keepAfterUnpost: true },
      false,
      "после отмены проведения осталось",
    ],
  ] as const)("%s — verified=false", async (_name, opts, post, text) => {
    const s = setup(opts);
    const { ref, posted } = await createAndPost(s);
    const r = post
      ? posted
      : await s.call("write.document.post_document", {
          entitySet: "Document_УстановкаЦенНоменклатуры",
          ref,
          post: false,
          confirm: true,
        });
    expect(r.structuredContent).toMatchObject({ done: true, verified: false, movementsVerified: false });
    expect(
      ((r.structuredContent["movements"] as { mismatches: string[] }).mismatches ?? []).join(" "),
    ).toContain(text);
  });

  it("движения не прочитались — не успех: verified=false и совет проверить в 1С", async () => {
    const s = setup({ failMovements: true });
    const { posted } = await createAndPost(s);
    expect(posted.structuredContent).toMatchObject({ done: true, verified: false, movementsVerified: false });
    expect(String(posted.structuredContent["note"])).toContain("прочитать не удалось");
  });

  it("проведение с expectedDataVersion тоже сверяет движения", async () => {
    const s = setup({ corruptPost: true });
    const preview = await s.call(SET, { priceKind: "Закупочная", lines, date: AT });
    const created = await s.call(SET, {
      priceKind: "Закупочная",
      lines,
      date: AT,
      confirm: true,
      operationId: preview.structuredContent["operationId"],
    });
    const ref = String(created.structuredContent["ref"]);
    const args = {
      entitySet: "Document_УстановкаЦенНоменклатуры",
      ref,
      post: true,
      expectedDataVersion: "v1",
    };
    const dry = await s.call("write.document.post_document", args);
    expect(dry.isError, dry.text.slice(0, 300)).toBeFalsy();
    const done = await s.call("write.document.post_document", {
      ...args,
      confirm: true,
      operationId: dry.structuredContent["operationId"],
    });
    expect(done.structuredContent).toMatchObject({ done: true, verified: false, movementsVerified: false });
  });

  it("предпросмотр на 80 строк укладывается в норму ответа и несёт operationId; 1001 строка — отказ", async () => {
    const s = setup();
    const many = ITEMS.slice(0, 80).map((i) => ({ itemRef: i.Ref_Key, price: 99.99 }));
    const r = await s.call(SET, { priceKind: "Розничная", lines: many, date: AT });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.text.length).toBeLessThan(45_000);
    expect(r.text).toContain('"operationId"');
    expect(r.structuredContent["prices"]).toHaveLength(80);
    const c = scaleCase(1001, 10, 0);
    const big = setup({ items: c.items, base: c.base });
    const over = await big.call(SET, { priceKind: "Закупочная", lines: c.lines, date: AT });
    expect(over.isError).toBe(true);
    expect(big.posts).toEqual([]);
  });

  it("папка номенклатуры вместо позиции — отказ", async () => {
    const s = setup();
    const r = await s.call(SET, { priceKind: "Закупочная", lines: [{ itemRef: G, price: 1 }], date: AT });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("папки");
  });

  it.each(["write.document.post_document", "write.entity.mark_for_deletion"])(
    "%s: ref не GUID — отказ на входе, в 1С ничего не уходит",
    async (name) => {
      const s = setup();
      const r = await s.call(name, {
        entitySet: "Document_УстановкаЦенНоменклатуры",
        ref: `${I(1)}/../Catalog_Номенклатура(guid'${I(1)}')?$format=json&x=`,
        confirm: true,
      });
      expect(r.isError).toBe(true);
      expect(s.posts).toEqual([]);
    },
  );
});

describe("пометка удаления документа цен в Бухгалтерии", () => {
  const ref = "11111111-1111-4111-8111-111111111111";
  const mark = async (withPriceSchema: boolean) => {
    const fake = fake1C(
      { READ_ONLY: "false", ODATA_WRITABLE: "true" },
      {
        data: (q) =>
          q.base === "Document_УстановкаЦенНоменклатуры"
            ? [{ Ref_Key: ref, Posted: true, DeletionMark: false }]
            : undefined,
      },
    );
    if (!withPriceSchema) {
      // Регистр цен не опубликован: устройство цен не подтверждено — пометка как раньше.
      const meta = await fake.conn.getMetadata();
      const entities = new Map(
        [...meta.entities].filter(([k]) => !k.startsWith("InformationRegister_ЦеныНоменклатуры")),
      );
      vi.spyOn(fake.conn, "getMetadata").mockResolvedValue({ ...meta, entities });
    }
    const tools = (
      createServer(fake.ctx) as unknown as {
        _registeredTools: Record<
          string,
          {
            inputSchema: { parse: (a: unknown) => Record<string, unknown> };
            handler: (
              a: unknown,
              e: unknown,
            ) => Promise<{ isError?: boolean; content?: { text?: string }[] }>;
          }
        >;
      }
    )._registeredTools;
    const t = tools["write.entity.mark_for_deletion"]!;
    const r = await t.handler(
      t.inputSchema.parse({ entitySet: "Document_УстановкаЦенНоменклатуры", ref }),
      {},
    );
    return { r, text: (r.content ?? []).map((c) => c.text).join("") };
  };

  it("схема цен опубликована: проведённый документ не помечается — его цены продолжали бы действовать", async () => {
    const { r, text } = await mark(true);
    expect(r.isError).toBe(true);
    expect(text).toContain("Сначала отмените проведение");
  });

  it("схемы цен нет: помечается как раньше (предпросмотр)", async () => {
    const { r, text } = await mark(false);
    expect(r.isError, text).toBeFalsy();
  });
});

describe("пробелы тестов из ревью", () => {
  it("неактивная запись регистра (Active=false) не считается лишней", async () => {
    const s = setup({ inactiveMovement: true });
    const { posted } = await createAndPost(s);
    expect(posted.structuredContent).toMatchObject({
      verified: true,
      movementsVerified: true,
      movements: { found: 2 },
    });
  });

  it("обе табличные части заполнены — механизм по регистру с записями (здесь 2.5)", async () => {
    const s = setup({ recentRows: "both", legacyEmpty: true });
    const r = await s.call(SET, { priceKind: "Закупочная", lines, date: AT });
    expect(r.structuredContent["mechanism"]).toBe("2.5");
  });

  it("страницы get_prices: продолжение с другими видами цен не берёт чужое чтение", async () => {
    const s = setup();
    const first = await s.call("read.analytics.get_prices", {
      priceKinds: ["Закупочная"],
      asOf: "2026-11-15",
      limit: 1,
    });
    expect(first.structuredContent["count"]).toBe(2);
    const other = await s.call("read.analytics.get_prices", {
      priceKinds: ["Розничная"],
      asOf: "2026-11-15",
      limit: 1,
      offset: 1,
    });
    expect(other.structuredContent["count"]).toBe(1);
  });
});

// 1.1: момент цены по умолчанию брался от «сейчас» при подтверждении. Предпросмотр до полуночи и подтверждение (или
// безопасный повтор после неизвестного исхода) после неё давали другой день — отказ «данные не совпадают с dry-run».
// Старые тесты всегда передавали date явно.
describe("момент цены по умолчанию — время предпросмотра", () => {
  it("подтверждение через сутки без date: документ с моментом предпросмотра, без отказа", async () => {
    const s = setup();
    const preview = await s.call(SET, { priceKind: "Закупочная", lines });
    expect(preview.isError, preview.text.slice(0, 300)).toBeFalsy();
    const day = String((preview.structuredContent["payload"] as Record<string, unknown>)["Date"]).slice(
      0,
      10,
    );
    vi.useFakeTimers({ toFake: ["Date"], now: Date.now() + 24 * 3600_000 });
    try {
      const created = await s.call(SET, {
        priceKind: "Закупочная",
        lines,
        confirm: true,
        operationId: preview.structuredContent["operationId"],
      });
      expect(created.isError, created.text.slice(0, 300)).toBeFalsy();
      const ref = String(created.structuredContent["ref"]).toLowerCase();
      expect(String(s.docs.get(ref)!.body["Date"]).slice(0, 10)).toBe(day);
    } finally {
      vi.useRealTimers();
    }
  });
});

/**
 * Корпус масштаба: число строк, длина названий, доля отклонений, предел ответа. PRICE_DUMP=<каталог> — записать
 * нормализованные ответы и запросы в 1С для сравнения версий (scratchpad), без него — ничего не пишет.
 */
const BIG = Array.from({ length: 1001 }, (_, n) => n + 1);
function scaleCase(n: number, nameLen: number, share: number) {
  const items = BIG.map((i) => ({
    Ref_Key: I(i),
    Description: `Позиция ${i} `.padEnd(nameLen, "ж"),
    Артикул: `A-${i}`,
    Parent_Key: G,
    IsFolder: false,
    DeletionMark: share > 0 && i % 13 === 0,
  }));
  const base: Rec[] = [];
  const lines = BIG.slice(0, n).map((i, j) => {
    const noCurrent = share > 0 && i % 5 === 0;
    if (!noCurrent) base.push(rec(I(i), K_BASE, share > 0 && i % 11 === 0 ? 0 : 100, "2026-11-01T00:00:01"));
    if (share > 0 && i % 9 === 0) base.push(rec(I(i), K_BASE, 101, "2026-12-01T00:00:00"));
    const dev = share === 1 || (share > 0 && j % 10 === 0);
    return { itemRef: I(i), price: dev ? (j % 2 ? 150 : 70) : j % 10 === 3 ? 120 : 105 };
  });
  return { items, base, lines };
}

const normalize = (v: unknown): unknown => {
  const ids = new Map<string, string>();
  return JSON.parse(
    JSON.stringify(v).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, (g) =>
      /^a0000000-/.test(g) || /^[cdef]0000000-/.test(g) || /^0{8}-/.test(g)
        ? g
        : (ids.get(g) ?? ids.set(g, `id${ids.size}`).get(g)!),
    ),
  );
};

describe("корпус масштаба: дамп для сравнения версий", () => {
  const dir = process.env["PRICE_DUMP"];
  const cases: Array<[number, number, number, number]> = [];
  for (const n of (process.env["PRICE_DUMP_N"] ?? "1,2,40,80").split(",").map(Number))
    for (const len of [10, 100, 150])
      for (const share of [0, 0.1, 1])
        for (const limit of [45_000, 10_000]) cases.push([n, len, share, limit]);
  it.runIf(!!dir).each(cases)("n=%i длина=%i доля=%f предел=%i", async (n, len, share, limit) => {
    const c = scaleCase(n, len, share);
    const s = setup({ items: c.items, base: c.base, maxResponseChars: limit });
    const gets: string[] = [];
    const inner = globalThis.fetch;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
      gets.push(`${init?.method ?? "GET"} ${decodeURIComponent(url.replace(/^https?:\/\/[^/]+/, ""))}`);
      return inner(input, init);
    });
    const p = await s.call(SET, { priceKind: "Закупочная", lines: c.lines, date: AT });
    const cf = await s.call(SET, {
      priceKind: "Закупочная",
      lines: c.lines,
      date: AT,
      confirm: true,
      operationId: p.structuredContent["operationId"],
    });
    const { writeFileSync, mkdirSync } = await import("node:fs");
    mkdirSync(dir!, { recursive: true });
    writeFileSync(
      join(dir!, `n${n}-l${len}-s${share}-m${limit}.json`),
      JSON.stringify(
        normalize({
          preview: { len: p.text.length, isError: !!p.isError, sc: p.structuredContent },
          confirm: { len: cf.text.length, isError: !!cf.isError, sc: cf.structuredContent },
          gets,
        }),
        null,
        1,
      ),
    );
  });
});

/** Ожидание по данным случая, независимо от кода: действующая цена на AT, более поздняя, пометка, порог в копейках. */
function expected(c: ReturnType<typeof scaleCase>) {
  const del = new Set(c.items.filter((i) => i.DeletionMark).map((i) => i.Ref_Key));
  const rows = c.lines.map((l) => {
    const recs = c.base.filter((r) => r.Номенклатура_Key === l.itemRef && r.ВидЦены_Key === K_BASE);
    const cur = recs
      .filter((r) => r.Period <= AT)
      .sort((a, b) => a.Period.localeCompare(b.Period))
      .pop();
    const later = recs.some((r) => r.Period >= AT);
    const c100 = cur ? Math.round(cur.Цена * 100) : undefined;
    const over = c100 !== undefined && Math.abs(Math.round(l.price * 100) - c100) * 5 > c100;
    return {
      cur: c100,
      later,
      del: del.has(l.itemRef),
      over,
      dev: c100 === undefined || later || del.has(l.itemRef) || over,
    };
  });
  const n = (f: (r: (typeof rows)[number]) => boolean) => rows.filter(f).length;
  return {
    lines: rows.length,
    withoutCurrent: n((r) => r.cur === undefined),
    overThreshold: n((r) => r.over),
    laterPrices: n((r) => r.later),
    deletionMarked: n((r) => r.del),
    deviations: n((r) => r.dev),
  };
}

describe("масштаб: до 1000 строк, ответ в пределе, сводка по всем строкам", () => {
  const cases: Array<[number, number, number, number]> = [];
  for (const n of [40, 81, 300, 1000])
    for (const len of [10, 150])
      for (const share of [0, 0.1, 1])
        for (const limit of [45_000, 10_000]) cases.push([n, len, share, limit]);
  it.each(cases)("n=%i длина=%i доля=%f предел=%i", async (n, len, share, limit) => {
    const c = scaleCase(n, len, share);
    const s = setup({ items: c.items, base: c.base, maxResponseChars: limit });
    const gets: string[] = [];
    const inner = globalThis.fetch;
    vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
      const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
      if ((init?.method ?? "GET") === "GET") gets.push(decodeURIComponent(url));
      return inner(input, init);
    });
    const p = await s.call(SET, { priceKind: "Закупочная", lines: c.lines, date: AT });
    expect(p.isError, p.text.slice(0, 300)).toBeFalsy();
    const exp = expected(c);
    for (const r of [p]) {
      expect(r.text.length).toBeLessThanOrEqual(limit);
      expect(r.structuredContent["summary"]).toMatchObject(exp);
      const prices = r.structuredContent["prices"] as Array<Record<string, unknown>>;
      if (r.structuredContent["pricesComplete"] === false) {
        expect(r.structuredContent).toMatchObject({
          pricesTotal: n,
          pricesShown: prices.length,
          deviationsTotal: exp.deviations,
        });
        const devLines = new Set(
          c.lines.flatMap((_, i) => (expected({ ...c, lines: [c.lines[i]!] }).deviations ? [i + 1] : [])),
        );
        for (const row of prices) expect(devLines.has(Number(row["line"]))).toBe(true);
        // Пометка удаления и более поздняя цена — первыми.
        const rank = (row: Record<string, unknown>) =>
          row["deletionMark"] ? 0 : row["laterPriceFrom"] ? 1 : 2;
        expect(prices.map(rank)).toEqual([...prices.map(rank)].sort((a, b) => a - b));
        if (prices.length < exp.deviations)
          expect(String(r.structuredContent["notes"])).toContain("самые важные первыми");
      } else expect(prices).toHaveLength(n);
    }
    expect(p.structuredContent["operationId"]).toEqual(expect.any(String));
    // Больше 80 строк — цены по виду цен целиком, без пачек по ссылкам. Больше 200 — и справочник целиком
    // (readPageAttributes): 1000 строк — не больше 15 запросов вместо 150.
    if (n > 80) {
      expect(gets.filter((g) => /ЦеныНоменклатуры.*Номенклатура_Key eq/.test(g))).toEqual([]);
      // Вид цен целиком — только поля цены: со всеми полями на рабочей УТ 20,8 МБ вместо 1,4.
      const wide = gets.filter((g) => /ЦеныНоменклатуры_RecordType(\/SliceLast\(|\?.*Period ge)/.test(g));
      expect(wide).toHaveLength(2);
      for (const g of wide)
        expect(g).toMatch(/\$select=Номенклатура_Key,ВидЦены_Key,Характеристика_Key,Цена,Period/);
    }
    if (n > 200) {
      expect(gets.filter((g) => /Catalog_Номенклатура.*Ref_Key eq/.test(g))).toEqual([]);
      expect(gets.length).toBeLessThanOrEqual(15);
    }
    const cf = await s.call(SET, {
      priceKind: "Закупочная",
      lines: c.lines,
      date: AT,
      confirm: true,
      operationId: p.structuredContent["operationId"],
    });
    expect(cf.isError, cf.text.slice(0, 300)).toBeFalsy();
    expect(cf.text.length).toBeLessThanOrEqual(limit);
    expect(cf.structuredContent["ref"]).toEqual(expect.any(String));
    expect(cf.structuredContent["summary"]).toMatchObject(exp);
    expect(s.posts).toHaveLength(1);
    const body = [...s.docs.values()][0]!.body;
    expect(body["Товары"]).toHaveLength(n);
  });

  it("100 строк с длинными названиями: prices полный, строки документа не повторены в view", async () => {
    const c = scaleCase(100, 150, 0.1);
    const s = setup({ items: c.items, base: c.base });
    const r = await s.call(SET, { priceKind: "Закупочная", lines: c.lines, date: AT });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.text.length).toBeLessThanOrEqual(45_000);
    expect(r.structuredContent["prices"]).toHaveLength(100);
    expect((r.structuredContent["prices"] as Array<Record<string, unknown>>)[0]).toHaveProperty("ref");
    expect(r.structuredContent["pricesComplete"]).toBeUndefined();
    const view = r.structuredContent["view"] as { tables: Record<string, unknown[]> };
    expect(Object.values(view.tables).some((t) => t.length === 100)).toBe(false);
    expect(String(r.structuredContent["notes"])).toContain("view не повторены");
  });

  it("документ на 1000 строк: проведение, отмена, чтение, статус операции — в пределе ответа", async () => {
    const c = scaleCase(1000, 150, 1);
    const s = setup({ items: c.items, base: c.base, maxResponseChars: 10_000 });
    const args = { priceKind: "Закупочная", lines: c.lines, date: AT };
    const p = await s.call(SET, args);
    const operationId = p.structuredContent["operationId"];
    const cf = await s.call(SET, { ...args, confirm: true, operationId });
    const ref = String(cf.structuredContent["ref"]);
    const doc = { entitySet: "Document_УстановкаЦенНоменклатуры", ref };
    const calls: Array<[string, Record<string, unknown>]> = [
      ["write.document.post_document", { ...doc, post: true }],
      ["write.document.post_document", { ...doc, post: true, confirm: true }],
      ["write.document.post_document", { ...doc, post: false, confirm: true }],
      ["read.document.get_document", doc],
      ["write.operation.status", { operationId }],
      [SET, { ...args, confirm: true, operationId }],
    ];
    for (const [name, a] of calls) {
      const r = await s.call(name, a);
      expect(r.isError, `${name}: ${r.text.slice(0, 300)}`).toBeFalsy();
      expect(r.text.length, name).toBeLessThanOrEqual(10_000);
      if (name === "write.document.post_document" && a["confirm"])
        expect(r.structuredContent["movementsVerified"], r.text.slice(0, 300)).toBe(true);
    }
    expect(s.posts).toHaveLength(1);
  });

  it("порог 20 %: ровно 20 % — не отклонение, больше — отклонение, текущая 0 — отклонение", async () => {
    const base = [1, 2, 3, 4, 5].map((i) => rec(I(i), K_BASE, 100, "2026-11-01T00:00:01"));
    base.push(rec(I(6), K_BASE, 0, "2026-11-01T00:00:01"), rec(I(7), K_BASE, 30, "2026-11-01T00:00:01"));
    const s = setup({ base });
    const prices = [120, 120.01, 80, 79, 100, 5, 36];
    const r = await s.call(SET, {
      priceKind: "Закупочная",
      lines: prices.map((price, i) => ({ itemRef: I(i + 1), price })),
      date: AT,
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(r.structuredContent["summary"]).toEqual({
      lines: 7,
      withoutCurrent: 0,
      up: 4,
      down: 2,
      unchanged: 1,
      deviationThresholdPct: 20,
      overThreshold: 3,
      maxIncreasePct: 20,
      maxDecreasePct: -21,
      laterPrices: 0,
      deletionMarked: 0,
      deviations: 3,
    });
  });

  it("у вида цен записей больше предела чтения — те же цены чтением по ссылкам", async () => {
    const c = scaleCase(100, 10, 0.1);
    const wide = setup({ items: c.items, base: c.base });
    const capped = setup({ items: c.items, base: c.base, analyticsMaxRows: 20 });
    const args = { priceKind: "Закупочная", lines: c.lines, date: AT };
    const a = await wide.call(SET, args);
    const b = await capped.call(SET, args);
    expect(b.isError, b.text.slice(0, 300)).toBeFalsy();
    expect(b.structuredContent["prices"]).toEqual(a.structuredContent["prices"]);
    expect(b.structuredContent["summary"]).toEqual(a.structuredContent["summary"]);
  });
});
