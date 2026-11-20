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

const ITEMS = [1, 2, 3].map((i) => ({
  Ref_Key: I(i),
  Description: `Позиция ${i}`,
  Артикул: `A-${i}`,
  Parent_Key: G,
  IsFolder: false,
  DeletionMark: false,
}));

interface Rec {
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
  /** Табличная часть последних проведённых документов: legacy или 2.5. */
  recentRows?: "Товары" | "Товары2_5" | "both";
  schema?: FakeSchema;
}

function setup(o: Options = {}) {
  const docs = new Map<string, { body: Record<string, unknown>; posted: boolean; version: number }>();
  const base: Rec[] = [
    rec(I(1), K_BASE, 100, "2026-11-01T00:00:01"),
    rec(I(1), K_MARKUP, 150, "2026-11-01T00:00:01"),
    rec(I(2), K_BASE, 200, "2026-11-01T00:00:01"),
    rec(I(2), K_BASE, 210, "2026-12-01T00:00:00"), // установлена на будущее
  ];
  const posts: string[] = [];
  const movementsOf = (ref: string): Rec[] => {
    const d = docs.get(ref);
    if (!d?.posted) return [];
    return ((d.body["Товары"] as Record<string, unknown>[]) ?? []).map((r) => ({
      Period: String(d.body["Date"]),
      Номенклатура_Key: String(r["Номенклатура_Key"]),
      Характеристика_Key: String(r["Характеристика_Key"]),
      ВидЦены_Key: String(r["ВидЦены_Key"]),
      Цена: Number(r["Цена"]) + (o.corruptPost ? 1 : 0),
      Recorder: ref,
    }));
  };
  const records = () => [...base, ...[...docs.keys()].flatMap(movementsOf)];
  const guids = (text: string) =>
    [...text.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());

  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-pricing-")),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) => {
        const filter = q.params.get("$filter") ?? "";
        if (q.base === "Catalog_ВидыЦен") return KINDS;
        if (q.base === "Catalog_Номенклатура") {
          if (filter.includes("IsFolder eq true"))
            return [{ Ref_Key: G, Description: "Группа", Parent_Key: EMPTY, IsFolder: true }];
          const refs = new Set(guids(filter));
          if (filter.includes("Ref_Key eq")) return ITEMS.filter((i) => refs.has(i.Ref_Key));
          if (filter.includes("Parent_Key eq")) return ITEMS.filter((i) => refs.has(i.Parent_Key));
          return ITEMS;
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
          DataVersion: `v${d.version}`,
          Товары2_5: [],
        });
      }
      if (method === "GET") {
        // Последние проведённые документы — по ним определяется механизм цен.
        const rows = o.recentRows ?? "Товары";
        const sample = { Товары: rows !== "Товары2_5" ? [{}] : [], Товары2_5: rows !== "Товары" ? [{}] : [] };
        return json(200, { value: [sample] });
      }
    }
    if (path.startsWith("InformationRegister_ЦеныНоменклатуры(Recorder=")) {
      const ref = /Recorder='([0-9a-f-]{36})'/i.exec(path)![1]!.toLowerCase();
      return json(200, { Recorder: ref, Recorder_Type: DOC_TYPE, RecordSet: movementsOf(ref) });
    }
    if (path.startsWith("InformationRegister_ЦеныНоменклатуры25")) return json(200, { value: [] });
    if (path.startsWith("InformationRegister_ЦеныНоменклатуры_RecordType")) {
      const all = guids(path);
      const kinds = new Set(all.filter((g) => KINDS.some((k) => k.Ref_Key === g)));
      const items = new Set(all.filter((g) => ITEMS.some((i) => i.Ref_Key === g)));
      const pick = (r: Rec) =>
        (!kinds.size || kinds.has(r.ВидЦены_Key)) && (!items.size || items.has(r.Номенклатура_Key));
      if (path.includes("/SliceLast(")) {
        const at = /Period=datetime'([^']+)'/.exec(path)?.[1];
        const last = new Map<string, Rec>();
        for (const r of records()
          .filter((x) => pick(x) && (!at || x.Period <= at))
          .sort((a, b) => a.Period.localeCompare(b.Period)))
          last.set(`${r.Номенклатура_Key}|${r.Характеристика_Key}|${r.ВидЦены_Key}`, r);
        const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
        return json(200, { value: skip ? [] : [...last.values()] });
      }
      const after = /Period ge datetime'([^']+)'/.exec(path)?.[1] ?? "";
      const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
      return json(200, { value: skip ? [] : records().filter((r) => pick(r) && r.Period >= after) });
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
    expect((payload["Товары"] as unknown[])[0]).toMatchObject({
      Номенклатура_Key: I(1),
      Характеристика_Key: EMPTY,
      Упаковка_Key: EMPTY,
      ВидЦены_Key: K_BASE,
      Цена: 120,
      ЦенаИзмененаВручную: true,
      Валюта_Key: CUR,
    });
  });

  it.each([
    ["повтор позиции", { lines: [lines[0], lines[0]] }, "повторяется"],
    ["три знака после запятой", { lines: [{ itemRef: I(1), price: 1.005 }] }, "двух знаков"],
    ["цена на тот же момент уже есть", { date: "2026-11-01T00:00:01" }, "уже есть цена этого вида"],
    ["вид цен не действует", { priceKind: "Старая" }, "не действует"],
    ["позиции нет в справочнике", { lines: [{ itemRef: uuid("a", 99), price: 1 }] }, "Нет в справочнике"],
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
    expect((r.structuredContent["payload"] as Record<string, unknown>)["Товары2_5"]).toHaveLength(2);
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

  it("в Бухгалтерии чтение цен отказывает с причиной", async () => {
    const s = setup({ schema: "bp30" });
    const r = await s.call("read.analytics.get_prices", {});
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Управления торговлей");
  });
});
