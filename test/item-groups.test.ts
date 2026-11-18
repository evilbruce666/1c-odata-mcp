import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { PAGE_READ_ENTRIES, PAGE_READ_TTL_MS, pageCached } from "../src/odata/cache.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/**
 * Номенклатура УТ на синтетическом корпусе в 6000 позиций: список справочника (list_items), отбор по нескольким группам
 * за вызов (groups) и листание длинного списка остатков и продаж. Метрика приёмки:
 * - list_items отдаёт все позиции группы, в том числе без остатка и продаж;
 * - groups даёт те же позиции и итоги, что вызовы по одной группе, без задвоения у вложенных групп;
 * - склеенные страницы равны count, без повторов, а регистр за всё листание читается один раз.
 * Имена выдуманы, схема — снимок УТ 11.5.
 */

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const G = (i: number) => uuid("d", i);
const I = (i: number) => uuid("a", i);
const W = (i: number) => uuid("b", i);
const C = (i: number) => uuid("c", i);
const EMPTY = "00000000-0000-0000-0000-000000000000";

// «Бренд А» — в двух ветках; «Смартфоны» под «Распродажей» без номера, под «Техникой» с номером.
const GROUPS = [
  { Ref_Key: G(1), Description: "Техника", Parent_Key: EMPTY, Code: "Г1" },
  { Ref_Key: G(2), Description: "01. Смартфоны", Parent_Key: G(1), Code: "Г2" },
  { Ref_Key: G(3), Description: "Бренд А", Parent_Key: G(2), Code: "Г3" },
  { Ref_Key: G(4), Description: "Бренд Б", Parent_Key: G(2), Code: "Г4" },
  { Ref_Key: G(5), Description: "Планшеты", Parent_Key: G(1), Code: "Г5" },
  { Ref_Key: G(6), Description: "Распродажа", Parent_Key: EMPTY, Code: "Г6" },
  { Ref_Key: G(7), Description: "Смартфоны", Parent_Key: G(6), Code: "Г7" },
  { Ref_Key: G(8), Description: "Бренд А", Parent_Key: G(7), Code: "Г8" },
  { Ref_Key: G(9), Description: "Пустая", Parent_Key: EMPTY, Code: "Г9" },
  { Ref_Key: G(10), Description: "Старое", Parent_Key: G(1), Code: "Г10", DeletionMark: true },
].map((g) => ({ DeletionMark: false, ...g, IsFolder: true }));

const N = 6000;
const HOME = [G(3), G(4), G(5), G(8), G(10)];
/** Позиция i: папка по кругу, каждая 50-я помечена на удаление, характеристики — у каждой 4-й. */
const ITEMS = Array.from({ length: N }, (_, i) => ({
  Ref_Key: I(i),
  Description: `Позиция синтетическая ${i}`,
  Артикул: `ART-${i}`,
  Code: `К-${String(i).padStart(5, "0")}`,
  Parent_Key: HOME[i % HOME.length]!,
  DeletionMark: i % 50 === 0,
  ИспользованиеХарактеристик: i % 4 === 0 ? "ИндивидуальныеДляНоменклатуры" : "НеИспользовать",
  IsFolder: false,
}));

/** Остаток есть у 4 позиций из 5: у каждой 5-й движения нет — её видно только в list_items. */
const hasStock = (i: number) => i % 5 !== 0;
const bal = (item: string, inStock: number, toShip: number, ch = EMPTY, wh = W(1)) => ({
  Номенклатура_Key: item,
  Характеристика_Key: ch,
  Склад_Key: wh,
  Назначение_Key: EMPTY,
  Помещение_Key: EMPTY,
  Серия_Key: EMPTY,
  КонтролируемаяПартия_Key: EMPTY,
  ВНаличииBalance: inStock,
  КОтгрузкеBalance: toShip,
});
const BALANCE = ITEMS.flatMap((it, i) =>
  !hasStock(i)
    ? []
    : i % 4 === 0
      ? [bal(it.Ref_Key, 1, 0, C(1)), bal(it.Ref_Key, (i % 7) + 1, i % 3 === 0 ? 1 : 0, C(2), W(2))]
      : [bal(it.Ref_Key, (i % 7) + 1, i % 3 === 0 ? 1 : 0)],
);

/** Продажи у 3 позиций из 4, у каждой 9-й ещё и возврат. */
const sale = (item: string, quantity: number, operation = "РеализацияВРозницу") => ({
  АналитикаУчетаНоменклатуры_Key: uuid("e", 1),
  АналитикаУчетаНоменклатуры: { Номенклатура_Key: item, Характеристика_Key: EMPTY },
  Склад: W(1),
  Склад_Type: "StandardODATA.Catalog_Склады",
  ХозяйственнаяОперация: operation,
  КоличествоTurnover: quantity,
  СуммаВыручкиTurnover: Math.round(quantity * 100.1 * 100) / 100,
  СуммаВыручкиБезНДСTurnover: Math.round((quantity * 100.1 * 100) / 1.2) / 100,
});
const SALES = ITEMS.flatMap((it, i) =>
  i % 4 === 0
    ? []
    : [
        sale(it.Ref_Key, (i % 5) + 1),
        ...(i % 9 === 0 ? [sale(it.Ref_Key, -1, "ВозвратТоваровОтКлиентаПрошлыхПериодов")] : []),
      ],
);

const guidsIn = (text: string) =>
  new Set([...text.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase()));

function data(counter: { balance: number; turnovers: number; wholeCatalog: number }, balance = BALANCE) {
  return (q: FakeQuery) => {
    const filter = q.params.get("$filter") ?? "";
    if (q.fn === "Balance") {
      counter.balance++;
      return balance;
    }
    if (q.fn === "Turnovers") {
      counter.turnovers++;
      return SALES;
    }
    if (q.base === "Document_ОтчетОРозничныхПродажах") return [];
    if (q.base === "Catalog_Номенклатура") {
      if (filter.includes("IsFolder eq true")) return GROUPS;
      const refs = guidsIn(filter);
      if (filter.includes("Parent_Key eq")) return ITEMS.filter((i) => refs.has(i.Parent_Key));
      if (filter.includes("Ref_Key eq")) return [...ITEMS, ...GROUPS].filter((i) => refs.has(i.Ref_Key));
      if (filter.includes("IsFolder eq false")) {
        counter.wholeCatalog++;
        return ITEMS;
      }
    }
    return undefined;
  };
}

type Json = Record<string, unknown>;
interface Row extends Json {
  ref: string;
  characteristicRef?: string;
  warehouseRef?: string;
  matchedGroupRef?: string;
}

async function setup(env: Record<string, string> = {}, balance = BALANCE) {
  const counter = { balance: 0, turnovers: 0, wholeCatalog: 0 };
  const fake = fake1C({ READ_ONLY: "true", ...env }, { schema: "ut11", data: data(counter, balance) });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (name: string, args: Json) => {
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Json & { items: Row[] };
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  /** Все страницы по nextOffset: склеенный список и первая страница (на ней сводки). */
  const all = async (name: string, args: Json) => {
    const rows: Row[] = [];
    let first: (Json & { items: Row[] }) | undefined;
    let calls = 0;
    for (let offset = 0; calls < 500; calls++) {
      const r = await call(name, { ...args, offset });
      expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
      first ??= r.structuredContent;
      rows.push(...r.structuredContent.items);
      const next = r.structuredContent["nextOffset"];
      if (next === undefined) break;
      offset = next as number;
    }
    return { rows, first: first!, calls: calls + 1 };
  };
  return { call, all, counter, fake, client };
}

const STOCK = "read.analytics.get_warehouse_stock";
const SALES_TOOL = "read.analytics.get_sales_register";
const ITEMS_TOOL = "read.catalog.list_items";
const PERIOD = { from: "2026-10-01", to: "2026-10-31" };
const key = (r: Row) => `${r.ref}|${r.characteristicRef ?? ""}|${r.warehouseRef ?? ""}`;
const inGroups = (groups: string[]) => ITEMS.filter((it) => groups.includes(it.Parent_Key));

describe("листание длинного списка: регистр читается один раз", () => {
  it("остатки 6000 позиций: склеено столько, сколько count, без повторов, одно чтение регистра", async () => {
    const { all, counter, fake } = await setup();
    const { rows, first, calls } = await all(STOCK, { limit: 1000 });
    const expected = BALANCE.length; // строка на позицию и характеристику, склад сворачивается
    expect(first["count"]).toBe(expected);
    expect(rows).toHaveLength(expected);
    expect(new Set(rows.map(key)).size).toBe(expected);
    expect(calls).toBeGreaterThan(10);
    expect(counter.balance).toBe(1);
    // Сведения о позициях — из справочника целиком, а не пачками по 20 на странице: было около 1400 запросов.
    expect(fake.requests.length).toBeLessThan(2 * calls);
    expect(rows.every((r) => typeof r["item"] === "string" && r["item"] !== r.ref)).toBe(true);
  });

  it("первая страница без отбора — сведения по ссылкам: справочник целиком только для продолжения", async () => {
    // Следующих страниц может не быть: справочник на 100 тыс. позиций ради одной страницы стоил бы вдвое дороже (ревью).
    const { call, counter } = await setup();
    const r = await call(STOCK, { limit: 200 });
    expect(counter.wholeCatalog).toBe(0);
    await call(STOCK, { limit: 200, offset: r.structuredContent["nextOffset"] as number });
    expect(counter.wholeCatalog).toBeGreaterThan(0); // 6000 позиций — две страницы по 5000
  });

  it("дедлайн вызова останавливает чтение регистра, а не только ожидание ответа", async () => {
    // pageCached сначала отсоединял чтение от вызова: запрос к 1С доживал до конца и держал очередь (ревью).
    const { call } = await setup({ ODATA_TOOL_DEADLINE_MS: "200", ODATA_RETRIES: "0" });
    const real = globalThis.fetch;
    const balance = { aborted: 0, done: 0 };
    vi.stubGlobal("fetch", (url: string | URL, init?: RequestInit) => {
      if (!decodeURIComponent(String(url)).includes("/Balance")) return real(url, init);
      return new Promise<Response>((resolve, reject) => {
        const t = setTimeout(() => {
          balance.done++;
          resolve(real(url, init));
        }, 1000);
        init?.signal?.addEventListener("abort", () => {
          clearTimeout(t);
          balance.aborted++;
          reject(new DOMException("aborted", "AbortError"));
        });
      });
    });
    const r = await call(STOCK, { limit: 10 });
    expect(r.isError).toBe(true);
    await new Promise((res) => setTimeout(res, 1200));
    expect(balance).toEqual({ aborted: 1, done: 0 });
  });

  it("продолжение, ждавшее чужое чтение, которое прервали, читает само, а не отдаёт чужую ошибку", async () => {
    const { call, counter } = await setup({ ODATA_TOOL_DEADLINE_MS: "300", ODATA_RETRIES: "0" });
    const real = globalThis.fetch;
    let slowFirst = true;
    vi.stubGlobal("fetch", (url: string | URL, init?: RequestInit) => {
      if (!decodeURIComponent(String(url)).includes("/Balance") || !slowFirst) return real(url, init);
      slowFirst = false;
      return new Promise<Response>((resolve, reject) => {
        const t = setTimeout(() => resolve(real(url, init)), 2000);
        init?.signal?.addEventListener("abort", () => {
          clearTimeout(t);
          reject(new DOMException("aborted", "AbortError"));
        });
      });
    });
    // Первая страница медленная и упрётся в дедлайн; продолжение приходит, пока она читает, и ждёт её чтение.
    const first = call(STOCK, { limit: 1000 });
    await new Promise((res) => setTimeout(res, 100));
    const next = await call(STOCK, { limit: 1000, offset: 200 });
    expect((await first).isError).toBe(true);
    expect(next.isError, next.text.slice(0, 200)).toBeFalsy();
    expect(next.structuredContent.items.length).toBeGreaterThan(0);
    expect(counter.balance).toBe(1); // прерванное чтение до подделки не дошло, своё — одно
  });

  it("snapshot: в окне кэша данные те же; окно истекло и в 1С провели документ — dataChanged", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const balance = BALANCE.slice(0, 600);
    const { call } = await setup({}, balance);
    const first = await call(STOCK, { limit: 1000 });
    const snapshot = first.structuredContent["snapshot"] as string;
    expect(snapshot).toMatch(/^[0-9a-f]{16}$/);
    const next = first.structuredContent["nextOffset"] as number;
    const same = await call(STOCK, { limit: 1000, offset: next, snapshot });
    expect(same.structuredContent["dataChanged"]).toBeUndefined();
    balance.push(bal(I(1), 5, 0)); // новый остаток после первой страницы
    vi.setSystemTime(Date.now() + PAGE_READ_TTL_MS + 1);
    const changed = await call(STOCK, { limit: 1000, offset: next, snapshot });
    expect(changed.structuredContent["dataChanged"]).toBe(true);
    expect(changed.structuredContent["note"]).toContain("прочитайте заново с offset=0");
  });

  it("snapshot есть и у продаж", async () => {
    const { call } = await setup();
    const r = await call(SALES_TOOL, { ...PERIOD, limit: 10 });
    expect(r.structuredContent["snapshot"]).toMatch(/^[0-9a-f]{16}$/);
  });

  it("в простое чтение уходит из памяти по таймеру, а не ждёт следующего вызова", async () => {
    vi.useFakeTimers({ toFake: ["Date", "setTimeout"] });
    const owner = {};
    let loads = 0;
    const load = async () => ++loads;
    await pageCached(owner, "k", true, load);
    expect(await pageCached(owner, "k", false, load)).toBe(1); // в окне — то же чтение
    vi.advanceTimersByTime(PAGE_READ_TTL_MS);
    // Время запроса отстаёт от таймера: запись ушла именно по таймеру, а не по сроку при поиске.
    vi.setSystemTime(Date.now() - PAGE_READ_TTL_MS + 1);
    expect(await pageCached(owner, "k", false, load)).toBe(2);
  });

  it("вытесненные и заменённые чтения не держатся таймерами: таймеров не больше записей", async () => {
    // Таймер удаления держал запись в замыкании: 10 разных периодов и 5 повторов первой страницы — 15 выборок в памяти
    // вместо 4 (ревью).
    vi.useFakeTimers({ toFake: ["Date", "setTimeout", "clearTimeout"] });
    const owner = {};
    for (let i = 0; i < 10; i++) await pageCached(owner, `k${i}`, true, async () => i);
    expect(vi.getTimerCount()).toBe(PAGE_READ_ENTRIES);
    for (let i = 0; i < 5; i++) await pageCached(owner, "k9", true, async () => i);
    expect(vi.getTimerCount()).toBe(PAGE_READ_ENTRIES);
    vi.advanceTimersByTime(PAGE_READ_TTL_MS);
    expect(vi.getTimerCount()).toBe(0);
  });

  it("справочник не поместился в предел — следующие страницы не пробуют читать его целиком заново", async () => {
    const { call, counter } = await setup({ ODATA_ANALYTICS_MAX_ROWS: "1000" }, BALANCE.slice(0, 500));
    let offset = 0;
    const attempts: number[] = [];
    for (let page = 0; page < 4; page++) {
      const r = await call(STOCK, { limit: 100, offset });
      expect(r.isError, r.text.slice(0, 200)).toBeFalsy();
      attempts.push(counter.wholeCatalog);
      offset = r.structuredContent["nextOffset"] as number;
    }
    expect(attempts[0]).toBe(0); // первая страница — по ссылкам
    expect(attempts[1]).toBeGreaterThan(0); // первая попытка на продолжении
    expect(attempts[3]).toBe(attempts[1]); // дальше — сразу по ссылкам
  });

  it("справочник больше предела выборки — сведения по ссылкам, как раньше, ответ не ломается", async () => {
    // Остатки — 500 строк, в пределе; справочник на 6000 позиций — за ним.
    const { call, fake } = await setup({ ODATA_ANALYTICS_MAX_ROWS: "1000" }, BALANCE.slice(0, 500));
    const r = await call(STOCK, { limit: 100 });
    expect(r.structuredContent.items).toHaveLength(100);
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.structuredContent.items[0]!["item"]).toMatch(/^Позиция синтетическая/);
    expect(fake.violations).toEqual([]);
  });

  it("продажи 6000 позиций: то же для оборотов", async () => {
    const { all, counter } = await setup();
    const { rows, first } = await all(SALES_TOOL, { ...PERIOD, limit: 1000 });
    const withSales = ITEMS.filter((_, i) => i % 4 !== 0).length;
    expect(first["count"]).toBe(withSales);
    expect(new Set(rows.map(key)).size).toBe(withSales);
    expect(rows).toHaveLength(withSales);
    expect(counter.turnovers).toBe(1);
  });

  it("первая страница всегда читает заново: новый вызов с offset=0 видит свежие данные", async () => {
    const { call, counter } = await setup();
    await call(STOCK, { limit: 50 });
    await call(STOCK, { limit: 50 });
    expect(counter.balance).toBe(2);
  });

  it("окно истекло — следующая страница читает регистр заново", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const { call, counter } = await setup();
    const r = await call(STOCK, { limit: 1000 });
    vi.setSystemTime(Date.now() + PAGE_READ_TTL_MS + 1);
    await call(STOCK, { limit: 1000, offset: r.structuredContent["nextOffset"] as number });
    expect(counter.balance).toBe(2);
  });

  it("другие аргументы — другое чтение; старые записи вытесняются, а не копятся", async () => {
    const { call, counter } = await setup();
    const dates = Array.from({ length: PAGE_READ_ENTRIES + 1 }, (_, i) => `2026-10-${String(i + 10)}`);
    for (const asOf of dates) await call(STOCK, { asOf, limit: 1000 });
    expect(counter.balance).toBe(dates.length);
    // Последнее чтение ещё в кэше: следующая страница без нового чтения.
    await call(STOCK, { asOf: dates.at(-1), limit: 1000, offset: 200 });
    expect(counter.balance).toBe(dates.length);
    // Первое вытеснено: его продолжение читает заново.
    await call(STOCK, { asOf: dates[0], limit: 1000, offset: 200 });
    expect(counter.balance).toBe(dates.length + 1);
  });
});

describe("groups: несколько групп за один вызов", () => {
  const LIST = ["Техника/Смартфоны/Бренд А", "Бренд Б", "Планшеты", "Распродажа/Смартфоны/Бренд А"];

  it("остатки: те же позиции и итоги, что вызовы по одной группе, регистр читается один раз", async () => {
    const { all, counter } = await setup();
    const many = await all(STOCK, { groups: LIST, limit: 1000 });
    expect(counter.balance).toBe(1);
    const single = [];
    for (const group of LIST) single.push(await all(STOCK, { group, limit: 1000 }));
    const singleRows = single.flatMap((s) => s.rows);
    expect(many.rows.map(key).sort()).toEqual(singleRows.map(key).sort());
    const total = (field: string) => single.reduce((t, s) => t + (s.first[field] as number), 0);
    expect(many.first["totalInStock"]).toBe(total("totalInStock"));
    expect(many.first["totalToShip"]).toBe(total("totalToShip"));
    // byGroup — итог каждой группы, как отдельный вызов; matchedGroupRef — группа из списка.
    const byGroup = many.first["byGroup"] as { ref: string; positions: number; inStock: number }[];
    const refs = (many.first["itemGroups"] as { ref: string }[]).map((g) => g.ref);
    expect(refs).toEqual([G(3), G(4), G(5), G(8)]);
    byGroup.forEach((g, i) => {
      expect(g.ref).toBe(refs[i]);
      expect(g.positions).toBe(single[i]!.first["count"]);
      expect(g.inStock).toBe(single[i]!.first["totalInStock"]);
    });
    for (const r of many.rows)
      expect(r.matchedGroupRef).toBe(ITEMS.find((it) => it.Ref_Key === r.ref)!.Parent_Key);
  });

  it("продажи: те же позиции и итоги до копейки, что вызовы по одной группе", async () => {
    const { all, counter } = await setup();
    const many = await all(SALES_TOOL, { ...PERIOD, groups: LIST, limit: 1000 });
    expect(counter.turnovers).toBe(1);
    const single = [];
    for (const group of LIST) single.push(await all(SALES_TOOL, { ...PERIOD, group, limit: 1000 }));
    expect(many.rows.map(key).sort()).toEqual(single.flatMap((s) => s.rows.map(key)).sort());
    const cents = (v: unknown) => Math.round((v as number) * 100);
    const total = (field: string) => single.reduce((t, s) => t + cents(s.first[field]), 0);
    for (const field of ["totalQuantity", "sold", "returns", "revenue", "revenueExVat"])
      expect(cents(many.first[field]), field).toBe(total(field));
    const byGroup = many.first["byGroup"] as Record<string, number>[];
    byGroup.forEach((g, i) => {
      expect(g["positions"]).toBe(single[i]!.first["count"]);
      expect(cents(g["revenue"])).toBe(cents(single[i]!.first["revenue"]));
      expect(g["quantity"]).toBe(single[i]!.first["totalQuantity"]);
    });
  });

  it("вложенные группы: позиция входит один раз и относится к самой глубокой", async () => {
    const { all } = await setup();
    // «Бренд Б», а не «Бренд А»: у позиций «Бренда А» по построению нет остатка, проверка была бы пустой.
    const many = await all(STOCK, { groups: ["Техника", "Техника/Смартфоны/Бренд Б"], limit: 1000 });
    const whole = await all(STOCK, { group: "Техника", limit: 1000 });
    expect(many.rows.map(key).sort()).toEqual(whole.rows.map(key).sort());
    expect(new Set(many.rows.map(key)).size).toBe(many.rows.length);
    for (const r of many.rows) {
      const home = ITEMS.find((it) => it.Ref_Key === r.ref)!.Parent_Key;
      expect(r.matchedGroupRef).toBe(home === G(4) ? G(4) : G(1));
    }
    expect(many.rows.filter((r) => r.matchedGroupRef === G(4)).length).toBeGreaterThan(0);
    expect(many.rows.filter((r) => r.matchedGroupRef === G(1)).length).toBeGreaterThan(0);
    const byGroup = many.first["byGroup"] as { positions: number; inStock: number }[];
    expect(byGroup.reduce((t, g) => t + g.positions, 0)).toBe(whole.first["count"]);
    expect(byGroup.reduce((t, g) => t + g.inStock, 0)).toBe(whole.first["totalInStock"]);
  });

  it("одна и та же группа дважды — как один раз", async () => {
    const { call } = await setup();
    const r = await call(STOCK, { groups: ["Планшеты", G(5)], limit: 5 });
    expect(r.isError, r.text).toBeFalsy();
    expect(r.structuredContent["itemGroups"]).toHaveLength(1);
  });

  it("group и groups вместе — отказ; без groups ответ прежний, без новых полей", async () => {
    const { call } = await setup();
    const both = await call(STOCK, { group: "Планшеты", groups: ["Бренд Б"] });
    expect(both.isError).toBe(true);
    expect(both.text).toContain("group или groups");
    const one = await call(STOCK, { group: "Планшеты", limit: 5 });
    expect(one.structuredContent["itemGroup"]).toMatchObject({ ref: G(5) });
    expect(one.structuredContent).not.toHaveProperty("byGroup");
    expect(one.structuredContent.items.some((r) => "matchedGroupRef" in r)).toBe(false);
  });

  it("неизвестная группа в списке — отказ с её названием и похожими группами", async () => {
    const { call } = await setup();
    const r = await call(STOCK, { groups: ["Планшеты", "Техника/Смартфоны/Бренд В"] });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Бренд В");
    // Похожие — по последней части пути.
    expect(r.text).toContain("Техника/01. Смартфоны/Бренд А");
    expect(r.text).toContain("Распродажа/Смартфоны/Бренд А");
  });
});

describe("отказ «группа не найдена» подсказывает похожие", () => {
  it("номер в названии и путь: подсказка по последней части", async () => {
    const { call } = await setup();
    const r = await call(STOCK, { group: "Техника/Смартфон" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Похожие группы");
    expect(r.text).toContain("Техника/01. Смартфоны");
    expect(r.text).toContain("Распродажа/Смартфоны");
  });

  it("ничего похожего — так и сказано, без случайного списка", async () => {
    const { call } = await setup();
    const r = await call(STOCK, { group: "Холодильники" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Похожих групп нет");
  });
});

describe("read.catalog.list_items", () => {
  const nonDeleted = (list: typeof ITEMS) => list.filter((it) => !it.DeletionMark);

  it("все позиции группы с подгруппами, включая позиции без остатка и продаж", async () => {
    const { all } = await setup();
    const { rows, first } = await all(ITEMS_TOOL, { group: "Техника/Смартфоны", limit: 1000 });
    const want = nonDeleted(inGroups([G(3), G(4)]));
    expect(first["count"]).toBe(want.length);
    expect(rows.map((r) => r.ref).sort()).toEqual(want.map((it) => it.Ref_Key).sort());
    // Позиции без остатка тоже здесь: их видно только в этом списке.
    const noStock = want.filter((it) => !hasStock(Number(it.Ref_Key.slice(-12))));
    expect(noStock.length).toBeGreaterThan(0);
    for (const it of noStock) expect(rows.some((r) => r.ref === it.Ref_Key)).toBe(true);
    expect(rows[0]).toMatchObject({
      isFolder: false,
      deletionMark: false,
      path: "Техника/01. Смартфоны/Бренд А",
    });
  });

  it("поля позиции: артикул, код, папка, характеристики", async () => {
    const { call } = await setup();
    const r = await call(ITEMS_TOOL, { query: "ART-1001", group: "Бренд Б" });
    const row = r.structuredContent.items.find((x) => x.ref === I(1001))!;
    expect(row).toMatchObject({
      name: "Позиция синтетическая 1001",
      article: "ART-1001",
      code: "К-01001",
      groupRef: G(4),
      path: "Техника/01. Смартфоны/Бренд Б",
      characteristicsUsage: "НеИспользовать",
    });
  });

  it("помеченные на удаление — только с includeDeleted", async () => {
    const { all } = await setup();
    const shown = await all(ITEMS_TOOL, { limit: 1000 });
    expect(shown.first["count"]).toBe(nonDeleted(ITEMS).length);
    const withDeleted = await all(ITEMS_TOOL, { includeDeleted: true, limit: 1000 });
    expect(withDeleted.first["count"]).toBe(N);
    expect(withDeleted.rows.filter((r) => r["deletionMark"] === true)).toHaveLength(N / 50);
  });

  it("одноимённые группы в разных ветках — по пути; список групп с matchedGroupRef", async () => {
    const { all, call } = await setup();
    const ambiguous = await call(ITEMS_TOOL, { group: "Бренд А" });
    expect(ambiguous.isError).toBe(true);
    expect(ambiguous.text).toContain("Распродажа/Смартфоны/Бренд А");
    const { rows } = await all(ITEMS_TOOL, {
      groups: ["Техника/Смартфоны/Бренд А", "Распродажа/Смартфоны/Бренд А"],
      limit: 1000,
    });
    expect(rows).toHaveLength(nonDeleted(inGroups([G(3), G(8)])).length);
    for (const r of rows)
      expect(r.matchedGroupRef).toBe(ITEMS.find((it) => it.Ref_Key === r.ref)!.Parent_Key);
  });

  it("папки: подгруппы группы с путями; пустая группа — пустой список с пометкой", async () => {
    const { call } = await setup();
    const folders = await call(ITEMS_TOOL, { group: "Техника", kind: "folders" });
    expect(folders.structuredContent.items.map((r) => r["path"])).toEqual([
      "Техника/01. Смартфоны",
      "Техника/01. Смартфоны/Бренд А",
      "Техника/01. Смартфоны/Бренд Б",
      "Техника/Планшеты",
    ]);
    const deleted = await call(ITEMS_TOOL, { group: "Техника", kind: "folders", includeDeleted: true });
    expect(deleted.structuredContent.items.map((r) => r["path"])).toContain("Техника/Старое");
    const empty = await call(ITEMS_TOOL, { group: "Пустая" });
    expect(empty.isError, empty.text).toBeFalsy();
    expect(empty.structuredContent["count"]).toBe(0);
    expect(empty.structuredContent["note"]).toContain("нет позиций");
  });

  it("справочник не опубликован — понятный отказ", async () => {
    const { call, fake } = await setup();
    const available = await fake.conn.available();
    vi.spyOn(fake.conn, "available").mockResolvedValue(
      new Set([...available].filter((s) => !s.startsWith("Catalog_Номенклатура"))),
    );
    const r = await call(ITEMS_TOOL, {});
    expect(r.isError).toBe(true);
    expect(r.text).toContain("Номенклатура");
  });

  it("ref из списка совпадает с ref в остатках: остаток по группе — подмножество списка", async () => {
    const { all } = await setup();
    // В «Бренде А» у позиций нет остатка (каждая 5-я по построению), в «Бренде Б» — есть.
    const list = await all(ITEMS_TOOL, { group: "Техника/Смартфоны", includeDeleted: true, limit: 1000 });
    const stock = await all(STOCK, { group: "Техника/Смартфоны", limit: 1000 });
    const refs = new Set(list.rows.map((r) => r.ref));
    for (const r of stock.rows) expect(refs.has(r.ref)).toBe(true);
    expect(list.rows.length).toBeGreaterThan(new Set(stock.rows.map((r) => r.ref)).size);
  });
});

// Замер для отчёта (вызовы, чтения регистра, запросы к 1С): только при METRICS=1.
it.runIf(process.env["METRICS"] === "1")("замер: полное листание и groups", async () => {
  const out: Record<string, unknown> = {};
  for (const [label, name, args] of [
    ["stock all", STOCK, { limit: 1000 }],
    ["sales all", SALES_TOOL, { ...PERIOD, limit: 1000 }],
    [
      "stock 4 groups",
      STOCK,
      {
        groups: ["Техника/Смартфоны/Бренд А", "Бренд Б", "Планшеты", "Распродажа/Смартфоны/Бренд А"],
        limit: 1000,
      },
    ],
    ["list_items all", ITEMS_TOOL, { limit: 1000 }],
  ] as const) {
    const { all, counter, fake } = await setup();
    const t0 = Date.now();
    const r = await all(name, args as Json);
    out[label] = {
      calls: r.calls,
      rows: r.rows.length,
      count: r.first["count"],
      registerReads: counter.balance + counter.turnovers,
      requests1C: fake.requests.length,
      ms: Date.now() - t0,
    };
  }
  console.log("METRICS " + JSON.stringify(out));
});
