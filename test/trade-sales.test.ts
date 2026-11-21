import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Продажи и закупки Управления торговлей. Состав — как у документов, которые приняла и провела копия рабочей УТ 11.5
 * (21.11.2026): налогообложение НДС, график оплаты и склад в строке коннектор ставит сам (1С через OData их не
 * заполняет), серии — статус 2 и таблица «Серии» со складом. Проведение сверяется по складу, расчётам и сериям.
 */

const g = (p: string, n: number) => `${p.repeat(8)}-0000-0000-0000-${String(n).padStart(12, "0")}`;
const ORG = g("b", 1);
const WH = g("a", 1);
const CP = g("e", 1);
const PARTNER = g("e", 2);
const CONTRACT = g("e", 3);
const AGR = g("e", 4);
const AGR2 = g("e", 5);
const ITEM = g("c", 1);
const SVC = g("c", 2);
const SERIES = g("c", 3);
const NOVAT = g("d", 1);
const VAT20 = g("d", 2);
const RUB = g("a", 9);
const BASE_SALE = g("9", 1);
const BASE_PUR = g("9", 2);

const TABLES: Record<string, Record<string, unknown>[]> = {
  Catalog_Организации: [{ Ref_Key: ORG, Description: "ТЕСТ", Code: "1" }],
  Catalog_Склады: [{ Ref_Key: WH, Description: "Основной", Code: "1" }],
  Catalog_Контрагенты: [{ Ref_Key: CP, Description: "ZZ-TEST", Code: "1", Партнер_Key: PARTNER }],
  Catalog_СтавкиНДС: [
    { Ref_Key: NOVAT, Description: "Без НДС", Code: "1" },
    { Ref_Key: VAT20, Description: "20%", Code: "2" },
  ],
  Catalog_Валюты: [{ Ref_Key: RUB, Description: "RUB", Code: "643" }],
  Catalog_Номенклатура: [
    { Ref_Key: ITEM, Description: "ZZ-TEST товар", ТипНоменклатуры: "Товар" },
    { Ref_Key: SVC, Description: "ZZ-TEST услуга", ТипНоменклатуры: "Услуга" },
  ],
};

function pick(rows: Record<string, unknown>[], filter: string): Record<string, unknown>[] {
  const refs = [...filter.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());
  if (refs.length) return rows.filter((r) => refs.includes(String(r["Ref_Key"]).toLowerCase()));
  const s = /(Code|Description) eq '((?:[^']|'')*)'/.exec(filter);
  if (s) return rows.filter((r) => String(r[s[1]!]) === s[2]!.replace(/''/g, "'"));
  return rows;
}

const SETS = [
  "Document_РеализацияТоваровУслуг",
  "Document_ПриобретениеТоваровУслуг",
  "Document_ВозвратТоваровОтКлиента",
  "Document_ВозвратТоваровПоставщику",
  "Document_СчетНаОплатуКлиенту",
  "Document_ЗаказКлиента",
  "Document_ЗаказПоставщику",
  "Document_СчетФактураВыданный",
  "Document_СчетФактураПолученный",
];
type Dir = "Receipt" | "Expense";
const DIRS: Record<string, { stock: Dir; party?: [string, Dir] }> = {
  Document_РеализацияТоваровУслуг: { stock: "Expense", party: ["РасчетыСКлиентами", "Receipt"] },
  Document_ПриобретениеТоваровУслуг: { stock: "Receipt", party: ["РасчетыСПоставщиками", "Expense"] },
  Document_ВозвратТоваровПоставщику: { stock: "Expense", party: ["РасчетыСПоставщиками", "Receipt"] },
  Document_ВозвратТоваровОтКлиента: { stock: "Receipt", party: ["РасчетыСКлиентами", "Expense"] },
};

/** Как проводит 1С: склад по строкам со складом, расчёты на сумму, серии по таблице «Серии». */
function movesOf(set: string, doc: Record<string, unknown>): Record<string, Record<string, unknown>[]> {
  const d = DIRS[set];
  if (!d) return {};
  const rows = (doc["Товары"] as Record<string, unknown>[]) ?? [];
  const out: Record<string, Record<string, unknown>[]> = {
    ТоварыНаСкладах: rows
      .filter((r) => r["Склад_Key"])
      .map((r) => ({
        RecordType: d.stock,
        Склад_Key: r["Склад_Key"],
        Номенклатура_Key: r["Номенклатура_Key"],
        ВНаличии: r["Количество"],
      })),
    ДвиженияСерийТоваров: ((doc["Серии"] as Record<string, unknown>[]) ?? []).map((s) => ({
      Серия_Key: s["Серия_Key"],
      Количество: s["Количество"],
    })),
  };
  const op = doc["ХозяйственнаяОперация"];
  if (d.party && op !== "ПриемНаКомиссию" && op !== "ВозвратТоваровКомитенту")
    out[d.party[0]] = [
      { RecordType: d.party[1], Сумма: 0, КОплате: doc["СуммаДокумента"] },
      { RecordType: d.party[1], Сумма: doc["СуммаДокумента"] },
    ];
  return out;
}

function setup(
  o: {
    schema?: FakeSchema;
    agreements?: Record<string, unknown>[];
    lastVat?: string | null;
    /** Налогообложение последнего проведённого документа — по набору (нет набора — нет документов). */
    lastVatBySet?: Record<string, string>;
    spoil?: (reg: string, rows: Record<string, unknown>[]) => Record<string, unknown>[];
    basePosted?: boolean;
    /** Флаги ордерной схемы основного склада. */
    orderScheme?: Record<string, unknown>;
  } = {},
) {
  const posts: { set: string; body: Record<string, unknown> }[] = [];
  const patches: Record<string, unknown>[] = [];
  const docs = new Map<
    string,
    {
      set: string;
      body: Record<string, unknown>;
      posted: boolean;
      moves: Record<string, Record<string, unknown>[]>;
      v: number;
    }
  >();
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-ts-")),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) => {
        if (q.fn) return undefined;
        if (q.base === "Catalog_СоглашенияСКлиентами" || q.base === "Catalog_СоглашенияСПоставщиками")
          return (
            o.agreements ?? [
              { Ref_Key: AGR, Description: "Основное", Организация_Key: ORG, Статус: "Действует" },
            ]
          );
        const f = q.params.get("$filter") ?? "";
        if (o.lastVatBySet && SETS.includes(q.base) && f.includes("Posted eq true"))
          return o.lastVatBySet[q.base]
            ? [
                {
                  Number: "0ТУТ-000200",
                  Date: "2026-11-21T09:00:00",
                  НалогообложениеНДС: o.lastVatBySet[q.base],
                },
              ]
            : [];
        if (SETS.includes(q.base) && f.includes("Posted eq true"))
          return o.lastVat === null
            ? []
            : [
                {
                  Number: "0ТУТ-000156",
                  Date: "2026-11-20T09:00:00",
                  НалогообложениеНДС: o.lastVat ?? "ПродажаНеОблагаетсяНДС",
                },
              ];
        if (q.base === "Catalog_Склады" && o.orderScheme)
          return pick(
            TABLES["Catalog_Склады"]!.map((w) => ({ ...w, ...o.orderScheme })),
            f,
          );
        const rows = TABLES[q.base];
        return rows ? pick(rows, f) : undefined;
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
    if (SETS.includes(set)) {
      if (method === "POST" && !ref) {
        const res = await inner(input, init);
        if (!res.ok) return res;
        const body = JSON.parse(String(init?.body)) as Record<string, unknown>;
        posts.push({ set, body });
        docs.set(String(body["Ref_Key"]).toLowerCase(), { set, body, posted: false, moves: {}, v: 1 });
        return json(201, { ...body, Number: "0000-000001", DataVersion: "v1" });
      }
      // Основания: проведённые реализация и приобретение.
      if (method === "GET" && (ref === BASE_SALE || ref === BASE_PUR))
        return json(200, {
          Ref_Key: ref,
          Posted: o.basePosted ?? true,
          Организация_Key: ORG,
          Контрагент_Key: CP,
          Партнер_Key: PARTNER,
          Договор_Key: CONTRACT,
          Валюта_Key: RUB,
        });
      const d = ref ? docs.get(ref) : undefined;
      if (d && /\/(Post|Unpost)\?/.test(path)) {
        d.posted = path.includes("/Post?");
        d.moves = d.posted
          ? Object.fromEntries(
              Object.entries(movesOf(d.set, d.body)).map(([r, rows]) => [
                r,
                o.spoil ? o.spoil(r, rows) : rows,
              ]),
            )
          : {};
        d.v++;
        return json(200, {});
      }
      if (d && method === "GET")
        return json(200, { ...d.body, Ref_Key: ref, Posted: d.posted, DataVersion: `v${d.v}` });
      if (d && method === "PATCH") {
        const patch = JSON.parse(String(init?.body)) as Record<string, unknown>;
        patches.push(patch);
        Object.assign(d.body, patch);
        return json(200, { ...d.body, Ref_Key: ref });
      }
    }
    const reg = /^AccumulationRegister_([^(]+)\(Recorder='([0-9a-f-]{36})'/i.exec(path);
    if (reg) {
      const rows = (docs.get(reg[2]!.toLowerCase())?.moves[reg[1]!] ?? []).map((r) => ({
        ...r,
        Active: true,
      }));
      if (!rows.length)
        return json(404, {
          "odata.error": { code: "9", message: { lang: "ru", value: "Экземпляр сущности не найден" } },
        });
      return json(200, { RecordSet: rows });
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
  const create = async (name: string, args: Record<string, unknown>) => {
    const p = await call(name, args);
    expect(p.isError, p.text.slice(0, 400)).toBeFalsy();
    const d = await call(name, { ...args, confirm: true, operationId: p.sc["operationId"] });
    expect(d.isError, d.text.slice(0, 400)).toBeFalsy();
    return { preview: p, ref: String(posts.at(-1)!.body["Ref_Key"]) };
  };
  const post = async (entitySet: string, ref: string, p = true) => {
    await call("write.document.post_document", { entitySet, ref, post: p });
    return call("write.document.post_document", { entitySet, ref, post: p, confirm: true });
  };
  return { call, create, post, posts, patches, docs, fake };
}

const SHIP = "write.sales.create_shipment";
const PUR = "write.purchase.create_purchase";
const line = (q = 2, price = 50, extra: Record<string, unknown> = {}) => ({
  nomenclatureRef: ITEM,
  quantity: q,
  price,
  ...extra,
});
const base = { organization: "ТЕСТ", counterpartyRef: CP, warehouse: "Основной" };
const payloadOf = (r: { sc: Record<string, unknown> }) => r.sc["payload"] as Record<string, unknown>;

describe("продажи и закупки УТ: состав", () => {
  it("реализация: партнёр, единственное соглашение, налогообложение из последней реализации, график оплаты, склад в строке", async () => {
    const s = setup();
    const r = await s.call(SHIP, { ...base, contractRef: CONTRACT, lines: [line()], date: "2026-11-21" });
    expect(r.isError, r.text).toBeFalsy();
    expect(payloadOf(r)).toMatchObject({
      ХозяйственнаяОперация: "РеализацияКлиенту",
      Партнер_Key: PARTNER,
      Контрагент_Key: CP,
      Договор_Key: CONTRACT,
      Соглашение_Key: AGR,
      НалогообложениеНДС: "ПродажаНеОблагаетсяНДС",
      ЦенаВключаетНДС: true,
      Статус: "Отгружено",
      Склад_Key: WH,
      СуммаДокумента: 100,
      Товары: [
        {
          Номенклатура_Key: ITEM,
          Количество: 2,
          КоличествоУпаковок: 2,
          Цена: 50,
          Сумма: 100,
          СуммаСНДС: 100,
          СтавкаНДС_Key: NOVAT,
          Склад_Key: WH,
        },
      ],
      ЭтапыГрафикаОплаты: [
        {
          ВариантОплаты: "КредитПослеОтгрузки",
          ДатаПлатежа: "2026-11-21T00:00:00",
          ПроцентПлатежа: 100,
          СуммаПлатежа: 100,
        },
      ],
    });
    const notes = JSON.stringify(r.sc["notes"]);
    expect(notes).toContain("единственное у партнёра");
    expect(notes).toContain("№ 0ТУТ-000156");
    expect(s.fake.violations).toEqual([]);
  });

  it("НДС: цена с НДС — сумма с НДС и НДС из неё; без — НДС сверху; серия — статус 2 и таблица со складом", async () => {
    const s = setup({ lastVat: "ПродажаОблагаетсяНДС" });
    const incl = await s.call(SHIP, {
      ...base,
      lines: [line(1, 120, { vatRate: "НДС20", seriesRef: SERIES })],
    });
    expect(payloadOf(incl)["Товары"]).toMatchObject([
      { Сумма: 120, СуммаНДС: 20, СуммаСНДС: 120, СтавкаНДС_Key: VAT20, СтатусУказанияСерий: 2 },
    ]);
    expect(payloadOf(incl)["Серии"]).toEqual([
      { LineNumber: 1, Номенклатура_Key: ITEM, Серия_Key: SERIES, Количество: 1, Склад_Key: WH },
    ]);
    const excl = await s.call(SHIP, {
      ...base,
      sumIncludesVat: false,
      lines: [line(1, 100, { vatRate: "НДС20" })],
    });
    expect(payloadOf(excl)).toMatchObject({
      ЦенаВключаетНДС: false,
      СуммаДокумента: 120,
      Товары: [{ Сумма: 100, СуммаНДС: 20, СуммаСНДС: 120 }],
    });
    const warn = await s.call(SHIP, { ...base, lines: [line()] });
    expect(JSON.stringify(warn.sc["notes"])).toContain("ставка «Без НДС»");
    // Не облагается — предупреждения нет («НеОблагается» не путается с «Облагается»).
    const ok = await setup().call(SHIP, { ...base, lines: [line()] });
    expect(JSON.stringify(ok.sc["notes"])).not.toContain("ставка «Без НДС»");
  });

  it("акт в УТ — реализация с услугой: без склада, с содержанием, с заметкой", async () => {
    const s = setup();
    const r = await s.call("write.sales.create_act", {
      organization: "ТЕСТ",
      counterpartyRef: CP,
      lines: [{ nomenclatureRef: SVC, quantity: 1, price: 50, content: "Услуга ZZ" }],
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(r.sc["willCreate"]).toBe("Document_РеализацияТоваровУслуг");
    const p = payloadOf(r);
    expect(p).not.toHaveProperty("Склад_Key");
    expect(p["Товары"]).toEqual([
      expect.objectContaining({ Номенклатура_Key: SVC, Содержание: "Услуга ZZ", Сумма: 50 }),
    ]);
    expect(JSON.stringify(p["Товары"])).not.toContain("Склад_Key");
    expect(JSON.stringify(r.sc["notes"])).toContain("продаются реализацией");
  });

  it("приобретение: закупка по умолчанию, приём на комиссию — operationKind; налогообложение 1С заполняет сама", async () => {
    const s = setup();
    const r = await s.call(PUR, { ...base, lines: [line()] });
    expect(payloadOf(r)).toMatchObject({ ХозяйственнаяОперация: "ЗакупкаУПоставщика", Склад_Key: WH });
    expect(payloadOf(r)).not.toHaveProperty("НалогообложениеНДС");
    expect(payloadOf(r)).not.toHaveProperty("Соглашение_Key");
    const c = await s.call(PUR, {
      ...base,
      operationKind: "ПриемНаКомиссию",
      agreementRef: AGR,
      lines: [line()],
    });
    expect(payloadOf(c)).toMatchObject({ ХозяйственнаяОперация: "ПриемНаКомиссию", Соглашение_Key: AGR });
  });

  it("возвраты: поставщику — по приобретению в шапке и строках; от клиента — по реализации в строках", async () => {
    const s = setup();
    const a = await s.call("write.warehouse.create_return_to_supplier", {
      ...base,
      basisRef: BASE_PUR,
      lines: [line(1)],
    });
    expect(payloadOf(a)).toMatchObject({
      ДокументПоступления_Key: BASE_PUR,
      Товары: [{ ДокументПоступления_Key: BASE_PUR, Склад_Key: WH }],
    });
    const b = await s.call("write.warehouse.create_return_from_customer", {
      ...base,
      basisRef: BASE_SALE,
      lines: [line(1)],
    });
    expect(payloadOf(b)).toMatchObject({
      ХозяйственнаяОперация: "ВозвратТоваровОтКлиента",
      Соглашение_Key: AGR,
      Товары: [
        {
          ДокументРеализации: BASE_SALE,
          ДокументРеализации_Type: "StandardODATA.Document_РеализацияТоваровУслуг",
        },
      ],
    });
    const k = await s.call("write.warehouse.create_return_to_supplier", {
      ...base,
      operationKind: "ВозвратТоваровКомитенту",
      basisRef: BASE_PUR,
      lines: [line(1)],
    });
    expect(payloadOf(k)).toMatchObject({ ХозяйственнаяОперация: "ВозвратТоваровКомитенту" });
    expect(s.fake.violations).toEqual([]);
  });

  it("счёт клиенту — сумма и график оплаты без строк; заказы — статусы по умолчанию и явные", async () => {
    const s = setup();
    const inv = await s.call("write.sales.create_invoice", {
      organization: "ТЕСТ",
      counterpartyRef: CP,
      lines: [line(2, 50, { vatRate: "БезНДС" })],
      paymentDueDate: "2026-11-30",
      basisRef: BASE_SALE,
    });
    expect(payloadOf(inv)).toMatchObject({
      СуммаДокумента: 100,
      ДокументОснование: BASE_SALE,
      ЭтапыГрафикаОплаты: [{ ДатаПлатежа: "2026-11-30T00:00:00", ПроцентПлатежа: 100, СуммаПлатежа: 100 }],
    });
    expect(payloadOf(inv)).not.toHaveProperty("Товары");
    const co = await s.call("write.sales.create_customer_order", { ...base, lines: [line()] });
    expect(payloadOf(co)).toMatchObject({
      Статус: "КОбеспечению",
      ХозяйственнаяОперация: "РеализацияКлиенту",
      Соглашение_Key: AGR,
    });
    const so = await s.call("write.purchase.create_supplier_order", {
      ...base,
      lines: [line()],
      status: "Согласован",
    });
    expect(payloadOf(so)).toMatchObject({
      Статус: "Согласован",
      ХозяйственнаяОперация: "ЗакупкаУПоставщика",
    });
    expect(s.fake.violations).toEqual([]);
  });

  it("счета-фактуры — на основании проведённого документа: контрагент и договор из основания; полученный — номер и дата обязательны", async () => {
    const s = setup();
    const iss = await s.call("write.sales.create_issued_invoice", { baseDocumentRef: BASE_SALE });
    expect(payloadOf(iss)).toMatchObject({
      Организация_Key: ORG,
      Контрагент: CP,
      Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
      ДокументОснование: BASE_SALE,
      КодВидаОперации: "01",
      ДокументыОснования: [{ ДокументОснование: BASE_SALE }],
      Договор: CONTRACT,
    });
    const noNum = await s.call("write.purchase.create_received_invoice", { baseDocumentRef: BASE_PUR });
    expect(noNum.text).toContain("incomingNumber");
    const rec = await s.call("write.purchase.create_received_invoice", {
      baseDocumentRef: BASE_PUR,
      incomingNumber: "77",
      incomingDate: "2026-11-20",
    });
    expect(payloadOf(rec)).toMatchObject({
      Number: "77",
      ДатаСоставления: "2026-11-20T00:00:00",
      ДокументыОснования: [
        {
          ДокументОснование: BASE_PUR,
          ДокументОснование_Type: "StandardODATA.Document_ПриобретениеТоваровУслуг",
        },
      ],
    });
    const unposted = setup({ basePosted: false });
    const r = await unposted.call("write.sales.create_issued_invoice", { baseDocumentRef: BASE_SALE });
    expect(r.text).toContain("не проведён");
  });
});

describe("продажи и закупки УТ: налогообложение по умолчанию", () => {
  it("заказов клиента у организации ещё нет — налогообложение из последней реализации, с заметкой", async () => {
    const s = setup({ lastVatBySet: { Document_РеализацияТоваровУслуг: "ПродажаОблагаетсяНДС" } });
    const r = await s.call("write.sales.create_customer_order", {
      ...base,
      lines: [line(1, 120, { vatRate: "НДС20" })],
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(payloadOf(r)["НалогообложениеНДС"]).toBe("ПродажаОблагаетсяНДС");
    expect(JSON.stringify(r.sc["notes"])).toContain("№ 0ТУТ-000200");
  });
});

describe("продажи и закупки УТ: отказы без записи", () => {
  it.each([
    [
      SHIP,
      { ...base, lines: [line()] },
      {
        agreements: [
          { Ref_Key: AGR, Description: "А", Организация_Key: ORG },
          { Ref_Key: AGR2, Description: "Б", Организация_Key: ORG },
        ],
      },
      "несколько",
    ],
    [SHIP, { ...base, lines: [line()] }, { agreements: [] }, "нет действующих"],
    [SHIP, { ...base, lines: [line()] }, { lastVat: null }, "vatTaxation"],
    [
      "write.warehouse.create_return_from_customer",
      { ...base, operationKind: "ВозвратОтРозничногоПокупателя", basisRef: BASE_SALE, lines: [line()] },
      {},
      "розница",
    ],
    [
      "write.sales.create_customer_order",
      { ...base, operationKind: "ПоставкаПодПринципала", lines: [line()] },
      {},
      "принципала",
    ],
    ["write.warehouse.create_return_to_supplier", { ...base, lines: [line()] }, {}, "basisRef"],
    [
      "write.sales.create_services_act",
      { counterpartyRef: CP, nomenclatureGroupRef: g("9", 9), lines: [line()] },
      {},
      "create_act",
    ],
    [
      "write.purchase.create_supplier_invoice",
      { counterpartyRef: CP, lines: [line()] },
      {},
      "create_payment_request",
    ],
    [SHIP, { ...base, lines: [line(1, 10, { vatRate: "НДС18" })] }, {}, "vatRate"],
  ] as const)("%s %o → «%s»", async (tool, args, o, text) => {
    const s = setup(o as Parameters<typeof setup>[0]);
    const r = await s.call(tool, args as Record<string, unknown>);
    expect(r.isError, r.text).toBe(true);
    expect(r.text).toContain(text);
    expect(s.posts).toEqual([]);
  });
});

describe("продажи и закупки УТ: проведение со сверкой", () => {
  it.each([
    [
      SHIP,
      "Document_РеализацияТоваровУслуг",
      { ...base, lines: [line(2, 50, { seriesRef: SERIES }), line(1, 10)] },
    ],
    [PUR, "Document_ПриобретениеТоваровУслуг", { ...base, lines: [line()] }],
    [
      PUR,
      "Document_ПриобретениеТоваровУслуг",
      { ...base, operationKind: "ПриемНаКомиссию", lines: [line()] },
    ],
    [
      "write.warehouse.create_return_to_supplier",
      "Document_ВозвратТоваровПоставщику",
      { ...base, basisRef: BASE_PUR, lines: [line(1)] },
    ],
    [
      "write.warehouse.create_return_from_customer",
      "Document_ВозвратТоваровОтКлиента",
      { ...base, basisRef: BASE_SALE, lines: [line(1)] },
    ],
  ] as const)("%s / %s: склад, расчёты, серии совпали; отмена — пусто", async (tool, set, args) => {
    const s = setup();
    const { ref } = await s.create(tool, args);
    const p = await s.post(set, ref);
    expect(p.isError, p.text).toBeFalsy();
    expect(p.sc).toMatchObject({ done: true, verified: true, movementsVerified: true });
    const u = await s.post(set, ref, false);
    expect(u.sc).toMatchObject({ verified: true, movementsVerified: true });
  });

  it("другое количество на складе, другая сумма расчётов, нет движения серии — verified=false", async () => {
    const cases: Array<[Parameters<typeof setup>[0], string]> = [
      [
        { spoil: (r, rows) => (r === "ТоварыНаСкладах" ? rows.map((x) => ({ ...x, ВНаличии: 1 })) : rows) },
        "склад",
      ],
      [
        {
          spoil: (r, rows) =>
            r === "РасчетыСКлиентами" ? rows.map((x) => ({ ...x, Сумма: Number(x["Сумма"]) ? 1 : 0 })) : rows,
        },
        "РасчетыСКлиентами",
      ],
      [{ spoil: (r, rows) => (r === "ДвиженияСерийТоваров" ? [] : rows) }, "серия"],
    ];
    for (const [o, text] of cases) {
      const s = setup(o);
      const { ref } = await s.create(SHIP, { ...base, lines: [line(2, 50, { seriesRef: SERIES })] });
      const p = await s.post("Document_РеализацияТоваровУслуг", ref);
      expect(p.sc).toMatchObject({ verified: false, movementsVerified: false });
      expect(JSON.stringify(p.sc["movements"])).toContain(text);
    }
  });

  it("ордерный склад: склад и серии не сверяются (их меняют ордера), расчёты сверяются; схема с более поздней даты — сверка", async () => {
    // Как при ордерной схеме: документ «в наличии» не меняет и серий не двигает.
    const spoil = (r: string, rows: Record<string, unknown>[]) =>
      r === "ТоварыНаСкладах" || r === "ДвиженияСерийТоваров" ? [] : rows;
    const flagOut = {
      ИспользоватьОрдернуюСхемуПриОтгрузке: true,
      ДатаНачалаОрдернойСхемыПриОтгрузке: "2026-01-01T00:00:00",
    };
    const s = setup({ spoil, orderScheme: flagOut });
    const { ref } = await s.create(SHIP, {
      ...base,
      date: "2026-11-21",
      lines: [line(2, 50, { seriesRef: SERIES })],
    });
    const p = await s.post("Document_РеализацияТоваровУслуг", ref);
    expect(p.sc).toMatchObject({ done: true, verified: true, movementsVerified: true });
    expect(String(p.sc["movementsNote"])).toContain("ордерной схемой («Основной»)");
    // Расчёты по-прежнему сверяются: порча суммы видна и на ордерном складе.
    const bad = setup({
      orderScheme: flagOut,
      spoil: (r, rows) =>
        r === "РасчетыСКлиентами" ? rows.map((x) => ({ ...x, Сумма: 1 })) : spoil(r, rows),
    });
    const b = await bad.create(SHIP, { ...base, date: "2026-11-21", lines: [line()] });
    expect((await bad.post("Document_РеализацияТоваровУслуг", b.ref)).sc).toMatchObject({ verified: false });
    // Приобретение смотрит флаг поступления, а не отгрузки: отгрузочная схема его не выключает.
    const pur = setup({ spoil, orderScheme: flagOut });
    const c = await pur.create(PUR, { ...base, date: "2026-11-21", contractRef: CONTRACT, lines: [line()] });
    expect((await pur.post("Document_ПриобретениеТоваровУслуг", c.ref)).sc).toMatchObject({
      verified: false,
    });
    // Приём на комиссию на ордерном складе: расчётов с поставщиком нет, склад не сверяется — признака сверки нет.
    const flagIn = { ИспользоватьОрдернуюСхемуПриПоступлении: true };
    const com = setup({ spoil, orderScheme: flagIn });
    const e = await com.create(PUR, {
      ...base,
      contractRef: CONTRACT,
      operationKind: "ПриемНаКомиссию",
      lines: [line()],
    });
    const ep = await com.post("Document_ПриобретениеТоваровУслуг", e.ref);
    expect(ep.sc).toMatchObject({ done: true, verified: true });
    expect(ep.sc).not.toHaveProperty("movementsVerified");
    expect(String(ep.sc["movementsNote"])).toContain("Сверять было нечего");
    // Схема с 01.12 — на дату документа её нет: сверка идёт и видит пустой склад.
    const later = setup({
      spoil,
      orderScheme: { ...flagOut, ДатаНачалаОрдернойСхемыПриОтгрузке: "2026-12-01T00:00:00" },
    });
    const d = await later.create(SHIP, { ...base, date: "2026-11-21", lines: [line()] });
    expect((await later.post("Document_РеализацияТоваровУслуг", d.ref)).sc).toMatchObject({
      verified: false,
    });
  });

  it("ревью выпуска 5: услуга со складом в строке, расчёты в у.е., статус не «Отгружено», нечего сверять", async () => {
    // Услуга в реализации товаров: склад в строке есть, движения по складу 1С не делает — расхождения нет.
    const svc = setup({
      spoil: (r, rows) =>
        r === "ТоварыНаСкладах" ? rows.filter((x) => x["Номенклатура_Key"] !== SVC) : rows,
    });
    const a = await svc.create(SHIP, {
      ...base,
      lines: [line(), { nomenclatureRef: SVC, quantity: 1, price: 30 }],
    });
    expect((await svc.post("Document_РеализацияТоваровУслуг", a.ref)).sc).toMatchObject({
      verified: true,
      movementsVerified: true,
    });
    // Расчёты в валюте договора: сумма в регистре другая — не сверяется, заметка; склад сверен.
    const USD = g("a", 8);
    const fx = setup({
      spoil: (r, rows) =>
        r === "РасчетыСКлиентами"
          ? rows.map((x) => ({ ...x, Сумма: Number(x["Сумма"]) ? 1.25 : 0, Валюта_Key: USD }))
          : rows,
    });
    const b = await fx.create(SHIP, { ...base, lines: [line()] });
    fx.docs.get(b.ref)!.body["Валюта_Key"] = RUB;
    const pb = await fx.post("Document_РеализацияТоваровУслуг", b.ref);
    expect(pb.sc).toMatchObject({ verified: true, movementsVerified: true });
    expect(String(pb.sc["movementsNote"])).toContain("в валюте договора");
    // В той же валюте порча суммы видна, как раньше.
    const same = setup({
      spoil: (r, rows) =>
        r === "РасчетыСКлиентами" ? rows.map((x) => ({ ...x, Сумма: 1, Валюта_Key: RUB })) : rows,
    });
    const c = await same.create(SHIP, { ...base, lines: [line()] });
    same.docs.get(c.ref)!.body["Валюта_Key"] = RUB;
    expect((await same.post("Document_РеализацияТоваровУслуг", c.ref)).sc).toMatchObject({ verified: false });
    // Статус реализации сменили: движения не известны — без сверки и без признака «сверено».
    const st = setup({ spoil: () => [] });
    const d = await st.create(SHIP, { ...base, lines: [line()] });
    st.docs.get(d.ref)!.body["Статус"] = "ВПути";
    const pd = await st.post("Document_РеализацияТоваровУслуг", d.ref);
    expect(pd.sc).toMatchObject({ done: true, verified: true });
    expect(pd.sc).not.toHaveProperty("movementsVerified");
    expect(String(pd.sc["movementsNote"])).toContain("«ВПути»");
    // Приём на комиссию без складов в строках и без движений: сверять нечего — признака нет.
    const nothing = setup({ spoil: () => [] });
    const e = await nothing.create(PUR, {
      ...base,
      contractRef: CONTRACT,
      operationKind: "ПриемНаКомиссию",
      lines: [line()],
    });
    for (const r of nothing.docs.get(e.ref)!.body["Товары"] as Record<string, unknown>[])
      delete r["Склад_Key"];
    const pe = await nothing.post("Document_ПриобретениеТоваровУслуг", e.ref);
    expect(pe.sc).not.toHaveProperty("movementsVerified");
    expect(String(pe.sc["movementsNote"])).toContain("Сверять было нечего");
  });

  it("заказ, счёт и счёт-фактура — только состояние документа, с пояснением", async () => {
    const s = setup();
    const { ref } = await s.create("write.sales.create_customer_order", { ...base, lines: [line()] });
    const p = await s.post("Document_ЗаказКлиента", ref);
    expect(p.sc).toMatchObject({ done: true, verified: true });
    expect(String(p.sc["movementsNote"])).toContain("состояние документа");
  });
});

describe("продажи и закупки УТ: строки, копия, поля", () => {
  const SET = "Document_РеализацияТоваровУслуг";
  const twice = async (s: ReturnType<typeof setup>, name: string, args: Record<string, unknown>) => {
    const p = await s.call(name, args);
    if (p.isError) return p;
    return s.call(name, {
      ...args,
      confirm: true,
      ...(p.sc["operationId"] ? { operationId: p.sc["operationId"] } : {}),
    });
  };
  const made = async (s: ReturnType<typeof setup>, lines = [line(2, 50)]) => {
    const { ref } = await s.create(SHIP, { ...base, lines });
    for (const r of s.docs.get(ref)!.body["Товары"] as Record<string, unknown>[])
      r["ОбъектРасчетов_Key"] = g("8", 1);
    return ref;
  };

  it("ревью выпуска 5: копейки «половина вверх», скидка в копии пропорционально, склады только в строках — отказ", async () => {
    // 0.57 с НДС 20%: 0.095 → 0.10 (двоичная дробь давала 0.09).
    const s = setup();
    const r = await s.call(SHIP, { ...base, lines: [line(1, 0.57, { vatRate: "НДС20" })] });
    expect(payloadOf(r)["Товары"]).toMatchObject([{ Сумма: 0.57, СуммаНДС: 0.1 }]);
    // Копия: скидка суммой масштабируется вместе с суммой.
    const ref = await made(s, [line(2, 100)]);
    Object.assign((s.docs.get(ref)!.body["Товары"] as Record<string, unknown>[])[0]!, {
      Сумма: 180,
      СуммаСНДС: 180,
      СуммаРучнойСкидки: 20,
      ПроцентРучнойСкидки: 10,
    });
    await s.create("write.document.copy_document", {
      entitySet: SET,
      ref,
      lines: [{ lineNumber: 1, fields: { Количество: 4 } }],
    });
    expect(s.posts.at(-1)!.body["Товары"]).toEqual([
      expect.objectContaining({ Количество: 4, Сумма: 360, СуммаРучнойСкидки: 40, ПроцентРучнойСкидки: 10 }),
    ]);
    // Склада в шапке нет, а у строк есть свой: склад новой строки не угадать — отказ без записи.
    const multi = await made(s);
    s.docs.get(multi)!.body["Склад_Key"] = "00000000-0000-0000-0000-000000000000";
    const patches = s.patches.length;
    const add = await s.call("write.document.add_document_line", {
      entitySet: SET,
      ref: multi,
      line: line(1, 5),
    });
    expect(add.isError).toBe(true);
    expect(add.text).toContain("Склад в шапке");
    expect(s.patches).toHaveLength(patches);
  });

  it("замена строк: та же сборка, итог и единственный этап графика пересчитаны; строки сверены", async () => {
    const s = setup();
    const ref = await made(s);
    const r = await twice(s, "write.document.update_document_lines", {
      entitySet: SET,
      ref,
      lines: [line(3, 10)],
    });
    expect(r.isError, r.text).toBeFalsy();
    expect(s.patches.at(-1)).toMatchObject({
      Товары: [{ Количество: 3, КоличествоУпаковок: 3, Сумма: 30, Склад_Key: WH }],
      СуммаДокумента: 30,
      ЭтапыГрафикаОплаты: [{ СуммаПлатежа: 30, СуммаВзаиморасчетов: 30 }],
    });
    expect(r.sc).toMatchObject({ verified: true });
  });

  it("добавление сохраняет поля прежних строк; удаление уносит серии позиции; проведённый — отказ", async () => {
    const s = setup();
    const ref = await made(s, [line(1, 10), line(1, 20, { nomenclatureRef: SVC, seriesRef: SERIES })]);
    const add = await twice(s, "write.document.add_document_line", { entitySet: SET, ref, line: line(1, 5) });
    expect(add.isError, add.text).toBeFalsy();
    expect(s.patches.at(-1)!["Товары"]).toMatchObject([{ ОбъектРасчетов_Key: g("8", 1) }, {}, { Сумма: 5 }]);
    expect(s.patches.at(-1)).toMatchObject({ СуммаДокумента: 35 });
    const rm = await twice(s, "write.document.remove_document_line", { entitySet: SET, ref, lineNumber: 2 });
    expect(rm.isError, rm.text).toBeFalsy();
    expect(s.patches.at(-1)).toMatchObject({ Серии: [], СуммаДокумента: 15 });
    await s.post(SET, ref);
    const posted = await s.call("write.document.add_document_line", {
      entitySet: SET,
      ref,
      line: line(1, 5),
    });
    expect(posted.text).toContain("проведён");
  });

  it("график оплаты из нескольких этапов — отказ; счёт-фактура строк не правит", async () => {
    const s = setup();
    const ref = await made(s);
    s.docs.get(ref)!.body["ЭтапыГрафикаОплаты"] = [{ СуммаПлатежа: 50 }, { СуммаПлатежа: 50 }];
    const r = await s.call("write.document.update_document_lines", {
      entitySet: SET,
      ref,
      lines: [line(1, 10)],
    });
    expect(r.text).toContain("несколько этапов");
    const iss = await s.create("write.sales.create_issued_invoice", { baseDocumentRef: BASE_SALE });
    const sf = await s.call("write.document.add_document_line", {
      entitySet: "Document_СчетФактураВыданный",
      ref: iss.ref,
      line: line(1, 5),
    });
    expect(sf.text).toContain("строк товаров нет");
  });

  it("поля: белый список, проведённый — отказ; копия — без объектов расчётов, количество — пропорционально суммам", async () => {
    const s = setup();
    const ref = await made(s);
    const ok = await twice(s, "write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { Комментарий: "ZZ", Статус: "Отгружено" },
    });
    expect(ok.isError, ok.text).toBeFalsy();
    const bad = await s.call("write.entity.update_entity", {
      entitySet: SET,
      ref,
      fields: { Склад_Key: WH },
    });
    expect(bad.text).toContain("Склад_Key");
    const { preview } = await s.create("write.document.copy_document", {
      entitySet: SET,
      ref,
      date: "2026-11-25",
      lines: [{ lineNumber: 1, fields: { Количество: 4 } }],
    });
    const body = s.posts.at(-1)!.body;
    expect(body["Товары"]).toEqual([
      expect.objectContaining({ Количество: 4, КоличествоУпаковок: 4, Сумма: 200, СуммаСНДС: 200 }),
    ]);
    expect(JSON.stringify(body)).not.toContain("ОбъектРасчетов_Key");
    expect(body).toMatchObject({
      СуммаДокумента: 200,
      ЭтапыГрафикаОплаты: [{ ДатаПлатежа: "2026-11-25T00:00:00", СуммаПлатежа: 200 }],
    });
    expect(JSON.stringify(preview.sc["notes"])).toContain("без объектов расчётов");
    const badRow = await s.call("write.document.copy_document", {
      entitySet: SET,
      ref,
      lines: [{ lineNumber: 1, fields: { Цена: 1 } }],
    });
    expect(badRow.text).toContain("Цена");
  });
});

describe("продажи и закупки в Бухгалтерии — как раньше", () => {
  it("параметры УТ — отказ; заказы — отказ", async () => {
    const s = setup({ schema: "bp30" });
    const a = await s.call(SHIP, {
      counterpartyRef: CP,
      agreementRef: AGR,
      lines: [{ nomenclatureRef: ITEM, quantity: 1, price: 10 }],
    });
    expect(a.text).toContain("только для Управления торговлей");
    const b = await s.call("write.sales.create_customer_order", {
      counterpartyRef: CP,
      lines: [{ nomenclatureRef: ITEM, quantity: 1, price: 10 }],
    });
    expect(b.text).toContain("документ Управления торговлей");
    for (const r of [a, b]) expect(r.isError).toBe(true);
    expect(s.posts).toEqual([]);
  });
});
