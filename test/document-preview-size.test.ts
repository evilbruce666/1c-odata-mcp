import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";
import { SNAPSHOT_REF as id } from "./helpers/metadata-snapshot.js";

/**
 * Предпросмотр документа любого размера укладывается в предел ответа (ODATA_MAX_RESPONSE_CHARS), operationId в нём
 * есть; итоги — по всем строкам. Корпус: все инструменты создания документов Бухгалтерии и Управления торговлей и
 * правка строк, документ из N разных позиций с названиями заданной длины.
 */

const item = (i: number) => `c0c0c0c0-0000-4000-8000-${String(i).padStart(12, "0")}`;
const W2 = "a0a0a0a0-0000-4000-8000-000000000002";
const DOC = "d0d0d0d0-0000-4000-8000-000000000001";

interface Opts {
  schema: FakeSchema;
  n: number;
  nameLen: number;
  limit?: number;
  /** Позиций в справочнике (по умолчанию — столько, сколько строк). */
  catalog?: number;
  /** Вид всех позиций ведёт серии (а серии в строках не указаны). */
  series?: boolean;
  /** Дополнительные переменные окружения сервера. */
  env?: Record<string, string>;
}

const KIND = "b0b0b0b0-0000-4000-8000-000000000001";

function setup(o: Opts) {
  const name = (i: number) =>
    `Позиция ${String(i).padStart(4, "0")} `.padEnd(o.nameLen, "ж").slice(0, o.nameLen);
  const known = new Set(Array.from({ length: Math.max(o.catalog ?? 0, o.n) }, (_, i) => item(i + 1)));
  const docRows = Array.from({ length: o.n }, (_, i) => ({
    LineNumber: i + 1,
    Номенклатура_Key: item(i + 1),
    Количество: 1 + (i % 7),
    Цена: 10 + i,
    Сумма: (1 + (i % 7)) * (10 + i),
  }));
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-dps-")),
      ...(o.limit ? { ODATA_MAX_RESPONSE_CHARS: String(o.limit) } : {}),
      ...o.env,
    },
    {
      schema: o.schema,
      data: (q: FakeQuery) => {
        if (q.fn) return undefined;
        if (q.base === "Document_РеализацияТоваровУслуг" && q.params.get("$select") === null)
          return [
            { Ref_Key: DOC, Posted: false, DeletionMark: false, Товары: docRows },
            ...[...posted.values()].map((b) => ({ ...b, Posted: false, DeletionMark: false })),
          ];
        if (
          o.series &&
          q.base === "Catalog_ВидыНоменклатуры" &&
          q.params.get("$select") === "Ref_Key,ИспользоватьСерии"
        )
          return [{ Ref_Key: KIND, ИспользоватьСерии: true }];
        if (q.base !== "Catalog_Номенклатура") return undefined;
        // Только чтения названий: иначе — синтетическая строка схемы (сборщику нужны все её поля).
        const select = (q.params.get("$select") ?? "").split(",").filter(Boolean);
        const simple = new Set([
          ...["Ref_Key", "Description", "Code", "Number", "Date", "ВидНоменклатуры_Key"],
          // Сведения позиций (inventory: справочник целиком и по ссылкам).
          ...["Артикул", "Parent_Key", "DeletionMark", "ИспользованиеХарактеристик", "IsFolder"],
        ]);
        if (!select.length || !select.every((f) => simple.has(f))) return undefined;
        const f = q.params.get("$filter") ?? "";
        const refs = [...f.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());
        const rows = (refs.length ? refs.filter((r) => known.has(r)) : [...known]).map((r) => {
          const i = Number(r.slice(-12));
          const all: Record<string, unknown> = {
            Ref_Key: r,
            Description: name(i),
            Code: String(i),
            DeletionMark: false,
            IsFolder: false,
            ВидНоменклатуры_Key: o.series ? KIND : null,
          };
          return Object.fromEntries(select.map((k) => [k, all[k] ?? null]));
        });
        return rows; // $skip/$top применяет подделка
      },
    },
  );
  const log: string[] = [];
  const posted = new Map<string, Record<string, unknown>>();
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    const method = (init?.method ?? "GET").toUpperCase();
    log.push(`${method} ${decodeURIComponent(url.replace(/^.*?standard\.odata\//, ""))}`);
    if (method === "POST" && typeof init?.body === "string" && !/\/(Post|Unpost)\?/.test(url)) {
      const body = JSON.parse(init.body) as Record<string, unknown>;
      if (typeof body["Ref_Key"] === "string") posted.set(body["Ref_Key"].toLowerCase(), body);
    }
    return inner(input, init);
  });
  const call = async (tool: string, args: Record<string, unknown>) => {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({ name: tool, arguments: args })) as {
      isError?: boolean;
      structuredContent?: Record<string, unknown>;
      content: { text: string }[];
    };
    return {
      isError: !!r.isError,
      sc: r.structuredContent ?? {},
      text: r.content.map((c) => c.text).join(""),
    };
  };
  return { call, log, fake, posted };
}

const lines = (n: number, priced = true) =>
  Array.from({ length: n }, (_, i) => ({
    nomenclatureRef: item(i + 1),
    quantity: 1 + (i % 7),
    ...(priced ? { price: 10 + i } : {}),
  }));

type Case = { schema: FakeSchema; tool: string; args: (n: number) => Record<string, unknown> };
const NOVAT = { vatTaxation: "ПродажаНеОблагаетсяНДС" };
export const CASES: Case[] = [
  ...(
    [
      ["write.sales.create_invoice", (n) => ({ counterpartyRef: id, lines: lines(n) })],
      ["write.sales.create_shipment", (n) => ({ counterpartyRef: id, contractRef: id, lines: lines(n) })],
      ["write.sales.create_act", (n) => ({ counterpartyRef: id, contractRef: id, lines: lines(n) })],
      ["write.purchase.create_purchase", (n) => ({ counterpartyRef: id, contractRef: id, lines: lines(n) })],
      ["write.purchase.create_supplier_invoice", (n) => ({ counterpartyRef: id, lines: lines(n) })],
      [
        "write.warehouse.create_transfer",
        (n) => ({ fromWarehouse: id, toWarehouse: W2, lines: lines(n, false) }),
      ],
      ["write.warehouse.create_writeoff", (n) => ({ warehouse: id, lines: lines(n, false) })],
      ["write.warehouse.create_surplus", (n) => ({ warehouse: id, lines: lines(n) })],
      ["write.warehouse.create_return_from_customer", (n) => ({ counterpartyRef: id, lines: lines(n) })],
      ["write.warehouse.create_return_to_supplier", (n) => ({ counterpartyRef: id, lines: lines(n) })],
      [
        "write.document.add_document_line",
        () => ({ entitySet: "Document_РеализацияТоваровУслуг", ref: DOC, line: lines(1)[0] }),
      ],
    ] as [string, Case["args"]][]
  ).map(([tool, args]) => ({ schema: "bp30" as FakeSchema, tool, args })),
  ...(
    [
      ["write.sales.create_invoice", (n) => ({ counterpartyRef: id, lines: lines(n) })],
      [
        "write.sales.create_shipment",
        (n) => ({ counterpartyRef: id, contractRef: id, ...NOVAT, lines: lines(n) }),
      ],
      ["write.sales.create_customer_order", (n) => ({ counterpartyRef: id, ...NOVAT, lines: lines(n) })],
      ["write.purchase.create_purchase", (n) => ({ counterpartyRef: id, contractRef: id, lines: lines(n) })],
      ["write.purchase.create_supplier_order", (n) => ({ counterpartyRef: id, lines: lines(n) })],
      [
        "write.warehouse.create_transfer",
        (n) => ({ fromWarehouse: id, toWarehouse: W2, lines: lines(n, false) }),
      ],
      [
        "write.warehouse.create_writeoff",
        (n) => ({ warehouse: id, expenseItem: id, lines: lines(n, false) }),
      ],
      [
        "write.warehouse.create_surplus",
        (n) => ({ warehouse: id, priceKind: id, incomeItem: id, lines: lines(n) }),
      ],
      [
        "write.warehouse.create_return_from_customer",
        (n) => ({ counterpartyRef: id, ...NOVAT, lines: lines(n) }),
      ],
      [
        "write.document.add_document_line",
        () => ({ entitySet: "Document_РеализацияТоваровУслуг", ref: DOC, line: lines(1)[0] }),
      ],
      [
        "write.document.update_document_lines",
        (n) => ({ entitySet: "Document_РеализацияТоваровУслуг", ref: DOC, lines: lines(n) }),
      ],
    ] as [string, Case["args"]][]
  ).map(([tool, args]) => ({ schema: "ut11" as FakeSchema, tool, args })),
  {
    schema: "bp30",
    tool: "write.document.update_document_lines",
    args: (n) => ({ entitySet: "Document_РеализацияТоваровУслуг", ref: DOC, lines: lines(n) }),
  },
];

/** Случайные ref и operationId — метками: сравнивается содержание, а не случайные значения. */
function normalize(text: string, sc: Record<string, unknown>): string {
  let out = text;
  for (const v of [sc["operationId"], sc["ref"]])
    if (typeof v === "string") out = out.split(v).join(`<${v === sc["ref"] ? "ref" : "op"}>`);
  return out;
}

const DUMP = process.env["DOC_DUMP"];
const DUMP_N = (process.env["DOC_DUMP_N"] ?? "1,5,20,50,80").split(",").map(Number);

describe.runIf(DUMP)("дамп предпросмотров (сравнение версий)", () => {
  it("дамп", async () => {
    mkdirSync(DUMP!, { recursive: true });
    for (const c of CASES)
      for (const n of DUMP_N)
        for (const nameLen of [10, 150])
          for (const limit of [45_000, 10_000]) {
            const s = setup({ schema: c.schema, n, nameLen, limit });
            const r = await s.call(c.tool, c.args(n));
            const sc = JSON.parse(JSON.stringify(r.sc)) as Record<string, unknown>;
            writeFileSync(
              join(DUMP!, `${c.schema}-${c.tool}-${n}-${nameLen}-${limit}.json`),
              JSON.stringify({
                isError: r.isError,
                len: r.text.length,
                sc: JSON.parse(normalize(JSON.stringify(sc), r.sc)),
                requests: s.log.map((l) => normalize(l, r.sc)),
              }),
            );
          }
  }, 600_000);
});

type Res = Awaited<ReturnType<ReturnType<typeof setup>["call"]>>;
const viewRows = (r: Res, table = "Товары") =>
  (r.sc["view"] as { tables?: Record<string, Record<string, unknown>[]> } | undefined)?.tables?.[table] ?? [];
const GUID_TEXT = /^[0-9a-f]{8}-[0-9a-f]{4}-/;

describe("предпросмотр документа: предел ответа на любом числе строк", () => {
  const grid = CASES.flatMap((c) =>
    [81, 300, 1000].flatMap((n) =>
      [10, 150].flatMap((nameLen) => [45_000, 10_000].map((limit) => ({ ...c, n, nameLen, limit }))),
    ),
  );
  it.each(grid)(
    "$schema $tool: $n строк, названия $nameLen, предел $limit",
    async (c) => {
      const s = setup({ schema: c.schema, n: c.n, nameLen: c.nameLen, limit: c.limit });
      const args = c.args(c.n);
      const r = await s.call(c.tool, args);
      if (r.isError) {
        // Отказ — тот же, что и на одной строке (у сборщика нет данных), а не из-за размера.
        const one = await setup({ schema: c.schema, n: 1, nameLen: c.nameLen, limit: c.limit }).call(
          c.tool,
          c.args(1),
        );
        expect(one.isError, r.text.slice(0, 300)).toBe(true);
        return;
      }
      expect(r.text.length).toBeLessThanOrEqual(c.limit);
      // Правка строк — без журнала операций и без operationId; у остальных он первым полем.
      if (c.tool !== "write.document.update_document_lines")
        expect(r.text.startsWith(`{"operationId":"${String(r.sc["operationId"])}"`)).toBe(true);
      const total = c.tool === "write.document.add_document_line" ? c.n + 1 : c.n;
      const own = c.tool !== "write.document.add_document_line";
      const rows = (r.sc["payload"] ?? r.sc["fields"]) as Record<string, unknown>;
      const summary = r.sc["summary"] as Record<string, Record<string, number>> | undefined;
      // Таблица строк документа: у акта — «Услуги», у остальных — «Товары». Счёт УТ строк товаров не хранит.
      const main = ["Товары", "Услуги"].find((k) => k in rows || (summary && k in summary));
      if (!main) return;
      if (!summary) {
        // Не сокращён: таблица целиком и в машинной части, и в виде.
        expect(rows[main]).toHaveLength(total);
        expect(viewRows(r, main)).toHaveLength(total);
        return;
      }
      expect(summary[main]!["lines"]).toBe(total);
      if (own) {
        const qty = (args["lines"] as { quantity: number }[]).reduce((t, l) => t + l.quantity, 0);
        expect(summary[main]!["Количество"]).toBe(qty);
      }
      const shown = viewRows(r, main);
      if (r.sc["linesComplete"] === false) {
        const n = (r.sc["linesShown"] as Record<string, number>)[main]!;
        expect(shown.length).toBe(n);
        expect(n).toBeGreaterThan(0);
        expect(n).toBeLessThan(total);
      } else expect(shown.length).toBe(total);
      // Показанные строки — начало документа, по порядку.
      if (own)
        shown.forEach((row, i) =>
          expect(row["Количество"], `строка ${i + 1}`).toBe(
            (args["lines"] as { quantity: number }[])[i]!.quantity,
          ),
        );
    },
    60_000,
  );
});

describe("предпросмотр документа: неухудшение и названия", () => {
  it("документ, который помещается, — без сокращения и без новых полей", async () => {
    const s = setup({ schema: "ut11", n: 20, nameLen: 10 });
    const r = await s.call("write.sales.create_shipment", CASES[12]!.args(20));
    expect(r.isError, r.text.slice(0, 300)).toBe(false);
    for (const k of ["summary", "rowsInView", "linesComplete", "linesShown"])
      expect(r.sc).not.toHaveProperty(k);
    expect((r.sc["payload"] as Record<string, unknown[]>)["Товары"]).toHaveLength(20);
    expect(viewRows(r)).toHaveLength(20);
  });

  it("ответ вплотную к пределу: запас под operationId учтён (t−1, t, t+1)", async () => {
    // Размер без operationId при пределе по умолчанию, затем предел так, чтобы ответ лёг ровно на границу.
    const s0 = setup({ schema: "ut11", n: 30, nameLen: 150, limit: 45_000 });
    const base = await s0.call("write.sales.create_shipment", CASES[12]!.args(30));
    const t = base.text.length;
    for (const limit of [t - 1, t, t + 1]) {
      const r = await setup({ schema: "ut11", n: 30, nameLen: 150, limit }).call(
        "write.sales.create_shipment",
        CASES[12]!.args(30),
      );
      expect(r.text.length, `предел ${limit}`).toBeLessThanOrEqual(limit);
      expect(r.sc).toHaveProperty(limit >= t ? "payload.Товары" : "rowsInView");
    }
  });

  it("больше 50 позиций — названия у всех строк: пачками по 50, с 200 — справочником целиком", async () => {
    for (const n of [80, 300]) {
      const s = setup({ schema: "ut11", n, nameLen: 10 });
      const r = await s.call("write.sales.create_shipment", CASES[12]!.args(n));
      const rows = viewRows(r);
      expect(rows.length).toBeGreaterThan(0);
      for (const row of rows) expect(String(row["Номенклатура"])).not.toMatch(GUID_TEXT);
      const names = s.log.filter((l) =>
        /Catalog_Номенклатура\?.*\$select=Ref_Key,Description(,Code)?&/.test(l),
      );
      const orChunks = names.filter((l) => l.includes(" or "));
      if (n === 80) expect(orChunks).toHaveLength(2);
      else expect(orChunks).toHaveLength(0);
    }
  });
});

describe("проверка серий: номенклатура документа на 1000 позиций — справочником целиком", () => {
  it.each([
    "write.warehouse.create_transfer",
    "write.warehouse.create_writeoff",
    "write.warehouse.create_surplus",
  ])(
    "%s",
    async (tool) => {
      const c = CASES.find((x) => x.schema === "ut11" && x.tool === tool)!;
      const s = setup({ schema: "ut11", n: 1000, nameLen: 10, catalog: 1500 });
      const r = await s.call(tool, c.args(1000));
      expect(r.isError, r.text.slice(0, 300)).toBe(false);
      const kinds = s.log.filter((l) => /Catalog_Номенклатура\?.*ВидНоменклатуры_Key/.test(l));
      expect(kinds.length).toBeGreaterThan(0);
      for (const l of kinds) expect(l).not.toContain(" or ");
      expect(s.log.filter((l) => l.startsWith("GET")).length).toBeLessThanOrEqual(15);
    },
    60_000,
  );

  it("200 и меньше позиций — по ссылкам, как раньше", async () => {
    const c = CASES.find((x) => x.schema === "ut11" && x.tool === "write.warehouse.create_transfer")!;
    const s = setup({ schema: "ut11", n: 199, nameLen: 10 });
    await s.call(c.tool, c.args(199));
    const kinds = s.log.filter((l) => /Catalog_Номенклатура\?.*ВидНоменклатуры_Key/.test(l));
    expect(kinds).toHaveLength(10);
    for (const l of kinds) expect(l).toContain(" or ");
  });
});

describe("мелочи пачки А/Б", () => {
  it("ровно 200 позиций — названия пачками по 50, как на 199; справочник целиком — с 201", async () => {
    for (const [n, chunks] of [
      [199, 4],
      [200, 4],
      [201, 0],
    ] as const) {
      const s = setup({ schema: "ut11", n, nameLen: 10, catalog: 1500 });
      const r = await s.call("write.sales.create_shipment", CASES[12]!.args(n));
      expect(r.isError, r.text.slice(0, 300)).toBe(false);
      for (const row of viewRows(r)) expect(String(row["Номенклатура"])).not.toMatch(GUID_TEXT);
      const names = s.log.filter((l) => /Catalog_Номенклатура\?.*\$select=Ref_Key,Description,Code&/.test(l));
      expect(
        names.filter((l) => l.includes(" or ")),
        `${n} позиций`,
      ).toHaveLength(chunks);
      // Сведения позиций по ссылкам (пачки по 20) — только у справочника целиком с откатом, не на 200.
      const attrs = s.log.filter((l) => /Catalog_Номенклатура\?.*Артикул.*Ref_Key eq/.test(l));
      expect(attrs, `${n} позиций`).toHaveLength(0);
    }
  });

  it("справочник больше предела выборки в проверке серий: второй предпросмотр сразу по ссылкам", async () => {
    const c = CASES.find((x) => x.schema === "ut11" && x.tool === "write.warehouse.create_transfer")!;
    const s = setup({
      schema: "ut11",
      n: 300,
      nameLen: 10,
      catalog: 1500,
      env: { ODATA_ANALYTICS_MAX_ROWS: "1000" },
    });
    const whole = () =>
      s.log.filter((l) => /Catalog_Номенклатура\?.*ВидНоменклатуры_Key/.test(l) && !l.includes(" or "))
        .length;
    const r1 = await s.call(c.tool, c.args(300));
    expect(r1.isError, r1.text.slice(0, 300)).toBe(false);
    const first = whole();
    expect(first).toBeGreaterThan(0); // попытка целиком была и не удалась
    s.log.length = 0;
    const r2 = await s.call(c.tool, c.args(300));
    expect(r2.isError, r2.text.slice(0, 300)).toBe(false);
    expect(whole()).toBe(0);
    expect(r2.sc["notes"]).toEqual(r1.sc["notes"]);
  });

  it("summary: поле есть не во всех строках — withValue называет, по скольким сложено", async () => {
    const { linesSummary } = await import("../src/tools/preview-fit.js");
    const rows = {
      Товары: [{ Сумма: 1.1, Количество: 1 }, { Количество: 2 }, { Сумма: 2.2, Количество: 3 }],
    };
    expect(linesSummary(rows, ["Товары"])).toEqual({
      Товары: { lines: 3, Количество: 6, Сумма: 3.3, withValue: { Сумма: 2 } },
    });
    expect(linesSummary({ Товары: [{ Сумма: 1 }, { Сумма: 2 }] }, ["Товары"])).toEqual({
      Товары: { lines: 2, Сумма: 3 },
    });
  });
});

describe("счета строк Бухгалтерии: регистр один раз, вид позиций пачками", () => {
  it.each(
    CASES.filter((c) => c.schema === "bp30" && c.tool !== "write.document.add_document_line").map((c) => [
      c.tool,
      c,
    ]),
  )("%s на 80 строк", async (_tool, c) => {
    const s = setup({ schema: "bp30", n: 80, nameLen: 10 });
    const r = await s.call(c.tool, c.args(80));
    expect(r.isError, r.text.slice(0, 300)).toBe(false);
    expect(s.log.filter((l) => /^GET Catalog_Номенклатура\(guid/.test(l))).toHaveLength(0);
    expect(
      s.log.filter((l) => l.includes("InformationRegister_СчетаУчетаНоменклатуры")).length,
    ).toBeLessThanOrEqual(1);
    expect(s.log.filter((l) => l.startsWith("GET")).length).toBeLessThanOrEqual(20);
  });
});

describe("сокращение предпросмотра делит место между таблицами", () => {
  const SERIES = "e0e0e0e0-0000-4000-8000-000000000001";
  it.each([
    [300, 45_000],
    [1000, 45_000],
    [300, 10_000],
    [1000, 10_000],
  ])("реализация УТ с сериями, %i строк, предел %i: и Товары, и Серии показаны долей", async (n, limit) => {
    const s = setup({ schema: "ut11", n, nameLen: 150, limit });
    const args = CASES[12]!.args(n);
    const r = await s.call("write.sales.create_shipment", {
      ...args,
      lines: (args["lines"] as Record<string, unknown>[]).map((l) => ({ ...l, seriesRef: SERIES })),
    });
    expect(r.isError, r.text.slice(0, 300)).toBe(false);
    expect(r.text.length).toBeLessThanOrEqual(limit);
    const shown = r.sc["linesShown"] as Record<string, number>;
    if (process.env["DOC_SHOW"])
      console.log(n, limit, JSON.stringify(shown), viewRows(r).length, viewRows(r, "Серии").length);
    const goods = viewRows(r).length;
    const series = viewRows(r, "Серии").length;
    expect(goods, JSON.stringify(shown)).toBeGreaterThan(1);
    expect(goods).toBeGreaterThanOrEqual(series * 0.2);
    expect(series).toBeGreaterThanOrEqual(goods * 0.2);
    const sum = r.sc["summary"] as Record<string, Record<string, number>>;
    expect(sum["Товары"]!["lines"]).toBe(n);
  });
});

describe("предпросмотр документа: уровень 1 и соседи на 1000 строк", () => {
  it("строки машинной части, повторённые в view, не повторяются — все строки видны", async () => {
    // 80 строк с длинными названиями: вдвоём (payload и view) не помещаются, по одному разу — да.
    const s = setup({ schema: "ut11", n: 80, nameLen: 150 });
    const r = await s.call("write.sales.create_shipment", CASES[12]!.args(80));
    expect(r.sc["rowsInView"]).toContain("Товары");
    expect(r.sc).not.toHaveProperty("linesComplete");
    expect(viewRows(r)).toHaveLength(80);
    expect(r.sc["payload"]).not.toHaveProperty("Товары");
    expect((r.sc["summary"] as Record<string, Record<string, number>>)["Товары"]!["lines"]).toBe(80);
    expect(r.text.length).toBeLessThanOrEqual(45_000);
  });

  it("реализация УТ на 1000 строк: подтверждение пишет все строки; проведение, чтение, статус, повтор — в пределе", async () => {
    const s = setup({ schema: "ut11", n: 1000, nameLen: 150 });
    const args = CASES[12]!.args(1000);
    const p = await s.call("write.sales.create_shipment", args);
    expect(p.sc["linesComplete"]).toBe(false);
    const confirm = { ...args, confirm: true, operationId: p.sc["operationId"] };
    const c = await s.call("write.sales.create_shipment", confirm);
    expect(c.isError, c.text.slice(0, 300)).toBe(false);
    const ref = String(c.sc["ref"]).toLowerCase();
    expect((s.posted.get(ref)?.["Товары"] as unknown[]).length).toBe(1000);
    const entitySet = "Document_РеализацияТоваровУслуг";
    const results = [
      c,
      await s.call("write.sales.create_shipment", confirm),
      await s.call("write.document.post_document", { entitySet, ref, post: true }),
      await s.call("write.document.post_document", { entitySet, ref, post: true, confirm: true }),
      await s.call("read.document.get_document", { entitySet, ref }),
      await s.call("write.operation.status", { operationId: p.sc["operationId"] }),
    ];
    expect(results[1]!.sc["replayed"]).toBe(true);
    for (const r of results) {
      expect(r.text.length).toBeLessThanOrEqual(45_000);
    }
    expect(results.slice(2).map((r) => r.isError)).toEqual([false, false, false, false]);
  }, 60_000);
});

describe("заметки предпросмотра не растут со строками", () => {
  it.each([
    "write.warehouse.create_transfer",
    "write.warehouse.create_writeoff",
    "write.warehouse.create_surplus",
  ])(
    "%s: 1000 позиций без серии у вида с сериями — предпросмотр и подтверждение в пределе",
    async (tool) => {
      // Живая проверка на копии УТ (23.11.2026): заметка перечисляла все позиции — 58 тыс. знаков на 1000 строк.
      const c = CASES.find((x) => x.schema === "ut11" && x.tool === tool)!;
      const s = setup({ schema: "ut11", n: 1000, nameLen: 60, series: true });
      const p = await s.call(tool, c.args(1000));
      expect(p.isError, p.text.slice(0, 300)).toBe(false);
      const note = (p.sc["notes"] as string[]).find((x) => x.includes("ведёт серии"))!;
      expect(note).toContain("ещё 990");
      expect(p.text.length).toBeLessThanOrEqual(45_000);
      const d = await s.call(tool, { ...c.args(1000), confirm: true, operationId: p.sc["operationId"] });
      expect(d.isError, d.text.slice(0, 300)).toBe(false);
      expect(d.text.length).toBeLessThanOrEqual(45_000);
    },
    60_000,
  );
});
