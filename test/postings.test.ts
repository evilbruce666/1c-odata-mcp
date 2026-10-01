import { describe, it, expect, vi, afterEach } from "vitest";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { Connection, type ServerContext } from "../src/context.js";
import { recorderFilter, REGISTER_RECORDS } from "../src/odata/accounting.js";
import { AggregateOverflowError } from "../src/odata/aggregate.js";
import { getDocumentPostings, registerRegisterTools } from "../src/tools/registers.js";
import { getDocumentPostingsResultSchema } from "../src/schemas/output.js";

/**
 * Проводки документа: настоящий Connection/ODataClient, подменён только global fetch.
 * Так проверяются и разбор $metadata, и то, что уходят ТОЛЬКО GET-запросы.
 *
 * ВАЖНО: фейковая 1С реализует отбор `Recorder eq cast(guid'…', 'Document_…')` так,
 * как мы ОЖИДАЕМ от реальной. Эти тесты не доказывают, что живая 1С понимает этот
 * синтаксис — только что код строит его и корректно обрабатывает ответ.
 */

const BASE = "http://1c.test/base/odata/standard.odata/";
const DOC = "Document_РегламентнаяОперация";
const DOC_REF = "919a75d1-7f6a-11f1-86c3-74563c4bf0d1";
const OTHER_REF = "aaaaaaaa-0000-0000-0000-000000000001";
const ORG = "e726d309-13a0-11e3-ae87-e8039ae81ce9";
const REG = "AccountingRegister_Хозрасчетный";
const RECS = `${REG}_RecordType`;
const CHART = "ChartOfAccounts_Хозрасчетный";
const ORGS = "Catalog_Организации";
const EMPTY = "00000000-0000-0000-0000-000000000000";
const acc = (n: number): string => `00000000-0000-0000-0000-${String(n).padStart(12, "0")}`;

const A20 = acc(20);
const A26 = acc(26);
const A9021 = acc(9021);
const A9008 = acc(9008);
const A001 = acc(1); // забалансовый
const CHART_ROWS = [
  { Ref_Key: A20, Code: "20.01", Description: "Основное производство" },
  { Ref_Key: A26, Code: "26", Description: "Общехозяйственные расходы" },
  { Ref_Key: A9021, Code: "90.02.1", Description: "Себестоимость продаж" },
  { Ref_Key: A9008, Code: "90.08.1", Description: "Управленческие расходы" },
  { Ref_Key: A001, Code: "001", Description: "Арендованные ОС" },
];

const prop = (name: string, type = "Edm.String"): string => `<Property Name="${name}" Type="${type}"/>`;
const RECORD_PROPS = [
  "Recorder",
  "Recorder_Type",
  "Period",
  "LineNumber",
  "Active",
  "AccountDr_Key",
  "AccountCr_Key",
  "Организация_Key",
  "Сумма",
  "КоличествоDr",
  "КоличествоCr",
  "ВалютнаяСуммаDr",
  "ВалютнаяСуммаCr",
  "ПодразделениеDr_Key",
  "ПодразделениеCr_Key",
  "Содержание",
  "ExtDimensionDr1",
  "ExtDimensionDr1_Type",
  "ExtDimensionDr2",
  "ExtDimensionDr2_Type",
  "ExtDimensionCr1",
  "ExtDimensionCr1_Type",
];

function metadataXml(recordProps: string[] = RECORD_PROPS): string {
  const type = (name: string, props: string[], keys = ["Ref_Key"]): string =>
    `<EntityType Name="${name}"><Key>${keys.map((k) => `<PropertyRef Name="${k}"/>`).join("")}</Key>${props
      .map((p) => prop(p))
      .join("")}</EntityType>`;
  const sets = [DOC, `${DOC}_Список`, "Catalog_Контрагенты", ORGS, CHART, REG, RECS];
  return `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx xmlns:edmx="http://schemas.microsoft.com/ado/2007/06/edmx" Version="1.0">
<edmx:DataServices m:DataServiceVersion="3.0" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">
<Schema Namespace="StandardODATA" xmlns="http://schemas.microsoft.com/ado/2009/11/edm">
${type(DOC, ["Ref_Key", "Number", "Date", "Posted", "DeletionMark", "Организация_Key", "ВидОперации", "Состояние", "Комментарий"])}
${type(`${DOC}_Список`, ["Ref_Key", "LineNumber", "Сумма"], ["Ref_Key", "LineNumber"])}
${type("Catalog_Контрагенты", ["Ref_Key", "Description"])}
${type(ORGS, ["Ref_Key", "Description"])}
${type(CHART, ["Ref_Key", "Code", "Description"])}
${type(REG, ["Recorder", "Recorder_Type", "RecordSet"], ["Recorder", "Recorder_Type"])}
${type(RECS, recordProps, ["Recorder", "Recorder_Type", "LineNumber"])}
<EntityContainer Name="EnterpriseV8" m:IsDefaultEntityContainer="true">
${sets.map((s) => `<EntitySet Name="${s}" EntityType="StandardODATA.${s}"/>`).join("\n")}
</EntityContainer>
</Schema>
</edmx:DataServices>
</edmx:Edmx>`;
}

type Row = Record<string, unknown>;
const rec = (line: number, dr: string, cr: string, sum: number | string, extra: Row = {}): Row => ({
  Recorder: DOC_REF,
  Recorder_Type: `StandardODATA.${DOC}`,
  Period: "2026-06-30T23:59:59",
  LineNumber: line,
  Active: true,
  AccountDr_Key: dr,
  AccountCr_Key: cr,
  Организация_Key: ORG,
  Сумма: sum,
  КоличествоDr: 0,
  КоличествоCr: 0,
  ВалютнаяСуммаDr: 0,
  ВалютнаяСуммаCr: 0,
  ПодразделениеDr_Key: EMPTY,
  ПодразделениеCr_Key: EMPTY,
  Содержание: "",
  ExtDimensionDr1: EMPTY,
  ExtDimensionDr1_Type: "",
  ExtDimensionDr2: EMPTY,
  ExtDimensionDr2_Type: "",
  ExtDimensionCr1: EMPTY,
  ExtDimensionCr1_Type: "",
  ...extra,
});

interface Fake1C {
  records?: Row[];
  recordProps?: string[];
  docExists?: boolean;
  cap?: number;
  /** Имитация 1С, которая проигнорировала $filter по регистратору. */
  ignoreRecorderFilter?: boolean;
}

interface Call {
  method: string;
  path: string;
}

function setup(opts: Fake1C = {}) {
  const calls: Call[] = [];
  const records = opts.records ?? [];
  const json = (body: unknown, status = 200): Response =>
    new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });

  vi.stubGlobal("fetch", async (input: string | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    const method = init?.method ?? "GET";
    const rel = decodeURIComponent(url.pathname.slice(new URL(BASE).pathname.length));
    calls.push({ method, path: decodeURIComponent(url.pathname + url.search) });
    const q = (k: string): string | undefined => url.searchParams.get(k) ?? undefined;
    const top = Number(q("$top") ?? 1e9);
    const skip = Number(q("$skip") ?? 0);
    const filter = q("$filter") ?? "";

    if (rel === "$metadata") return new Response(metadataXml(opts.recordProps), { status: 200 });

    if (rel.startsWith(`${DOC}(`)) {
      if (opts.docExists === false) {
        return json({ "odata.error": { code: "", message: { lang: "ru", value: "Не найдено" } } }, 404);
      }
      return json({
        Ref_Key: DOC_REF,
        Number: "0000-000061",
        Date: "2026-06-30T23:59:59",
        Posted: true,
        DeletionMark: false,
        Организация_Key: ORG,
        ВидОперации: "ЗакрытиеСчетов20_23_25_26",
        Состояние: "Выполнено",
        Комментарий: "",
      });
    }

    if (rel === RECS) {
      let pool = records;
      if (!opts.ignoreRecorderFilter) {
        const m = /^Recorder eq cast\(guid'([^']+)', '([^']+)'\)$/.exec(filter);
        if (!m) return json({ "odata.error": { message: { value: "bad filter" } } }, 400);
        pool = records.filter((r) => r["Recorder"] === m[1] && String(r["Recorder_Type"]).endsWith(m[2]!));
      }
      return json({ value: pool.slice(skip, skip + top) });
    }

    if (rel === CHART) {
      const keys = [...filter.matchAll(/Ref_Key eq guid'([^']+)'/g)].map((m) => m[1]);
      return json({ value: CHART_ROWS.filter((a) => keys.includes(a.Ref_Key)).slice(skip, skip + top) });
    }

    if (rel === ORGS) {
      return json({
        value: filter.includes(ORG)
          ? [{ Ref_Key: ORG, Description: 'ООО "Медицинский центр "Мегаполис"' }]
          : [],
      });
    }

    return json({ "odata.error": { message: { value: "unknown" } } }, 404);
  });

  // readOnly=false и writable=true — чтобы ПИШУЩИЙ запрос (если бы он был) дошёл
  // до fetch и попал в calls, а не был отсечён гардом клиента.
  const conn = new Connection(
    { name: "default", baseUrl: BASE, username: "u", password: "p", writable: true },
    {
      timeoutMs: 5_000,
      retries: 0,
      pageSize: 100,
      maxRows: 1000,
      analyticsMaxRows: opts.cap ?? 200_000,
      readOnly: false,
    },
  );
  const ctx = { db: () => conn } as unknown as ServerContext;
  return { conn, ctx, calls };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

const regCalls = (calls: Call[]): Call[] => calls.filter((c) => c.path.includes(`/${REG}`));

const SAMPLE = [
  rec(1, A9021, A20, 5000000.49, {
    ExtDimensionDr1: "bbbbbbbb-0000-0000-0000-000000000001",
    ExtDimensionDr1_Type: "StandardODATA.Catalog_НоменклатурныеГруппы",
    ПодразделениеCr_Key: "cccccccc-0000-0000-0000-000000000001",
  }),
  rec(2, A9021, A20, 2863638),
  rec(3, A9008, A26, 3243102.26, { КоличествоDr: 2.5 }),
  rec(4, A001, EMPTY, 100), // забалансовая: только Дт
];

async function callTool(ctx: ServerContext, args: Record<string, unknown>) {
  const server = new McpServer({ name: "t", version: "0" });
  registerRegisterTools(server, ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "c", version: "0" });
  await client.connect(b);
  try {
    return await client.callTool({ name: "read.accounting.get_document_postings", arguments: args });
  } finally {
    await client.close();
  }
}
const text = (r: Awaited<ReturnType<typeof callTool>>): string =>
  (r.content as Array<{ text: string }>)[0]!.text;

describe("recorderFilter", () => {
  it("один хелпер строит отбор по регистратору составного типа", () => {
    expect(recorderFilter(DOC, DOC_REF)).toBe(`Recorder eq cast(guid'${DOC_REF}', '${DOC}')`);
  });
  it("кривой GUID не попадает в фильтр", () => {
    expect(() => recorderFilter(DOC, "x' or 1 eq 1")).toThrow(/GUID/);
  });
});

describe("read.accounting.get_document_postings", () => {
  it("документ с несколькими проводками: шапка, проводки, счета, итоги, корреспонденции", async () => {
    const { ctx } = setup({ records: SAMPLE });
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    expect(r.isError).toBeFalsy();
    const s = r.structuredContent as Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
    expect(s["database"]).toBe("default");
    expect(s["document"]).toEqual({
      entitySet: DOC,
      ref: DOC_REF,
      number: "0000-000061",
      date: "2026-06-30T23:59:59",
      posted: true,
      deletionMark: false,
      organization: 'ООО "Медицинский центр "Мегаполис"',
      organizationRef: ORG,
      operation: "ЗакрытиеСчетов20_23_25_26",
      state: "Выполнено",
    });
    expect(s["postingsCount"]).toBe(4);
    // Дт: все 4 записи; Кт: без забалансовой (у неё нет счёта Кт).
    expect(s["debitTotal"]).toBe(11106840.75);
    expect(s["creditTotal"]).toBe(11106740.75);

    const p0 = s["postings"][0];
    expect(p0).toMatchObject({
      period: "2026-06-30T23:59:59",
      lineNumber: 1,
      active: true,
      amount: 5000000.49,
      organizationRef: ORG,
      debit: {
        accountCode: "90.02.1",
        accountName: "Себестоимость продаж",
        accountRef: A9021,
        dimensions: [
          { index: 1, type: "Catalog_НоменклатурныеГруппы", ref: "bbbbbbbb-0000-0000-0000-000000000001" },
        ],
      },
      credit: {
        accountCode: "20.01",
        accountRef: A20,
        divisionRef: "cccccccc-0000-0000-0000-000000000001",
      },
    });
    expect(p0.credit.dimensions).toBeUndefined(); // пустые субконто не выдумываем
    expect(p0.quantityDebit).toBeUndefined(); // нулевые количества не выводим
    expect(s["postings"][2].quantityDebit).toBe(2.5);
    expect(s["postings"][3].credit).toBeNull();

    expect(s["byCorrespondence"]).toEqual([
      { debitAccount: "90.02.1", creditAccount: "20.01", amount: 7863638.49, entries: 2 },
      { debitAccount: "90.08.1", creditAccount: "26", amount: 3243102.26, entries: 1 },
      { debitAccount: "001", creditAccount: null, amount: 100, entries: 1 },
    ]);
    expect(s["source"]).toEqual({ entitySet: RECS, filter: `Recorder eq cast(guid'${DOC_REF}', '${DOC}')` });
    expect(s["scan"]).toMatchObject({ rowsScanned: 4, windows: 1 });
    // Выходная схема (15): structuredContent валиден и напрямую по zod.
    expect(getDocumentPostingsResultSchema.safeParse(s).success).toBe(true);
  });

  it("отбор по регистратору уходит в 1С в $filter; весь регистр не запрашивается", async () => {
    const { ctx, calls } = setup({ records: [...SAMPLE, { ...rec(1, A20, A26, 1), Recorder: OTHER_REF }] });
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    expect(r.isError).toBeFalsy();
    const reg = regCalls(calls);
    expect(reg).toHaveLength(1);
    expect(reg[0]!.path).toContain(`/${RECS}?`);
    expect(reg[0]!.path).toContain(`$filter=Recorder eq cast(guid'${DOC_REF}', '${DOC}')`);
    expect(reg[0]!.path).toContain("$orderby=LineNumber asc");
    expect((r.structuredContent as { postingsCount: number }).postingsCount).toBe(4);
  });

  it("счета Дт/Кт резолвятся одним пакетным запросом к плану счетов, а не по проводке", async () => {
    const { ctx, calls } = setup({ records: SAMPLE });
    await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    const chart = calls.filter((c) => c.path.includes(`/${CHART}`));
    expect(chart).toHaveLength(1);
    for (const k of [A20, A26, A9021, A9008, A001]) expect(chart[0]!.path).toContain(`guid'${k}'`);
    expect(chart[0]!.path).not.toContain(EMPTY);
  });

  it("деньги копятся в копейках: 1000 × 0.10 = ровно 100.00", async () => {
    const records = Array.from({ length: 1000 }, (_, i) => rec(i + 1, A9021, A20, i % 2 ? 0.1 : "0.1"));
    const { conn } = setup({ records });
    const res = await getDocumentPostings(conn, DOC, DOC_REF);
    expect(res.debitTotal).toBe(100);
    expect(res.creditTotal).toBe(100);
    expect(res.byCorrespondence).toEqual([
      { debitAccount: "90.02.1", creditAccount: "20.01", amount: 100, entries: 1000 },
    ]);
  });

  it("проведённый документ без записей — пустые массивы и нули", async () => {
    const { ctx } = setup({ records: [] });
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    expect(r.isError).toBeFalsy();
    const s = r.structuredContent as Record<string, unknown>;
    expect(s["postingsCount"]).toBe(0);
    expect(s["postings"]).toEqual([]);
    expect(s["byCorrespondence"]).toEqual([]);
    expect(s["debitTotal"]).toBe(0);
    expect(s["creditTotal"]).toBe(0);
  });

  it("неактивные записи видны в postings, но не входят в итоги", async () => {
    const { conn } = setup({ records: [rec(1, A9021, A20, 10), rec(2, A9021, A20, 5, { Active: false })] });
    const res = await getDocumentPostings(conn, DOC, DOC_REF);
    expect(res.postingsCount).toBe(2);
    expect(res.postings[1]!.active).toBe(false);
    expect(res.debitTotal).toBe(10);
    expect(res.byCorrespondence[0]!.entries).toBe(1);
    expect(res.note).toMatch(/Неактивных записей: 1/);
  });

  it("невалидный GUID — отказ до любого запроса к 1С", async () => {
    const { ctx, conn, calls } = setup();
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: "not-a-guid" });
    expect(r.isError).toBe(true);
    await expect(getDocumentPostings(conn, DOC, "919a75d1' or 1 eq 1")).rejects.toThrow(/GUID/);
    expect(calls).toHaveLength(0);
  });

  it("несуществующий документ — явная ошибка not_found, а не пустые проводки", async () => {
    const { ctx, calls } = setup({ docExists: false, records: SAMPLE });
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: OTHER_REF });
    expect(r.isError).toBe(true);
    expect(text(r)).toMatch(/\[not_found\].*не найден/);
    expect(regCalls(calls)).toHaveLength(0);
  });

  it("нет обязательного поля в записях — громкая ошибка с именами полей, без значений", async () => {
    const broken = rec(1, A9021, A20, 777.77);
    delete broken["Сумма"];
    broken["СуммаНУDr"] = 555.55;
    const { ctx } = setup({ records: [broken] });
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    expect(r.isError).toBe(true);
    expect(r.structuredContent).toBeUndefined();
    expect(text(r)).toMatch(/нет полей Сумма/);
    expect(text(r)).toContain("СуммаНУDr"); // фактические имена полей
    expect(text(r)).not.toContain("555.55");
    expect(text(r)).not.toContain("777.77");
  });

  it("null/нечисловая сумма — ошибка, а не 0", async () => {
    const { conn } = setup({ records: [rec(1, A9021, A20, "abc")] });
    await expect(getDocumentPostings(conn, DOC, DOC_REF)).rejects.toThrow(/Сумма.*не является числом/);
    const { conn: c2 } = setup({ records: [rec(1, A9021, A20, 1, { Сумма: null })] });
    await expect(getDocumentPostings(c2, DOC, DOC_REF)).rejects.toThrow(/нет полей Сумма/);
  });

  it("в $metadata записей регистра нет полей проводки — ошибка до запроса к регистру", async () => {
    const { conn, calls } = setup({ recordProps: ["Recorder", "Recorder_Type", "Amount"] });
    const err = await getDocumentPostings(conn, DOC, DOC_REF).catch((e: Error) => e);
    expect(err).toBeInstanceOf(Error);
    expect((err as Error).message).toMatch(/нет полей Period, AccountDr_Key, AccountCr_Key, Сумма/);
    expect((err as Error).message).toContain("RecordSet"); // имена полей кандидата-набора
    expect(regCalls(calls)).toHaveLength(0);
  });

  it("1С проигнорировала отбор и вернула чужие записи — ошибка, не смешанный набор", async () => {
    const { conn } = setup({
      ignoreRecorderFilter: true,
      records: [rec(1, A9021, A20, 1), { ...rec(2, A20, A26, 2), Recorder: OTHER_REF }],
    });
    await expect(getDocumentPostings(conn, DOC, DOC_REF)).rejects.toThrow(/другого регистратора/);
  });

  it("пагинация: 2350 записей при странице 1000 — все выбраны, без потерь и дублей", async () => {
    const records = Array.from({ length: 2350 }, (_, i) => rec(i + 1, A9021, A20, 1));
    const { conn, calls } = setup({ records });
    const res = await getDocumentPostings(conn, DOC, DOC_REF);
    expect(res.postingsCount).toBe(2350);
    expect(res.debitTotal).toBe(2350);
    const reg = regCalls(calls);
    expect(reg).toHaveLength(3);
    expect(reg.map((c) => /\$skip=(\d+)/.exec(c.path)?.[1])).toEqual(["0", "1000", "2000"]);
  });

  it("переполнение потолка — AggregateOverflowError, частичного набора нет", async () => {
    const records = Array.from({ length: 30 }, (_, i) => rec(i + 1, A9021, A20, 1));
    const { conn, ctx } = setup({ records, cap: 10 });
    await expect(getDocumentPostings(conn, DOC, DOC_REF)).rejects.toBeInstanceOf(AggregateOverflowError);
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    expect(r.isError).toBe(true);
    expect(r.structuredContent).toBeUndefined();
  });

  it("никаких пишущих HTTP-методов — только GET (даже при разрешённой записи)", async () => {
    const { ctx, calls } = setup({ records: SAMPLE });
    await callTool(ctx, { documentEntity: DOC, documentRef: DOC_REF });
    await callTool(ctx, { documentEntity: DOC, documentRef: OTHER_REF });
    expect(calls.length).toBeGreaterThan(0);
    expect(new Set(calls.map((c) => c.method))).toEqual(new Set(["GET"]));
    // и ни одного обращения к действиям документа (Post/Unpost и т.п.)
    expect(calls.some((c) => /\/(Post|Unpost)\b/.test(c.path))).toBe(false);
  });

  it("вид документа: справочник, произвольная сущность и табличная часть отклоняются", async () => {
    const { ctx, conn, calls } = setup({ records: SAMPLE });
    for (const bad of ["Catalog_Контрагенты", REG, "Foo"]) {
      const r = await callTool(ctx, { documentEntity: bad, documentRef: DOC_REF });
      expect(r.isError).toBe(true);
      await expect(getDocumentPostings(conn, bad, DOC_REF)).rejects.toThrow(/не документ/);
    }
    expect(calls).toHaveLength(0); // отсечено до запросов
    await expect(getDocumentPostings(conn, `${DOC}_Список`, DOC_REF)).rejects.toThrow(/не документ/);
    // неопубликованный документ — подсказка про «Состав OData»
    const r = await callTool(ctx, { documentEntity: "Document_Несуществующий", documentRef: DOC_REF });
    expect(r.isError).toBe(true);
    expect(text(r)).toMatch(/не опубликованы в OData/);
  });

  it("конфигурация полей записей регистра собрана в одной константе", () => {
    expect(REGISTER_RECORDS.required).toEqual({
      period: "Period",
      accountDr: "AccountDr_Key",
      accountCr: "AccountCr_Key",
      amount: "Сумма",
    });
  });
});
