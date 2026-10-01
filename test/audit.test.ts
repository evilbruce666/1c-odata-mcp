import { describe, it, expect, vi, afterEach } from "vitest";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { Connection, type ServerContext } from "../src/context.js";
import { uuidTimestamp } from "../src/odata/uuid.js";
import {
  BASE_LIMITATIONS,
  REF_CREATED_AT_LIMITATIONS,
  REF_CREATED_AT_SOURCE,
  getDocumentHistory,
  registerAuditTools,
} from "../src/tools/audit.js";
import { getDocumentHistoryResultSchema } from "../src/schemas/output.js";

/**
 * Хронология документа: настоящий Connection/ODataClient, подменён только global fetch.
 *
 * ВАЖНО: фейковая 1С отвечает так, как мы ОЖИДАЕМ от реальной (структура взята из
 * живой пробы $metadata, но значения здесь синтетические). Живая проверка —
 * отдельно, эти тесты её не заменяют и живых фактов не утверждают.
 */

const BASE = "http://1c.test/base/odata/standard.odata/";
const DOC = "Document_РегламентнаяОперация";
const DOC_NO_RESP = "Document_БезОтветственного";
const V1_REF = "919a75d1-7f6a-11f1-86c3-74563c4bf0d1";
const V1_ISO = "2026-07-14T09:58:31.127Z"; // независимо посчитано (python uuid.UUID(...).time)
const V4_REF = "3f0e4c9a-8b1d-4e2f-9a6b-1c2d3e4f5a6b";
const ORG = "e726d309-13a0-11e3-ae87-e8039ae81ce9";
const USER = "d6953e8a-50b6-11e2-9ec4-c86000df0d7b";
const EMPTY = "00000000-0000-0000-0000-000000000000";
const REG = "AccountingRegister_Хозрасчетный";
const RECS = `${REG}_RecordType`;
const ORGS = "Catalog_Организации";
const USERS = "Catalog_Пользователи";
const DOC_DATE = "2026-06-30T23:59:59";

const REC_PROPS = [
  "Recorder",
  "Recorder_Type",
  "Period",
  "LineNumber",
  "AccountDr_Key",
  "AccountCr_Key",
  "Сумма",
];
const DOC_PROPS = [
  "Ref_Key",
  "DataVersion",
  "Number",
  "Date",
  "Posted",
  "DeletionMark",
  "ВидОперации",
  "Организация_Key",
  "Состояние",
  "Ответственный_Key",
  "Комментарий",
];

type Row = Record<string, unknown>;

interface Fake1C {
  docRef?: string;
  docExists?: boolean;
  doc?: Row;
  /** "records" — есть движения; "none" — нет; "unpublished" — регистр не опубликован; "400"/"500" — ошибка. */
  register?: "records" | "none" | "unpublished" | "400" | "500" | "ignoreFilter" | "ignoreTop";
  usersPublished?: boolean;
  user?: "ok" | "404" | "401" | "500" | "noDescription";
}

interface Call {
  method: string;
  path: string;
  url: URL;
}

function metadataXml(o: Fake1C): string {
  const prop = (n: string): string => `<Property Name="${n}" Type="Edm.String"/>`;
  const type = (name: string, props: string[], keys = ["Ref_Key"]): string =>
    `<EntityType Name="${name}"><Key>${keys.map((k) => `<PropertyRef Name="${k}"/>`).join("")}</Key>${props
      .map(prop)
      .join("")}</EntityType>`;
  const types = [
    type(DOC, DOC_PROPS),
    type(`${DOC}_УдалитьОшибки`, ["Ref_Key", "LineNumber", "Описание"], ["Ref_Key", "LineNumber"]),
    type(DOC_NO_RESP, ["Ref_Key", "Number", "Date"]),
    type(ORGS, ["Ref_Key", "Description"]),
    type("Catalog_Контрагенты", ["Ref_Key", "Description"]),
  ];
  const sets = [DOC, `${DOC}_УдалитьОшибки`, DOC_NO_RESP, ORGS, "Catalog_Контрагенты"];
  if (o.usersPublished !== false) {
    types.push(type(USERS, ["Ref_Key", "Description"]));
    sets.push(USERS);
  }
  if (o.register !== "unpublished") {
    types.push(type(REG, ["Recorder", "Recorder_Type", "RecordSet"], ["Recorder", "Recorder_Type"]));
    types.push(type(RECS, REC_PROPS, ["Recorder", "Recorder_Type", "LineNumber"]));
    sets.push(REG, RECS);
  }
  return `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx xmlns:edmx="http://schemas.microsoft.com/ado/2007/06/edmx" Version="1.0">
<edmx:DataServices m:DataServiceVersion="3.0" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">
<Schema Namespace="StandardODATA" xmlns="http://schemas.microsoft.com/ado/2009/11/edm">
${types.join("\n")}
<EntityContainer Name="EnterpriseV8" m:IsDefaultEntityContainer="true">
${sets.map((s) => `<EntitySet Name="${s}" EntityType="StandardODATA.${s}"/>`).join("\n")}
</EntityContainer>
</Schema>
</edmx:DataServices>
</edmx:Edmx>`;
}

function setup(o: Fake1C = {}) {
  const calls: Call[] = [];
  const docRef = o.docRef ?? V1_REF;
  const regMode = o.register ?? "records";
  const json = (body: unknown, status = 200): Response =>
    new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
  const err = (status: number): Response =>
    json({ "odata.error": { code: "", message: { lang: "ru", value: `ошибка ${status}` } } }, status);
  // Много записей регистратора — чтобы убедиться, что инструмент берёт не больше одной.
  const records: Row[] = Array.from({ length: 52 }, (_, i) => ({
    Recorder: docRef,
    Recorder_Type: `StandardODATA.${DOC}`,
    Period: DOC_DATE,
    LineNumber: i + 1,
  }));

  vi.stubGlobal("fetch", async (input: string | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    const method = init?.method ?? "GET";
    const rel = decodeURIComponent(url.pathname.slice(new URL(BASE).pathname.length));
    calls.push({ method, path: decodeURIComponent(url.pathname + url.search), url });
    const q = (k: string): string | undefined => url.searchParams.get(k) ?? undefined;
    const filter = q("$filter") ?? "";
    const top = Number(q("$top") ?? 1e9);

    if (rel === "$metadata") return new Response(metadataXml(o), { status: 200 });

    if (rel.startsWith(`${DOC}(`) || rel.startsWith(`${DOC_NO_RESP}(`)) {
      if (o.docExists === false) return err(404);
      return json({
        Ref_Key: docRef,
        DataVersion: "AAAAAACP4L8=",
        Number: "0000-000061",
        Date: DOC_DATE,
        Posted: false,
        DeletionMark: false,
        ВидОперации: "ЗакрытиеСчетов20_23_25_26",
        Организация_Key: ORG,
        Состояние: "Выполнено",
        Ответственный_Key: USER,
        ...o.doc,
      });
    }

    if (rel === ORGS) {
      return json({ value: filter.includes(ORG) ? [{ Ref_Key: ORG, Description: "ООО Тест" }] : [] });
    }

    if (rel.startsWith(`${USERS}(`)) {
      const mode = o.user ?? "ok";
      if (mode === "404") return err(404);
      if (mode === "401") return err(401);
      if (mode === "500") return err(500);
      if (mode === "noDescription") return json({ Ref_Key: USER });
      return json({ Ref_Key: USER, Description: "Бухгалтер Тестовый" });
    }

    if (rel === RECS) {
      if (regMode === "400") return err(400);
      if (regMode === "500") return err(500);
      if (regMode === "none") return json({ value: [] });
      if (regMode === "ignoreFilter") {
        return json({ value: [{ Recorder: V4_REF, Recorder_Type: `StandardODATA.${DOC}` }] });
      }
      if (regMode === "ignoreTop") return json({ value: records.slice(0, 5) });
      const m = /^Recorder eq cast\(guid'([^']+)', '([^']+)'\)$/.exec(filter);
      if (!m) return err(400);
      const pool = records.filter(
        (r) => r["Recorder"] === m[1] && String(r["Recorder_Type"]).endsWith(m[2]!),
      );
      return json({ value: pool.slice(0, top) });
    }

    return err(404);
  });

  // readOnly=false и writable=true — чтобы пишущий запрос (если бы он был) дошёл
  // до fetch и попал в calls, а не был отсечён гардом клиента.
  const conn = new Connection(
    { name: "default", baseUrl: BASE, username: "u", password: "p", writable: true },
    {
      timeoutMs: 5_000,
      retries: 0,
      pageSize: 100,
      maxRows: 1000,
      analyticsMaxRows: 200_000,
      readOnly: false,
    },
  );
  const ctx = { db: () => conn } as unknown as ServerContext;
  return { conn, ctx, calls };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

async function callTool(ctx: ServerContext, args: Record<string, unknown>) {
  const server = new McpServer({ name: "t", version: "0" });
  registerAuditTools(server, ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "c", version: "0" });
  await client.connect(b);
  try {
    return await client.callTool({ name: "read.audit.get_document_history", arguments: args });
  } finally {
    await client.close();
  }
}
const text = (r: Awaited<ReturnType<typeof callTool>>): string =>
  (r.content as Array<{ text: string }>)[0]!.text;

type History = Awaited<ReturnType<typeof getDocumentHistory>>;
const run = (o: Fake1C = {}, entity = DOC, ref = o.docRef ?? V1_REF) => {
  const s = setup(o);
  return { ...s, result: getDocumentHistory(s.conn, entity, ref) };
};

// ─── UUID ─────────────────────────────────────────────────────────────────────

describe("uuidTimestamp", () => {
  it("UUIDv1: известный GUID → ожидаемая метка UTC, node, clock sequence", () => {
    const r = uuidTimestamp(V1_REF);
    expect(r).toEqual({
      kind: "v1",
      version: 1,
      variant: "RFC4122",
      timestampIso: V1_ISO,
      timestamp100ns: "140033159111275985",
      clockSequence: 1731,
      node: "74563c4bf0d1",
      nodeMulticastBit: false,
    });
  });

  it("второй известный UUIDv1 (другой месяц) декодируется независимо", () => {
    const r = uuidTimestamp("525f9ffb-1eaa-11f1-86b8-74563c4bf0d1");
    expect(r.kind === "v1" && r.timestampIso).toBe("2026-03-13T07:00:30.046Z");
  });

  it("регистр букв и фигурные скобки не влияют на результат", () => {
    expect(uuidTimestamp(`{${V1_REF.toUpperCase()}}`)).toEqual(uuidTimestamp(V1_REF));
  });

  it("UUIDv4 — метки времени нет", () => {
    const r = uuidTimestamp(V4_REF);
    expect(r.kind).toBe("no-timestamp");
    expect(r.version).toBe(4);
    expect(r).not.toHaveProperty("timestampIso");
  });

  it("другие версии (3, 5, 6, 7) — метки времени нет", () => {
    for (const v of ["3", "5", "6", "7"]) {
      const g = `919a75d1-7f6a-${v}1f1-86c3-74563c4bf0d1`;
      expect(uuidTimestamp(g)).toMatchObject({ kind: "no-timestamp", version: Number(v) });
    }
  });

  it("версия 1, но вариант не RFC 4122 (NCS/Microsoft/Reserved) — метки времени нет", () => {
    expect(uuidTimestamp("919a75d1-7f6a-11f1-06c3-74563c4bf0d1")).toMatchObject({
      kind: "no-timestamp",
      variant: "NCS",
    });
    expect(uuidTimestamp("919a75d1-7f6a-11f1-c6c3-74563c4bf0d1")).toMatchObject({
      kind: "no-timestamp",
      variant: "Microsoft",
    });
    expect(uuidTimestamp("919a75d1-7f6a-11f1-e6c3-74563c4bf0d1")).toMatchObject({
      kind: "no-timestamp",
      variant: "Reserved",
    });
  });

  it("некорректная строка — исключение, а не «нет метки»", () => {
    for (const bad of [
      "",
      "not-a-guid",
      "919a75d1-7f6a-11f1-86c3-74563c4bf0d",
      "919a75d1x7f6a-11f1-86c3-74563c4bf0d1",
    ]) {
      expect(() => uuidTimestamp(bad)).toThrow(/GUID/);
    }
  });

  it("бит multicast в node отражается как есть (случайный node по RFC 4122)", () => {
    const r = uuidTimestamp("919a75d1-7f6a-11f1-86c3-75563c4bf0d1");
    expect(r.kind === "v1" && r.nodeMulticastBit).toBe(true);
  });
});

// ─── Инструмент ──────────────────────────────────────────────────────────────

describe("read.audit.get_document_history", () => {
  it("корректный документ: шапка, documentDate и refCreatedAt раздельно, ответственный, движения", async () => {
    const { ctx, calls } = setup();
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: V1_REF });
    expect(r.isError).toBeFalsy();
    const data = r.structuredContent as History;
    // SDK уже провалидировал structuredContent по outputSchema; проверим и явно.
    expect(getDocumentHistoryResultSchema.safeParse(data).success).toBe(true);

    expect(data.document).toEqual({
      entitySet: DOC,
      ref: V1_REF,
      number: "0000-000061",
      posted: false,
      deletionMark: false,
      organization: "ООО Тест",
      organizationRef: ORG,
      operation: "ЗакрытиеСчетов20_23_25_26",
      state: "Выполнено",
      dataVersion: "AAAAAACP4L8=",
    });
    expect(data.timestamps).toEqual({
      documentDate: DOC_DATE,
      refCreatedAt: {
        value: V1_ISO,
        source: REF_CREATED_AT_SOURCE,
        confidence: "derived",
        description: expect.stringContaining("Ref_Key creation timestamp derived from UUIDv1"),
      },
    });
    expect(data.responsible).toMatchObject({ ref: USER, name: "Бухгалтер Тестовый", resolution: "resolved" });
    expect(data.accountingMovements).toMatchObject({ status: "exists", exists: true, source: RECS });
    expect(data.limitations).toEqual([...BASE_LIMITATIONS, ...REF_CREATED_AT_LIMITATIONS]);
    expect(calls.every((c) => c.method === "GET")).toBe(true);
  });

  it("несколько свидетельств: Date и Ref_Key — отдельные записи evidence с разными полями-источниками", async () => {
    const { result } = run();
    const { evidence } = await result;
    expect(evidence.map((e) => e.sourceField)).toEqual(["Date", "Ref_Key"]);
    const [date, key] = evidence;
    expect(date).toMatchObject({ sourceEntity: DOC, sourceRef: V1_REF, timestamp: DOC_DATE });
    expect(date!.description).toMatch(/учётная дата/);
    expect(key).toMatchObject({
      sourceEntity: DOC,
      sourceRef: V1_REF,
      timestamp: V1_ISO,
      details: {
        uuidVersion: 1,
        uuidVariant: "RFC4122",
        node: "74563c4bf0d1",
        timestamp100ns: "140033159111275985",
      },
    });
    expect(key!.description).toMatch(/не доказанный идентификатор сервера/);
  });

  it("происхождение метки: refCreatedAt всегда с источником Ref_Key UUIDv1 и confidence=derived", async () => {
    const { timestamps } = await run().result;
    expect(timestamps.refCreatedAt).toMatchObject({
      source: "Ref_Key UUIDv1 timestamp",
      confidence: "derived",
    });
    expect(timestamps.refCreatedAt!.description).toMatch(
      /НЕ проверенное время создания, выполнения или проведения/,
    );
  });

  it("executedAt и modifiedAt никогда не возвращаются (через OData недоступны)", async () => {
    for (const o of [{}, { docRef: V4_REF }, { register: "none" as const }]) {
      const { timestamps, limitations } = await run(o).result;
      expect(timestamps).not.toHaveProperty("executedAt");
      expect(timestamps).not.toHaveProperty("modifiedAt");
      expect(timestamps).not.toHaveProperty("createdAt");
      expect(limitations.some((l) => l.includes("executedAt"))).toBe(true);
      expect(limitations.some((l) => l.includes("modifiedAt"))).toBe(true);
    }
  });

  it("некорректный GUID: ошибка до любого запроса к 1С", async () => {
    const { ctx, calls } = setup();
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: "x' or 1 eq 1" });
    expect(r.isError).toBe(true);
    expect(calls).toHaveLength(0);
    const s = setup();
    await expect(getDocumentHistory(s.conn, DOC, "not-a-guid")).rejects.toThrow(/GUID/);
    expect(s.calls).toHaveLength(0);
  });

  it("некорректный вид документа: не Document_*, не опубликован, табличная часть", async () => {
    const { ctx, calls } = setup();
    const r = await callTool(ctx, { documentEntity: "Catalog_Пользователи", documentRef: V1_REF });
    expect(r.isError).toBe(true);
    expect(calls).toHaveLength(0);

    const s = setup();
    await expect(getDocumentHistory(s.conn, "Catalog_Пользователи", V1_REF)).rejects.toThrow(/не документ/);
    await expect(getDocumentHistory(s.conn, "Document_НетТакого", V1_REF)).rejects.toThrow(/не опубликованы/);
    await expect(getDocumentHistory(s.conn, `${DOC}_УдалитьОшибки`, V1_REF)).rejects.toThrow(
      /табличная часть/,
    );
    // Ни одного запроса за пределами $metadata.
    expect(s.calls.every((c) => c.path.endsWith("$metadata"))).toBe(true);
  });

  it("документ не найден: явная ошибка not_found, без проверки регистра и ответственного", async () => {
    const { ctx, calls } = setup({ docExists: false });
    const r = await callTool(ctx, { documentEntity: DOC, documentRef: V1_REF });
    expect(r.isError).toBe(true);
    expect(text(r)).toMatch(/\[not_found\].*не найден/);
    expect(calls.some((c) => c.path.includes(REG) || c.path.includes(USERS))).toBe(false);
  });

  it("1С вернула документ без Ref_Key или с чужим Ref_Key — громкая ошибка", async () => {
    await expect(run({ doc: { Ref_Key: undefined } }).result).rejects.toThrow(/Ref_Key/);
    await expect(run({ doc: { Ref_Key: V4_REF } }, DOC, V1_REF).result).rejects.toThrow(/другим Ref_Key/);
  });

  it("источник хронологии найден (UUIDv1) → refCreatedAt есть", async () => {
    const { timestamps } = await run().result;
    expect(timestamps.refCreatedAt?.value).toBe(V1_ISO);
  });

  it("источника хронологии нет (UUIDv4) → refCreatedAt отсутствует, причина — в limitations и evidence", async () => {
    const { timestamps, evidence, limitations, document } = await run({ docRef: V4_REF }).result;
    expect(timestamps).toEqual({ documentDate: DOC_DATE });
    expect(document.ref).toBe(V4_REF);
    const key = evidence.find((e) => e.sourceField === "Ref_Key")!;
    expect(key).not.toHaveProperty("timestamp");
    expect(key.details).toEqual({ uuidVersion: 4, uuidVariant: "RFC4122" });
    expect(
      limitations.some((l) => l.startsWith("refCreatedAt не возвращается") && l.includes("версии 4")),
    ).toBe(true);
  });

  it("UUIDv4: в limitations нет утверждений, что refCreatedAt был выведен", async () => {
    const { limitations } = await run({ docRef: V4_REF }).result;
    for (const l of REF_CREATED_AT_LIMITATIONS) expect(limitations).not.toContain(l);
    const aboutRef = limitations.filter((l) => l.includes("refCreatedAt"));
    expect(aboutRef).toHaveLength(1);
    expect(aboutRef[0]).toMatch(/^refCreatedAt не возвращается: /);
    expect(limitations.some((l) => /выведено из UUIDv1|derived|отражает генерацию/.test(l))).toBe(false);
    // Общие ограничения при этом на месте.
    for (const l of BASE_LIMITATIONS) expect(limitations).toContain(l);
  });

  it("BASE_LIMITATIONS не упоминают refCreatedAt (верны для любой версии UUID)", () => {
    for (const l of BASE_LIMITATIONS) expect(l).not.toMatch(/refCreatedAt|UUID|Ref_Key/);
  });

  it("метка времени не сравнивается с часами MCP-сервера: результат детерминирован", async () => {
    // UUIDv1 с меткой далеко в будущем (≈ 5236 г.) — никаких суждений о часах.
    const future = "ffffffff-ffff-1fff-8000-000000000001";
    const first = await run({ docRef: future }).result;
    vi.useFakeTimers();
    try {
      vi.setSystemTime(new Date("6000-01-01T00:00:00Z"));
      const later = await run({ docRef: future }).result;
      expect(later.timestamps).toEqual(first.timestamps);
      expect(later.limitations).toEqual(first.limitations);
    } finally {
      vi.useRealTimers();
    }
    expect(first.timestamps.refCreatedAt?.value).toMatch(/^5236-/);
    expect(first.limitations).toEqual([...BASE_LIMITATIONS, ...REF_CREATED_AT_LIMITATIONS]);
  });

  it("Document.Date никогда не попадает в refCreatedAt / executedAt", async () => {
    // v4: единственная дата в ответе — documentDate.
    const v4 = await run({ docRef: V4_REF, doc: { Date: "2026-01-31T23:59:59" } }).result;
    expect(v4.timestamps).toEqual({ documentDate: "2026-01-31T23:59:59" });
    expect(JSON.stringify(v4.timestamps).match(/2026-01-31/g)).toHaveLength(1);
    // v1: refCreatedAt не зависит от Date.
    const a = await run({ doc: { Date: "2020-01-01T00:00:00" } }).result;
    const b = await run({ doc: { Date: "2030-12-31T23:59:59" } }).result;
    expect(a.timestamps.refCreatedAt).toEqual(b.timestamps.refCreatedAt);
    expect(a.timestamps.refCreatedAt!.value).toBe(V1_ISO);
    // Без Date — documentDate нет и ничем не подменяется.
    const noDate = await run({ docRef: V4_REF, doc: { Date: "" } }).result;
    expect(noDate.timestamps).toEqual({});
    expect(noDate.limitations.some((l) => l.includes("нет значения Date"))).toBe(true);
  });

  it("DataVersion возвращается без изменений и не влияет ни на одну метку времени", async () => {
    const a = await run({ doc: { DataVersion: "AAAAAACP4L8=" } }).result;
    const b = await run({ doc: { DataVersion: "//////////8=" } }).result;
    expect(a.document.dataVersion).toBe("AAAAAACP4L8=");
    expect(b.document.dataVersion).toBe("//////////8=");
    expect(a.timestamps).toEqual(b.timestamps);
    expect(a.evidence).toEqual(b.evidence);
    expect(JSON.stringify(a.timestamps)).not.toContain("AAAAAACP4L8=");
  });

  describe("движения по регистру бухгалтерии", () => {
    it("есть движения → status=exists, exists=true", async () => {
      const { accountingMovements, scan } = await run().result;
      expect(accountingMovements).toEqual({
        status: "exists",
        exists: true,
        source: RECS,
        filter: `Recorder eq cast(guid'${V1_REF}', '${DOC}')`,
      });
      expect(scan.rowsScanned).toBe(1);
    });

    it("движений нет → status=none, exists=false", async () => {
      const { accountingMovements } = await run({ register: "none" }).result;
      expect(accountingMovements).toMatchObject({ status: "none", exists: false, source: RECS });
    });

    it("регистр не опубликован → unsupported, а не exists=false", async () => {
      const { accountingMovements, limitations } = await run({ register: "unpublished" }).result;
      expect(accountingMovements.status).toBe("unsupported");
      expect(accountingMovements).not.toHaveProperty("exists");
      expect(limitations.some((l) => l.includes("не означает, что движений нет"))).toBe(true);
    });

    it("ошибка запроса (400/500) → error, а не exists=false; инструмент не падает", async () => {
      for (const mode of ["400", "500"] as const) {
        const { accountingMovements } = await run({ register: mode }).result;
        expect(accountingMovements.status).toBe("error");
        expect(accountingMovements).not.toHaveProperty("exists");
        expect(accountingMovements.detail).toMatch(/\[(bad_request|server)\]/);
      }
    });

    it("1С проигнорировала отбор по регистратору или $top → error", async () => {
      expect((await run({ register: "ignoreFilter" }).result).accountingMovements.status).toBe("error");
      const top = (await run({ register: "ignoreTop" }).result).accountingMovements;
      expect(top.status).toBe("error");
      expect(top.detail).toMatch(/\$top/);
    });

    it("проверка движений не влияет на метки времени", async () => {
      const a = await run({ register: "records" }).result;
      const b = await run({ register: "none" }).result;
      const c = await run({ register: "500" }).result;
      expect(a.timestamps).toEqual(b.timestamps);
      expect(a.timestamps).toEqual(c.timestamps);
      expect(a.evidence).toEqual(b.evidence);
    });
  });

  describe("ответственный", () => {
    it("Ответственный_Key разрешён через Catalog_Пользователи", async () => {
      const { responsible } = await run().result;
      expect(responsible).toEqual({
        ref: USER,
        name: "Бухгалтер Тестовый",
        resolution: "resolved",
        source: `Document.Ответственный_Key → ${USERS}.Description`,
      });
    });

    it("Ответственный_Key пуст → responsible отсутствует, причина в limitations", async () => {
      const { responsible, limitations } = await run({ doc: { Ответственный_Key: EMPTY } }).result;
      expect(responsible).toBeUndefined();
      expect(limitations.some((l) => l.includes("не заполнен"))).toBe(true);
    });

    it("у вида документа нет реквизита Ответственный_Key → responsible отсутствует", async () => {
      const { responsible, limitations, calls } = await (async () => {
        const s = setup({ docRef: V4_REF });
        const r = await getDocumentHistory(s.conn, DOC_NO_RESP, V4_REF);
        return { ...r, calls: s.calls };
      })();
      expect(responsible).toBeUndefined();
      expect(limitations.some((l) => l.includes("нет реквизита Ответственный_Key"))).toBe(true);
      expect(calls.some((c) => c.path.includes(USERS))).toBe(false);
    });

    it("сбой поиска имени не роняет инструмент и не превращается в «неизвестного автора»", async () => {
      for (const [mode, resolution] of [
        ["500", "lookup_failed"],
        ["401", "lookup_failed"],
        ["noDescription", "lookup_failed"],
        ["404", "not_found"],
      ] as const) {
        const { responsible, limitations, timestamps } = await run({ user: mode }).result;
        expect(responsible).toMatchObject({ ref: USER, resolution });
        expect(responsible).not.toHaveProperty("name");
        // Только ref + статус поиска: ни имени-заглушки, ни ролей автора/исполнителя.
        expect(Object.keys(responsible!).sort()).toEqual(["detail", "ref", "resolution", "source"]);
        expect(responsible!.detail).not.toMatch(/unknown|неизвестн/i);
        expect(limitations.some((l) => l.startsWith("Имя ответственного не получено"))).toBe(true);
        expect(timestamps.refCreatedAt?.value).toBe(V1_ISO);
      }
    });

    it("справочник пользователей не опубликован → catalog_not_published, ref сохранён", async () => {
      const { responsible } = await run({ usersPublished: false }).result;
      expect(responsible).toMatchObject({ ref: USER, resolution: "catalog_not_published" });
    });
  });

  describe("безопасность и ограниченность запросов", () => {
    it("только GET — даже когда клиенту разрешена запись", async () => {
      for (const o of [{}, { register: "none" as const }, { user: "500" as const }, { docRef: V4_REF }]) {
        const s = setup(o);
        await getDocumentHistory(s.conn, DOC, o.docRef ?? V1_REF);
        expect(s.calls.length).toBeGreaterThan(0);
        expect(s.calls.every((c) => c.method === "GET")).toBe(true);
      }
    });

    it("регистр: ровно один запрос, с серверным $filter по регистратору и $top=1, без $skip-листания", async () => {
      const s = setup(); // у регистратора 52 записи
      await getDocumentHistory(s.conn, DOC, V1_REF);
      const reg = s.calls.filter((c) => c.path.includes(`/${REG}`));
      expect(reg).toHaveLength(1);
      const u = reg[0]!.url;
      expect(u.searchParams.get("$filter")).toBe(`Recorder eq cast(guid'${V1_REF}', '${DOC}')`);
      expect(u.searchParams.get("$top")).toBe("1");
      expect(u.searchParams.has("$skip")).toBe(false);
    });

    it("нет полной выборки: каждый запрос — $metadata, GET по ключу или с $filter и $top", async () => {
      const s = setup();
      await getDocumentHistory(s.conn, DOC, V1_REF);
      for (const c of s.calls) {
        const rel = decodeURIComponent(c.url.pathname.slice(new URL(BASE).pathname.length));
        const byKey = /\(guid'[0-9a-f-]{36}'\)$/.test(rel);
        const bounded = c.url.searchParams.has("$filter") && c.url.searchParams.has("$top");
        expect(rel === "$metadata" || byKey || bounded, c.path).toBe(true);
      }
    });
  });
});
