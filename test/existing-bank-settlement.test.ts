import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Connection } from "../src/context.js";
import { parseConfig } from "../src/config/env.js";
import { createServer } from "../src/mcp/server.js";
import { ODataError } from "../src/odata/errors.js";
import { paymentAmountWarnings } from "../src/odata/existing-bank-settlement.js";
import { bp30Metadata, bp30Enums } from "./helpers/metadata-snapshot.js";
import { settlementAccountsFor } from "../src/odata/settlement-accounts.js";
import type { EntityMeta, ODataEntity } from "../src/types/odata.js";
import type { CallToolResult } from "@modelcontextprotocol/server";

const id = (n: number) => `${n}`.repeat(8) + "-1111-4111-8111-111111111111";
const ORG = id(1),
  CP = id(2),
  CTR = id(3),
  ACCOUNT = id(4),
  GROUP = id(5),
  CUR = id(6),
  DOC = id(7),
  OLD = id(8);
const EMPTY = "00000000-0000-0000-0000-000000000000";
const SET = "Document_ПоступлениеНаРасчетныйСчет",
  REG = "InformationRegister_СчетаУчетаРасчетовСКонтрагентами";
const CONTRACT = "ДоговорКонтрагента_Key",
  ACCOUNT_FIELD = "СчетУчетаРасчетовСКонтрагентом_Key",
  ROWS = "РасшифровкаПлатежа";
const TOOL = "write.document.apply_bank_settlement_settings";
const dirs: string[] = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(dirs.splice(0).map((d) => rm(d, { recursive: true, force: true })));
});
const data = (r: CallToolResult) => r.structuredContent as Record<string, unknown>;
async function fixture(entitySet = SET) {
  const dir = await mkdtemp(join(tmpdir(), "bank-settings-"));
  dirs.push(dir);
  const cfg = parseConfig({
    ODATA_BASE_URL: "https://example.invalid/odata/",
    ODATA_USERNAME: "synthetic",
    ODATA_PASSWORD: "synthetic",
    READ_ONLY: "false",
    ODATA_WRITABLE: "true",
    WRITE_JOURNAL_DIR: dir,
  });
  const conn = new Connection(
    { ...cfg.connections[0]!, writable: true },
    { ...cfg.behavior, writeJournalDir: dir, regulatedCurrencyCode: "933" },
  );
  let doc: ODataEntity = {
    Ref_Key: DOC,
    DataVersion: "v1",
    Posted: false,
    DeletionMark: false,
    Организация_Key: ORG,
    Контрагент_Key: CP,
    ВалютаДокумента_Key: CUR,
    ВидОперации: "РасчетыПоКредитамИЗаймам",
    [CONTRACT]: OLD,
    [ACCOUNT_FIELD]: OLD,
    СуммаДокумента: 150,
    Date: "2010-03-22T00:00:00",
    НазначениеПлатежа: "Synthetic loan",
    СубконтоДт1: "cost-analytic",
    СубконтоДт1_Type: "StandardODATA.Catalog_СтатьиЗатрат",
    [ROWS]: [
      {
        LineNumber: "1",
        [CONTRACT]: OLD,
        [ACCOUNT_FIELD]: OLD,
        СчетУчетаРасчетовПоАвансам_Key: OLD,
        СуммаПлатежа: 150,
        СуммаНДС: 0,
        СтавкаНДС: "НДС20",
        СтатьяДвиженияДенежныхСредств_Key: OLD,
        Субконто: "preserved",
      },
    ],
  };
  let contract: ODataEntity = {
    Ref_Key: CTR,
    DeletionMark: false,
    IsFolder: false,
    Owner_Key: CP,
    Организация_Key: ORG,
    ВалютаВзаиморасчетов_Key: CUR,
    ВидДоговора: "Прочее",
    РасчетыВУсловныхЕдиницах: false,
  };
  let counterparty: ODataEntity | Error = {
    Ref_Key: CP,
    IsFolder: false,
    DeletionMark: false,
    Parent_Key: GROUP,
    Description: "Synthetic counterparty",
  };
  let account: ODataEntity = {
    Ref_Key: ACCOUNT,
    Code: "67.2",
    Description: "Synthetic account",
    DeletionMark: false,
    IsFolder: false,
    ЗапретитьИспользоватьВПроводках: false,
  };
  let settings: ODataEntity[] = [
    {
      Организация_Key: EMPTY,
      Контрагент_Key: GROUP,
      Договор_Key: EMPTY,
      ВидРасчетовПоДоговору: "РасчетыВВалютеРегламентированногоУчета",
      СчетУчетаРасчетовСПоставщиком_Key: ACCOUNT,
      СчетУчетаАвансовВыданных_Key: ACCOUNT,
    },
  ];
  const meta = (
    name: string,
    properties: Record<string, string>,
    cls: EntityMeta["class"] = "catalog",
  ): EntityMeta => ({
    entitySet: name,
    entityType: name,
    class: cls,
    shortName: name,
    keys: ["Ref_Key"],
    properties: Object.entries(properties).map(([name, type]) => ({ name, type, nullable: true })),
    navigations: [],
  });
  const infer = (r: ODataEntity) =>
    Object.fromEntries(
      Object.entries(r).map(([k, v]) => [
        k,
        k.endsWith("_Key")
          ? "Edm.Guid"
          : typeof v === "boolean"
            ? "Edm.Boolean"
            : typeof v === "number"
              ? "Edm.Double"
              : "Edm.String",
      ]),
    );
  const entities = new Map<string, EntityMeta>([
    [
      entitySet,
      meta(
        entitySet,
        { ...infer(doc), [ROWS]: `Collection(StandardODATA.${entitySet}_${ROWS}_RowType)` },
        "document",
      ),
    ],
    [`${entitySet}_${ROWS}`, meta(`${entitySet}_${ROWS}`, infer((doc[ROWS] as ODataEntity[])[0]!))],
    [
      "Catalog_Контрагенты",
      meta("Catalog_Контрагенты", {
        Ref_Key: "Edm.Guid",
        Parent_Key: "Edm.Guid",
        DeletionMark: "Edm.Boolean",
        IsFolder: "Edm.Boolean",
      }),
    ],
    ["Catalog_ДоговорыКонтрагентов", meta("Catalog_ДоговорыКонтрагентов", infer(contract))],
    ["Catalog_Валюты", meta("Catalog_Валюты", { Ref_Key: "Edm.Guid", Code: "Edm.String" })],
    ["ChartOfAccounts_Хозрасчетный", meta("ChartOfAccounts_Хозрасчетный", infer(account), "chartOfAccounts")],
    [REG, meta(REG, infer(settings[0]!), "informationRegister")],
  ]);
  vi.spyOn(conn, "getMetadata").mockResolvedValue({ entities, odataVersion: "3.0" });
  let failure: "write" | "readback" | undefined;
  let readbackChange: object = {};
  let currency: ODataEntity | Error = { Code: "933" };
  const contractReads: string[] = [];
  const writes: Array<{ path: string; method: string; body: unknown; headers: unknown }> = [];
  vi.spyOn(conn.client, "request").mockImplementation(async (path, method = "GET", body, headers) => {
    if (method !== "GET") {
      writes.push({ path, method, body, headers });
      if (failure === "write") throw new ODataError({ kind: "timeout", message: "Synthetic timeout" });
      if (path.includes("/Unpost")) doc = { ...doc, Posted: false, DataVersion: "v2" };
      else if (path.includes("/Post")) doc = { ...doc, Posted: true, DataVersion: "v2" };
      else doc = { ...doc, ...(body as object), ...readbackChange, DataVersion: "v2" };
      return {} as never;
    }
    if (path.startsWith(entitySet)) {
      if (writes.length && failure === "readback")
        throw new ODataError({ kind: "auth", message: "Synthetic readback failure" });
      return structuredClone(doc) as never;
    }
    if (path.startsWith("Catalog_Договоры")) {
      contractReads.push(path);
      return contract as never;
    }
    if (path.startsWith("Catalog_Валюты")) {
      if (currency instanceof Error) throw currency;
      return currency as never;
    }
    if (path.startsWith("Catalog_Контрагенты")) {
      if (path.includes(GROUP))
        return { Ref_Key: GROUP, IsFolder: true, DeletionMark: false, Parent_Key: EMPTY } as never;
      if (counterparty instanceof Error) throw counterparty;
      return structuredClone(counterparty) as never;
    }
    if (path.startsWith(REG)) return { value: settings } as never;
    if (path.startsWith("ChartOfAccounts_"))
      return (path.includes("guid'") && !path.includes("$filter") ? account : { value: [account] }) as never;
    throw new Error(`Unexpected synthetic path ${path}`);
  });
  const tools = (
    createServer({ db: () => conn } as never) as unknown as {
      _registeredTools: Record<string, { handler: (args: object, extra: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  const run = async (extra: object = {}, tool = TOOL) => {
    expect(tools[tool], "bounded existing-bank tool is registered").toBeDefined();
    return tools[tool]!.handler(
      {
        database: "default",
        entitySet,
        ref: DOC,
        contractRef: CTR,
        side: "supplier",
        expectedDataVersion: "v1",
        confirm: false,
        ...extra,
      },
      {},
    );
  };
  return {
    run,
    conn,
    writes,
    contractReads,
    entities,
    doc: () => structuredClone(doc),
    setDoc: (v: object) => (doc = { ...doc, ...v }),
    setContract: (v: object) => (contract = { ...contract, ...v }),
    setCounterparty: (v: ODataEntity | Error) =>
      (counterparty = v instanceof Error ? v : { ...counterparty, ...v }),
    setAccount: (v: object) => (account = { ...account, ...v }),
    setSettings: (v: ODataEntity[]) => (settings = v),
    setCurrency: (v: typeof currency) => (currency = v),
    settings: () => settings,
    changeOnWrite: (v: object) => (readbackChange = v),
    fail: (v: typeof failure) => (failure = v),
  };
}

describe("existing bank inherited settlement", () => {
  it("applies all-org group supplier account with provenance while preserving header and row analytics", async () => {
    const f = await fixture(),
      before = f.doc(),
      p = await f.run();
    expect(p.isError).toBeFalsy();
    expect(data(p).resolution).toMatchObject({
      choice: { settle: ACCOUNT, provenance: { settle: { level: "группа", counterpartyRef: GROUP } } },
    });
    const r = await f.run({ confirm: true, operationId: data(p).operationId });
    expect(r.isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
    expect(f.writes[0]).toMatchObject({ method: "PATCH", headers: { "If-Match": "v1" } });
    expect(f.doc()).toEqual({
      ...before,
      DataVersion: "v2",
      [CONTRACT]: CTR,
      [ACCOUNT_FIELD]: ACCOUNT,
      [ROWS]: [
        {
          ...(before[ROWS] as object[])[0],
          [CONTRACT]: CTR,
          [ACCOUNT_FIELD]: ACCOUNT,
          СчетУчетаРасчетовПоАвансам_Key: ACCOUNT,
        },
      ],
    });
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
  });
  it("предупреждает до проведения, если сумма расшифровки не равна сумме документа (#61)", async () => {
    const f = await fixture();
    expect(data(await f.run()).warnings).toBeUndefined();
    const row = (f.doc()[ROWS] as ODataEntity[])[0]!;
    f.setDoc({ [ROWS]: [{ ...row, СуммаПлатежа: 0 }] });
    const p = await f.run();
    expect(p.isError).toBeFalsy();
    expect(String(data(p).warnings)).toContain(
      "сумма документа 150 не совпадает с суммой расшифровки платежа 0",
    );
    const r = await f.run({ confirm: true, operationId: data(p).operationId });
    expect(r.isError).toBeFalsy();
    expect(String(data(r).warnings)).toContain("align_bank_payment_amounts");
    expect((f.doc()[ROWS] as ODataEntity[])[0]!.СуммаПлатежа).toBe(0);
  });

  it("итог содержит документ из контрольного чтения; повтор — квитанция журнала без документа (#50)", async () => {
    const f = await fixture();
    const p = await f.run();
    const r = await f.run({ confirm: true, operationId: data(p).operationId });
    expect(data(r)).toMatchObject({
      verified: true,
      verificationScope: "document_state",
      readbackFetched: true,
      readback: { Ref_Key: DOC, DataVersion: "v2" },
    });
    const replay = await f.run({ confirm: true, operationId: data(p).operationId });
    expect(data(replay)).toMatchObject({ verificationScope: "journal_receipt", replayed: true });
    expect(data(replay).readback).toBeUndefined();
    expect(f.writes).toHaveLength(1);
  });

  it("uses individual override before all-org group", async () => {
    const f = await fixture();
    f.setSettings([
      ...f.settings(),
      {
        ...f.settings()[0],
        Контрагент_Key: CP,
        СчетУчетаРасчетовСПоставщиком_Key: OLD,
        СчетУчетаАвансовВыданных_Key: OLD,
      },
    ]);
    f.setAccount({ Ref_Key: OLD });
    expect(data(await f.run()).resolution).toMatchObject({ choice: { settle: OLD, level: "контрагент" } });
  });
  it.each([{ side: undefined }, { expectedDataVersion: "stale" }, { entitySet: "Catalog_Контрагенты" }])(
    "rejects unsupported or unbound inputs %j",
    async (v) => {
      const f = await fixture();
      expect((await f.run(v)).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it.each([
    { Posted: true },
    { DeletionMark: true },
    { Контрагент_Key: OLD },
    { Организация_Key: OLD },
    { ВалютаДокумента_Key: OLD },
    { [ROWS]: [{}, {}] },
  ])("rejects unsafe document %j", async (v) => {
    const f = await fixture();
    f.setDoc(v);
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each([
    { DeletionMark: true },
    { IsFolder: true },
    { Owner_Key: OLD },
    { Организация_Key: OLD },
    { ВалютаВзаиморасчетов_Key: OLD },
  ])("rejects unsuitable contract %j", async (v) => {
    const f = await fixture();
    f.setContract(v);
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each([
    { DeletionMark: true },
    { IsFolder: true },
    { ЗапретитьИспользоватьВПроводках: true },
    { ЗапретитьИспользоватьВПроводках: undefined },
  ])("rejects unusable account %j", async (v) => {
    const f = await fixture();
    f.setAccount(v);
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each(["settings", "doc"])("rejects changed %s between preview and confirmation", async (which) => {
    const f = await fixture(),
      p = await f.run();
    if (which === "doc") f.setDoc({ DataVersion: "v2" });
    else f.setSettings([{ ...f.settings()[0], Контрагент_Key: CP }]);
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each(["write", "readback"] as const)("holds unknown %s without re-sending PATCH", async (stage) => {
    const f = await fixture(),
      p = await f.run();
    f.fail(stage);
    const args = { confirm: true, operationId: data(p).operationId };
    expect((await f.run(args)).isError).toBe(true);
    expect((await f.run(args)).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
    expect((await f.conn.client.operationEntry(data(p).operationId as string))?.state).toBe(
      "outcome_unknown",
    );
  });
  it("supports explicit supplier side for other writeoff without clearing expense analytic", async () => {
    const f = await fixture("Document_СписаниеСРасчетногоСчета");
    f.setDoc({ ВидОперации: "ПрочееСписание" });
    const p = await f.run();
    expect(p.isError).toBeFalsy();
    expect(data(p).payload).not.toHaveProperty("СубконтоДт1");
    const before = f.doc();
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
    expect(f.doc().СубконтоДт1).toBe(before.СубконтоДт1);
    expect(f.doc().СуммаДокумента).toBe(before.СуммаДокумента);
  });
});

describe("explicit counterparty for an empty bank document", () => {
  it.each([
    { entitySet: "Document_СписаниеСРасчетногоСчета", operation: "ПрочееСписание" },
    { entitySet: SET, operation: "ПрочееПоступление" },
  ])(
    "fills zero CP in $entitySet from explicit selection and preserves zero row amounts/tax/article",
    async ({ entitySet, operation }) => {
      const f = await fixture(entitySet);
      f.setDoc({
        Контрагент_Key: EMPTY,
        ВидОперации: operation,
        [CONTRACT]: EMPTY,
        [ROWS]: [{ ...(f.doc()[ROWS] as object[])[0], [CONTRACT]: EMPTY, СуммаПлатежа: 0 }],
      });
      f.setAccount({ Code: "26" });
      // The saved exact contract setting supplies the account; no account argument is accepted.
      f.setSettings([{ ...f.settings()[0], Контрагент_Key: CP, Договор_Key: CTR }]);
      const before = f.doc();
      const p = await f.run({ counterpartyRef: CP });
      expect(p.isError).toBeFalsy();
      expect(data(p).observedCounterparty).toMatchObject({
        Ref_Key: CP,
        IsFolder: false,
        DeletionMark: false,
      });
      expect(data(p).resolution).toMatchObject({ choice: { settle: ACCOUNT, level: "договор" } });
      const expectedPayload = {
        Контрагент_Key: CP,
        [CONTRACT]: CTR,
        [ACCOUNT_FIELD]: ACCOUNT,
        [ROWS]: [
          {
            ...(before[ROWS] as object[])[0],
            [CONTRACT]: CTR,
            [ACCOUNT_FIELD]: ACCOUNT,
            СчетУчетаРасчетовПоАвансам_Key: ACCOUNT,
          },
        ],
      };
      expect(data(p).payload).toEqual(expectedPayload);
      expect(f.writes).toEqual([]);
      const r = await f.run({ counterpartyRef: CP, confirm: true, operationId: data(p).operationId });
      expect(r.isError).toBeFalsy();
      expect(f.writes).toEqual([
        expect.objectContaining({ method: "PATCH", body: expectedPayload, headers: { "If-Match": "v1" } }),
      ]);
      expect(f.doc()).toEqual({ ...before, ...expectedPayload, DataVersion: "v2" });
    },
  );
  it("requires explicit selection for zero CP instead of inferring from the contract", async () => {
    const f = await fixture();
    f.setDoc({ Контрагент_Key: EMPTY });
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each([undefined, null, "", "not-a-guid", 0])(
    "rejects unknown document CP %j even with explicit selection",
    async (value) => {
      const f = await fixture();
      f.setDoc({ Контрагент_Key: value });
      expect((await f.run({ counterpartyRef: CP })).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it.each([EMPTY, "not-a-guid", "", null])(
    "rejects invalid explicit selection %j",
    async (counterpartyRef) => {
      const f = await fixture();
      f.setDoc({ Контрагент_Key: EMPTY });
      expect((await f.run({ counterpartyRef })).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it("rejects replacement of populated document CP", async () => {
    const f = await fixture();
    expect((await f.run({ counterpartyRef: OLD })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("accepts the same explicit CP without adding it to PATCH", async () => {
    const f = await fixture();
    const p = await f.run({ counterpartyRef: CP });
    expect(p.isError).toBeFalsy();
    expect(data(p).payload).not.toHaveProperty("Контрагент_Key");
    expect(data(p).observedCounterparty).toMatchObject({ Ref_Key: CP });
    expect(
      (await f.run({ counterpartyRef: CP, confirm: true, operationId: data(p).operationId })).isError,
    ).toBeFalsy();
    expect(f.doc().Контрагент_Key).toBe(CP);
  });
  it.each([
    { Ref_Key: OLD },
    { DeletionMark: true },
    { DeletionMark: undefined },
    { IsFolder: true },
    { IsFolder: undefined },
    new ODataError({ kind: "not_found", message: "Synthetic missing counterparty" }),
  ])("rejects an unavailable or unsuitable selected CP %j", async (value) => {
    const f = await fixture();
    f.setDoc({ Контрагент_Key: EMPTY });
    f.setCounterparty(value);
    expect((await f.run({ counterpartyRef: CP })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each([{ Owner_Key: OLD }, { Организация_Key: OLD }, { ВалютаВзаиморасчетов_Key: OLD }])(
    "rejects selected CP/contract relation mismatch %j",
    async (value) => {
      const f = await fixture();
      f.setDoc({ Контрагент_Key: EMPTY });
      f.setContract(value);
      expect((await f.run({ counterpartyRef: CP })).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it.each(["counterparty", "contract", "document", "selection"])(
    "rejects changed %s after explicit-CP preview",
    async (which) => {
      const f = await fixture();
      f.setDoc({ Контрагент_Key: EMPTY });
      const p = await f.run({ counterpartyRef: CP });
      expect(p.isError).toBeFalsy();
      if (which === "counterparty") f.setCounterparty({ Description: "Changed synthetic counterparty" });
      else if (which === "contract") f.setContract({ Description: "Changed synthetic contract" });
      // Keep v1 to prove the observed document is also bound, beyond the version guard.
      else if (which === "document") f.setDoc({ Контрагент_Key: CP });
      const r = await f.run({
        counterpartyRef: which === "selection" ? OLD : CP,
        confirm: true,
        operationId: data(p).operationId,
      });
      expect(r.isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it("still rejects multiple rows when CP is explicitly selected", async () => {
    const f = await fixture();
    f.setDoc({ Контрагент_Key: EMPTY, [ROWS]: [{}, {}] });
    expect((await f.run({ counterpartyRef: CP })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
});

describe("existing bank mutation currency evidence", () => {
  it("preserves the legacy resolver fallback when strict mutation mode is absent", async () => {
    const f = await fixture();
    f.setSettings([{ ...f.settings()[0], Контрагент_Key: CP }]);
    f.setCurrency(new ODataError({ kind: "auth", message: "Synthetic legacy lookup failure" }));
    f.setContract({ РасчетыВУсловныхЕдиницах: undefined });
    expect(
      await settlementAccountsFor(f.conn, { orgKey: ORG, contractRef: CTR, side: "supplier" }),
    ).toMatchObject({
      kind: "РасчетыВВалютеРегламентированногоУчета",
      choice: { settle: ACCOUNT },
    });
    expect(f.writes).toEqual([]);
  });
  const unknownCases = [
    "auth",
    "timeout",
    "unpublished",
    "missing",
    "nonstring",
    "empty",
    "blank",
    "conditional-missing",
    "conditional-nonboolean",
    "conditional-unpublished",
  ] as const;
  for (const stage of ["preview", "confirm"] as const) {
    it.each(unknownCases)(`blocks ${stage} with %s classification evidence without PATCH`, async (reason) => {
      const f = await fixture();
      const preview = stage === "confirm" ? await f.run() : undefined;
      if (preview) expect(preview.isError).toBeFalsy();
      if (reason === "auth" || reason === "timeout")
        f.setCurrency(new ODataError({ kind: reason, message: "Synthetic currency lookup failure" }));
      else if (reason === "unpublished") f.entities.delete("Catalog_Валюты");
      else if (reason === "missing") f.setCurrency({});
      else if (reason === "nonstring") f.setCurrency({ Code: 933 });
      else if (reason === "empty") f.setCurrency({ Code: "" });
      else if (reason === "blank") f.setCurrency({ Code: " \t " });
      else if (reason === "conditional-missing") f.setContract({ РасчетыВУсловныхЕдиницах: undefined });
      else if (reason === "conditional-nonboolean") f.setContract({ РасчетыВУсловныхЕдиницах: "false" });
      else {
        const meta = f.entities.get("Catalog_ДоговорыКонтрагентов")!;
        f.entities.set("Catalog_ДоговорыКонтрагентов", {
          ...meta,
          properties: meta.properties.filter((p) => p.name !== "РасчетыВУсловныхЕдиницах"),
        });
      }
      const result = await f.run(preview ? { confirm: true, operationId: data(preview).operationId } : {});
      expect(result.isError).toBe(true);
      expect(f.writes).toEqual([]);
    });
  }
  it.each([
    { code: " 933 ", conditional: false, kind: "РасчетыВВалютеРегламентированногоУчета" },
    { code: "840", conditional: false, kind: "РасчетыВИностраннойВалюте" },
  ])("previews and confirms known classification $kind", async ({ code, conditional, kind }) => {
    const f = await fixture();
    f.setCurrency({ Code: code });
    f.setContract({ РасчетыВУсловныхЕдиницах: conditional });
    f.setSettings([{ ...f.settings()[0], ВидРасчетовПоДоговору: kind }]);
    const p = await f.run();
    expect(p.isError).toBeFalsy();
    expect(data(p).resolution).toMatchObject({ kind, choice: { settle: ACCOUNT } });
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
  });
  it("does not substitute a regulated-only setting for known foreign currency", async () => {
    const f = await fixture();
    f.setCurrency({ Code: "840" });
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("rejects conditional units even when currency code is missing", async () => {
    const f = await fixture();
    f.setContract({ РасчетыВУсловныхЕдиницах: true });
    f.setSettings([{ ...f.settings()[0], ВидРасчетовПоДоговору: "РасчетыВУсловныхЕдиницах" }]);
    f.setCurrency({});
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("uses the exact validated contract snapshot for classification on each invocation", async () => {
    const f = await fixture();
    const p = await f.run();
    expect(p.isError).toBeFalsy();
    expect(f.contractReads).toHaveLength(1);
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
    expect(f.contractReads).toHaveLength(2);
    expect(f.writes).toHaveLength(1);
  });
});

describe("mutation HTTP retry boundary", () => {
  it("does not retry a bodyless Post after network failure", async () => {
    const f = await fixture();
    vi.mocked(f.conn.client.request).mockRestore();
    const fetchMock = vi.fn(async () => {
      throw new Error("Synthetic connection loss");
    });
    vi.stubGlobal("fetch", fetchMock);
    try {
      await expect(f.conn.client.action(`Document_Synthetic(guid'${DOC}')/Post`)).rejects.toBeInstanceOf(
        ODataError,
      );
      expect(fetchMock).toHaveBeenCalledOnce();
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
describe("version guarded separate Post/Unpost", () => {
  it("explicitly re-posts a modified posted document once, returning the existing full readback", async () => {
    const f = await fixture();
    f.setDoc({ Posted: true, [ACCOUNT_FIELD]: ACCOUNT, НазначениеПлатежа: "Synthetic corrected payment" });
    const p = await f.run({ post: true, repost: true }, "write.document.post_document");
    expect(p.isError).toBeFalsy();
    expect(f.writes).toEqual([]);
    expect(data(p).path).toContain("/Post?$format=json&PostingModeOperational=false");
    const before = f.doc();
    const read = vi.spyOn(f.conn.client, "getEntity");
    const args = { post: true, repost: true, confirm: true, operationId: data(p).operationId };
    const result = await f.run(args, "write.document.post_document");
    expect(result.isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
    expect(f.writes[0]).toMatchObject({
      method: "POST",
      path: `${SET}(guid'${DOC}')/Post?$format=json&PostingModeOperational=false`,
      body: undefined,
    });
    // The existing contract preflight selects two fields; only the version check and
    // the single action readback fetch the full document.
    expect(read.mock.calls.filter(([path]) => path.endsWith("?$format=json"))).toHaveLength(2);
    expect(data(result)).toMatchObject({
      done: true,
      action: "Post",
      verified: true,
      verificationScope: "document_state",
      accountingMovementsVerified: false,
      readbackFetched: true,
      readback: { ...before, DataVersion: "v2" },
    });
    read.mockClear();
    const replay = await f.run(args, "write.document.post_document");
    expect(replay.isError).toBeFalsy();
    expect(data(replay)).toMatchObject({
      done: true,
      action: "Post",
      replayed: true,
      verified: false,
      verificationScope: "journal_receipt",
      accountingMovementsVerified: false,
      readbackFetched: false,
    });
    expect(data(replay).readback).toBeUndefined();
    expect(read).not.toHaveBeenCalled();
    expect(f.writes).toHaveLength(1);
    expect((await f.conn.client.operationEntry(String(data(p).operationId)))?.state).toBe("succeeded");
  });
  it.each([
    { post: true },
    { post: true, repost: false },
    { post: false, repost: true },
    { post: true, repost: true, expectedDataVersion: undefined },
  ])("rejects invalid or implicit re-post before writing: %j", async (args) => {
    const f = await fixture();
    f.setDoc({ Posted: true });
    const result = await f.run(args, "write.document.post_document");
    expect(result.isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("requires preview operationId and unchanged version/intent for explicit re-post", async () => {
    const f = await fixture();
    f.setDoc({ Posted: true });
    const args = { post: true, repost: true };
    expect((await f.run({ ...args, confirm: true }, "write.document.post_document")).isError).toBe(true);
    const p = await f.run(args, "write.document.post_document");
    const confirmation = { ...args, confirm: true, operationId: data(p).operationId };
    expect((await f.run({ ...confirmation, repost: false }, "write.document.post_document")).isError).toBe(
      true,
    );
    f.setDoc({ DataVersion: "stale" });
    expect((await f.run(confirmation, "write.document.post_document")).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each(["write", "readback"] as const)(
    "holds unknown explicit re-post after %s failure",
    async (stage) => {
      const f = await fixture();
      f.setDoc({ Posted: true });
      const p = await f.run({ post: true, repost: true }, "write.document.post_document");
      f.fail(stage);
      const args = { post: true, repost: true, confirm: true, operationId: data(p).operationId };
      expect((await f.run(args, "write.document.post_document")).isError).toBe(true);
      expect((await f.run(args, "write.document.post_document")).isError).toBe(true);
      expect(f.writes).toHaveLength(1);
      expect(f.writes[0]!.path).not.toContain("Unpost");
      const status = await f.run({ operationId: data(p).operationId }, "write.operation.status");
      expect(data(status).status).toBe("unverifiable");
      expect((await f.conn.client.operationEntry(String(data(p).operationId)))?.state).toBe(
        "outcome_unknown",
      );
    },
  );
  it.each([true, false])("previews and executes post=%s once with readback", async (post) => {
    const f = await fixture();
    f.setDoc({ Posted: !post });
    const p = await f.run({ post }, "write.document.post_document");
    expect(p.isError).toBeFalsy();
    expect(f.writes).toEqual([]);
    const args = { post, confirm: true, operationId: data(p).operationId };
    expect((await f.run(args, "write.document.post_document")).isError).toBeFalsy();
    expect((await f.run(args, "write.document.post_document")).isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
    expect(f.doc().Posted).toBe(post);
    expect(f.writes[0]!.path.includes("PostingModeOperational=false")).toBe(post);
  });
  it("rejects changed version before action", async () => {
    const f = await fixture(),
      p = await f.run({ post: true }, "write.document.post_document");
    f.setDoc({ DataVersion: "changed" });
    expect(
      (
        await f.run(
          { post: true, confirm: true, operationId: data(p).operationId },
          "write.document.post_document",
        )
      ).isError,
    ).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each(
    [true, false].flatMap((post) => (["write", "readback"] as const).map((stage) => ({ post, stage }))),
  )("holds unknown post=$post $stage and status cannot clear hold", async ({ post, stage }) => {
    const f = await fixture();
    f.setDoc({ Posted: !post });
    const p = await f.run({ post }, "write.document.post_document");
    f.fail(stage);
    const args = { post, confirm: true, operationId: data(p).operationId };
    expect((await f.run(args, "write.document.post_document")).isError).toBe(true);
    expect((await f.run(args, "write.document.post_document")).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
    const status = await f.run({ operationId: data(p).operationId }, "write.operation.status");
    expect(data(status).status).toBe("unverifiable");
    expect((await f.conn.client.operationEntry(data(p).operationId as string))?.state).toBe(
      "outcome_unknown",
    );
  });
});

describe("PR43 real BP30 compatibility and bounded readback", () => {
  it.each([
    { set: SET, operation: "ОплатаПокупателя", side: "customer", kind: "СПокупателем" },
    {
      set: "Document_СписаниеСРасчетногоСчета",
      operation: "ОплатаПоставщику",
      side: "supplier",
      kind: "СПоставщиком",
    },
  ])("previews $side with actual BP30 metadata and enum", async ({ set, operation, side, kind }) => {
    const f = await fixture(set);
    vi.mocked(f.conn.getMetadata).mockResolvedValue(bp30Metadata());
    f.setDoc({
      Контрагент: CP,
      Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
      ВидОперации: operation,
      [ROWS]: [
        {
          LineNumber: "1",
          [CONTRACT]: OLD,
          [ACCOUNT_FIELD]: OLD,
          СчетУчетаРасчетовПоАвансам_Key: OLD,
          СуммаПлатежа: 150,
        },
      ],
    });
    f.setContract({ ВидДоговора: kind });
    f.setSettings([
      {
        ...f.settings()[0],
        СчетУчетаРасчетовСПокупателем_Key: ACCOUNT,
        СчетУчетаАвансовПолученных_Key: ACCOUNT,
        СчетУчетаАвансовВыданных_Key: ACCOUNT,
      },
    ]);
    expect([...bp30Enums().values()].some((values) => values.has(operation))).toBe(true);
    const p = await f.run({ side });
    expect(p.isError, JSON.stringify(p)).toBeFalsy();
    expect(data(p).payload).not.toHaveProperty("Контрагент_Key");
    expect((data(p).payload as ODataEntity)[ROWS]).toMatchObject([
      { СчетУчетаРасчетовПоАвансам_Key: ACCOUNT },
    ]);
    expect((await f.run({ side, confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
  });
  it("changes both saved settlement and advance accounts in the payment row", async () => {
    const f = await fixture();
    f.setSettings([{ ...f.settings()[0], СчетУчетаАвансовВыданных_Key: ACCOUNT }]);
    const p = await f.run();
    expect(p.isError).toBeFalsy();
    expect((data(p).payload as ODataEntity)[ROWS]).toMatchObject([
      { СчетУчетаРасчетовПоАвансам_Key: ACCOUNT },
    ]);
  });
  it.each([CUR, OLD])(
    "rejects every conditional-unit contract before currency/settings lookup (%s)",
    async (contractCurrency) => {
      const f = await fixture();
      f.setContract({ РасчетыВУсловныхЕдиницах: true, ВалютаВзаиморасчетов_Key: contractCurrency });
      f.setCurrency(new ODataError({ kind: "auth", message: "must not read currency" }));
      const r = await f.run();
      expect(r.isError).toBe(true);
      expect(JSON.stringify(r)).toContain("Расчёты в условных единицах этим инструментом не поддерживаются");
      expect(f.writes).toEqual([]);
    },
  );
  it("accepts recalculated settlement rate fields while protecting payment amount", async () => {
    const f = await fixture();
    const p = await f.run();
    const row = (f.doc()[ROWS] as ODataEntity[])[0]!;
    f.changeOnWrite({
      [ROWS]: [
        {
          ...row,
          [CONTRACT]: CTR,
          [ACCOUNT_FIELD]: ACCOUNT,
          СчетУчетаРасчетовПоАвансам_Key: ACCOUNT,
          КурсВзаиморасчетов: 2,
          КратностьВзаиморасчетов: 1,
          СуммаВзаиморасчетов: 75,
        },
      ],
    });
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
  });
  it.each([{ СуммаДокумента: 151 }, { НазначениеПлатежа: "changed" }, { СубконтоДт1: "lost" }])(
    "holds changed protected business field %j",
    async (change) => {
      const f = await fixture();
      const p = await f.run();
      f.changeOnWrite(change);
      expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBe(true);
      expect((await f.conn.client.operationEntry(String(data(p).operationId)))?.state).toBe(
        "outcome_unknown",
      );
    },
  );
  it("legacy lookup skips hierarchy entirely even if metadata or CP state is unsuitable", async () => {
    const f = await fixture();
    f.entities.delete("Catalog_Контрагенты");
    f.setCounterparty({ DeletionMark: true });
    const get = vi.spyOn(f.conn.client, "getEntity");
    const r = await settlementAccountsFor(f.conn, { orgKey: ORG, contractRef: CTR, side: "supplier" });
    expect(r.kind).toBe("РасчетыВВалютеРегламентированногоУчета");
    expect(get.mock.calls.filter(([path]) => path.startsWith("Catalog_Контрагенты"))).toHaveLength(0);
  });
});

const belarusSchema = JSON.parse(
  readFileSync(new URL("./fixtures/metadata/belarus-2.1.json", import.meta.url), "utf8"),
) as {
  entities: Array<{ entitySet: string; keys: string[]; properties: EntityMeta["properties"] }>;
};
describe("actual Belarus 2.1 schema (platform 8.3.25), synthetic business values", () => {
  it.each([SET, "Document_СписаниеСРасчетногоСчета"])(
    "keeps GUID CP and Other/loan mode for %s",
    async (set) => {
      const f = await fixture(set);
      for (const entity of belarusSchema.entities) {
        const prior = f.entities.get(entity.entitySet);
        f.entities.set(entity.entitySet, {
          ...prior,
          ...entity,
          entityType: entity.entitySet,
          shortName: entity.entitySet,
          navigations: [],
          class: entity.entitySet.startsWith("Document_")
            ? "document"
            : entity.entitySet.startsWith("InformationRegister_")
              ? "informationRegister"
              : "catalog",
        });
      }
      const rowMeta = f.entities.get(`${set}_${ROWS}`)!;
      f.setDoc({
        ВидОперации: set === SET ? "ПрочееПоступление" : "ПрочееСписание",
        [ROWS]: [
          (f.doc()[ROWS] as ODataEntity[]).map((row) =>
            Object.fromEntries(
              Object.entries(row).filter(([field]) => rowMeta.properties.some((p) => p.name === field)),
            ),
          )[0],
        ],
      });
      const p = await f.run();
      expect(p.isError, JSON.stringify(p)).toBeFalsy();
      expect((data(p).payload as ODataEntity)[ROWS]).toMatchObject([
        { [ACCOUNT_FIELD]: ACCOUNT, СчетУчетаРасчетовПоАвансам_Key: ACCOUNT },
      ]);
      expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBeFalsy();
      expect(f.doc().Контрагент_Key).toBe(CP);
      expect(data(p).resolution).toMatchObject({
        choice: { settle: ACCOUNT, provenance: { settle: { level: "группа" } } },
      });
    },
  );
});

it("legacy lookup keeps its bounded register fallback instead of rejecting 500 rows", async () => {
  const f = await fixture();
  f.setSettings(Array.from({ length: 500 }, () => ({ ...f.settings()[0], Контрагент_Key: CP })));
  expect(
    await settlementAccountsFor(f.conn, { orgKey: ORG, contractRef: CTR, side: "supplier" }),
  ).toMatchObject({ choice: { settle: ACCOUNT } });
});
it.each(["StandardODATA.Catalog_ФизическиеЛица", "", undefined])(
  "rejects unknown composite counterparty discriminator %s",
  async (type) => {
    const f = await fixture();
    vi.mocked(f.conn.getMetadata).mockResolvedValue(bp30Metadata());
    f.setDoc({ Контрагент: CP, Контрагент_Type: type });
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  },
);
it("validates the inherited advance account as well as the settlement account", async () => {
  const f = await fixture();
  f.setSettings([{ ...f.settings()[0], СчетУчетаАвансовВыданных_Key: OLD }]);
  expect((await f.run()).isError).toBe(true);
  expect(f.writes).toEqual([]);
});
it("holds when server changes row payment, tax or cash-flow analytics", async () => {
  for (const change of [{ СуммаПлатежа: 151 }, { СуммаНДС: 7 }, { СтатьяДвиженияДенежныхСредств_Key: CP }]) {
    const f = await fixture();
    const p = await f.run();
    f.changeOnWrite({
      [ROWS]: [
        {
          ...(f.doc()[ROWS] as ODataEntity[])[0],
          [CONTRACT]: CTR,
          [ACCOUNT_FIELD]: ACCOUNT,
          СчетУчетаРасчетовПоАвансам_Key: ACCOUNT,
          ...change,
        },
      ],
    });
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
  }
});
it("fills only a zero composite counterparty of the published catalog type", async () => {
  const f = await fixture();
  vi.mocked(f.conn.getMetadata).mockResolvedValue(bp30Metadata());
  f.setDoc({
    Контрагент: EMPTY,
    Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
    ВидОперации: "ПрочееПоступление",
    [ROWS]: [],
  });
  const p = await f.run({ counterpartyRef: CP });
  expect(p.isError, JSON.stringify(p)).toBeFalsy();
  expect(data(p).payload).toMatchObject({ Контрагент: CP });
  expect(data(p).payload).not.toHaveProperty("Контрагент_Key");
  expect(
    (await f.run({ counterpartyRef: CP, confirm: true, operationId: data(p).operationId })).isError,
  ).toBeFalsy();
});

// Real schemas publish header business fields beyond СуммаДокумента.
describe("published header business values survive settlement PATCH", () => {
  it.each([
    { set: SET, field: "НДС_УСН", before: 0, after: 7 },
    { set: SET, field: "СуммаУслуг", before: 0, after: 11 },
    { set: "Document_СписаниеСРасчетногоСчета", field: "СуммаВПенсионный", before: 0, after: 3 },
    { set: "Document_СписаниеСРасчетногоСчета", field: "СуммаВСоцстрах", before: 0, after: 4 },
    { set: "Document_СписаниеСРасчетногоСчета", field: "НДСпоТС", before: false, after: true },
  ])("holds changed $field and never resends for $set", async ({ set, field, before, after }) => {
    const f = await fixture(set);
    for (const entity of belarusSchema.entities) {
      const prior = f.entities.get(entity.entitySet);
      f.entities.set(entity.entitySet, {
        ...prior,
        ...entity,
        entityType: entity.entitySet,
        shortName: entity.entitySet,
        navigations: [],
        class: entity.entitySet.startsWith("Document_")
          ? "document"
          : entity.entitySet.startsWith("InformationRegister_")
            ? "informationRegister"
            : "catalog",
      });
    }
    // НДСпоТС belongs to the outgoing schema.
    const header = f.entities.get(set)!;
    expect(header.properties.some((p) => p.name === field)).toBe(true);
    f.setDoc({
      [field]: before,
      [ROWS]: [],
      ВидОперации: set === SET ? "ПрочееПоступление" : "ПрочееСписание",
    });
    const preview = await f.run();
    expect(preview.isError, JSON.stringify(preview)).toBeFalsy();
    f.changeOnWrite({ [field]: after });
    const args = { confirm: true, operationId: data(preview).operationId };
    expect((await f.run(args)).isError).toBe(true);
    expect((await f.conn.client.operationEntry(String(args.operationId)))?.state).toBe("outcome_unknown");
    expect((await f.run(args)).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
  });
});

describe("paymentAmountWarnings (#61)", () => {
  it("совпадает — без предупреждения; нет строк (налоги, прочее) — молчим; суммы неизвестны — молчим", () => {
    expect(
      paymentAmountWarnings({ СуммаДокумента: 100 }, [{ СуммаПлатежа: 60 }, { СуммаПлатежа: 40 }]),
    ).toEqual([]);
    expect(paymentAmountWarnings({ СуммаДокумента: 100 }, [])).toEqual([]);
    expect(paymentAmountWarnings({ СуммаДокумента: 100 }, [{ СуммаПлатежа: "?" }])).toEqual([]);
    expect(paymentAmountWarnings({}, [{ СуммаПлатежа: 1 }])).toEqual([]);
  });
});

describe("align_bank_payment_amounts и запрет проведения при расхождении сумм (#61)", () => {
  const ALIGN = "write.document.align_bank_payment_amounts";
  const POST = "write.document.post_document";
  /** Строка с полями взаиморасчётов, как в БП 3.0; суммы строки — пустые. */
  async function broken() {
    const f = await fixture();
    const rowSet = `${SET}_${ROWS}`;
    const rowMeta = f.entities.get(rowSet)!;
    f.entities.set(rowSet, {
      ...rowMeta,
      properties: [
        ...rowMeta.properties,
        { name: "СуммаВзаиморасчетов", type: "Edm.Double", nullable: true },
        { name: "КурсВзаиморасчетов", type: "Edm.Double", nullable: true },
        { name: "КратностьВзаиморасчетов", type: "Edm.Int64", nullable: true },
      ],
    });
    const row = (f.doc()[ROWS] as ODataEntity[])[0]!;
    f.setDoc({
      [ROWS]: [
        {
          ...row,
          СуммаПлатежа: 0,
          СуммаВзаиморасчетов: 0,
          КурсВзаиморасчетов: 0,
          КратностьВзаиморасчетов: 0,
        },
      ],
    });
    return f;
  }
  const rowOf = (f: Awaited<ReturnType<typeof fixture>>) => (f.doc()[ROWS] as ODataEntity[])[0]!;

  it("предпросмотр «было → станет» без записи, подтверждение — один PATCH, остальное сохранено", async () => {
    const f = await broken();
    const p = await f.run({}, ALIGN);
    expect(p.isError).toBeFalsy();
    const changes = data(p).changes as Array<{ field: string; from: unknown; to: unknown }>;
    expect(changes).toEqual(
      expect.arrayContaining([
        { field: "СуммаПлатежа", from: 0, to: 150 },
        { field: "СуммаВзаиморасчетов", from: 0, to: 150 },
        { field: "КурсВзаиморасчетов", from: 0, to: 1 },
        { field: "КратностьВзаиморасчетов", from: 0, to: 1 },
        { field: "СуммаНДС", from: 0, to: 25 },
      ]),
    );
    expect(f.writes).toHaveLength(0);
    const r = await f.run({ confirm: true, operationId: data(p).operationId }, ALIGN);
    expect(r.isError).toBeFalsy();
    expect(data(r)).toMatchObject({ updated: true, verified: true });
    expect(f.writes).toHaveLength(1);
    expect(f.writes[0]).toMatchObject({ method: "PATCH", headers: { "If-Match": "v1" } });
    expect(rowOf(f)).toMatchObject({ СуммаПлатежа: 150, СуммаВзаиморасчетов: 150, СуммаНДС: 25 });
    expect(rowOf(f).Субконто).toBe("preserved");
    expect(f.doc().СуммаДокумента).toBe(150);
  });

  it.each([
    [
      "несколько строк",
      (f: Awaited<ReturnType<typeof broken>>) => f.setDoc({ [ROWS]: [rowOf(f), rowOf(f)] }),
      "Несколько строк",
    ],
    ["проведён", (f: Awaited<ReturnType<typeof broken>>) => f.setDoc({ Posted: true }), "проведён"],
    ["валюта", (f: Awaited<ReturnType<typeof broken>>) => f.setCurrency({ Code: "840" }), "валюте"],
    [
      "условные единицы",
      (f: Awaited<ReturnType<typeof broken>>) => f.setContract({ РасчетыВУсловныхЕдиницах: true }),
      "условных единицах",
    ],
    [
      "неизвестная ставка",
      (f: Awaited<ReturnType<typeof broken>>) => f.setDoc({ [ROWS]: [{ ...rowOf(f), СтавкаНДС: "НДСЧто" }] }),
      "Ставка НДС",
    ],
  ] as const)("отказ без записи: %s", async (_label, spoil, reason) => {
    const f = await broken();
    spoil(f);
    const r = await f.run({}, ALIGN);
    expect(r.isError).toBe(true);
    expect(JSON.stringify(r.content)).toContain(reason);
    expect(f.writes).toHaveLength(0);
  });

  it("Int64 строкой, как отдаёт 1С: «1» не считается изменением и отправляется строкой", async () => {
    const f = await broken();
    f.setDoc({ [ROWS]: [{ ...rowOf(f), КратностьВзаиморасчетов: "1", КурсВзаиморасчетов: 1 }] });
    const p = await f.run({}, ALIGN);
    const fields = (data(p).changes as Array<{ field: string }>).map((c) => c.field);
    expect(fields).not.toContain("КратностьВзаиморасчетов");
    expect(fields).not.toContain("КурсВзаиморасчетов");
    await f.run({ confirm: true, operationId: data(p).operationId }, ALIGN);
    const sent = (f.writes[0]!.body as { РасшифровкаПлатежа: ODataEntity[] }).РасшифровкаПлатежа[0]!;
    expect(sent.КратностьВзаиморасчетов).toBe("1");
  });

  it("неизменившиеся числа — «уже совпадают», даже если 1С отдаёт их строкой", async () => {
    const f = await broken();
    f.setDoc({
      [ROWS]: [
        {
          ...rowOf(f),
          СуммаПлатежа: 150,
          СуммаВзаиморасчетов: 150,
          СуммаНДС: 25,
          КурсВзаиморасчетов: 1,
          КратностьВзаиморасчетов: "1",
        },
      ],
    });
    expect(JSON.stringify((await f.run({}, ALIGN)).content)).toContain("уже совпадают");
  });

  it("суммы уже совпадают — отказ «изменение не нужно»", async () => {
    const f = await fixture();
    const rowSet = `${SET}_${ROWS}`;
    const rowMeta = f.entities.get(rowSet)!;
    f.entities.set(rowSet, {
      ...rowMeta,
      properties: [
        ...rowMeta.properties,
        { name: "СуммаВзаиморасчетов", type: "Edm.Double", nullable: true },
      ],
    });
    f.setDoc({ [ROWS]: [{ ...rowOf(f), СуммаПлатежа: 150, СуммаВзаиморасчетов: 150, СуммаНДС: 25 }] });
    const r = await f.run({}, ALIGN);
    expect(JSON.stringify(r.content)).toContain("уже совпадают");
  });

  it("проведение отказывает до отправки, пока суммы расходятся; после выравнивания — проводит", async () => {
    const f = await broken();
    // Простое проведение (без версии) — отказ на подтверждении, до отправки Post.
    const blocked = await f.run({ post: true, confirm: true, expectedDataVersion: undefined }, POST);
    expect(blocked.isError).toBe(true);
    expect(JSON.stringify(blocked.content)).toContain("Проведение не отправлено");
    expect(JSON.stringify(blocked.content)).toContain("align_bank_payment_amounts");
    expect(f.writes).toHaveLength(0);
    // Предпросмотр проведения показывает тот же отказ — до подтверждения.
    expect((await f.run({ post: true }, POST)).isError).toBe(true);

    const p = await f.run({}, ALIGN);
    await f.run({ confirm: true, operationId: data(p).operationId }, ALIGN);
    const pp = await f.run({ post: true, expectedDataVersion: "v2" }, POST);
    expect(pp.isError).toBeFalsy();
  });

  it("сумма совпадает, но пуста сумма взаиморасчётов — тоже отказ проведения", async () => {
    const f = await broken();
    f.setDoc({ [ROWS]: [{ ...rowOf(f), СуммаПлатежа: 150 }] });
    const r = await f.run({ post: true }, POST);
    expect(r.isError).toBe(true);
    expect(JSON.stringify(r.content)).toContain("не заполнена сумма взаиморасчётов");
  });
});
