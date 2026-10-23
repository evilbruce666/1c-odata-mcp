import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Connection } from "../src/context.js";
import { parseConfig } from "../src/config/env.js";
import { createServer } from "../src/mcp/server.js";
import { ODataError } from "../src/odata/errors.js";
import type { EntityMeta, ODataEntity } from "../src/types/odata.js";
import type { CallToolResult } from "@modelcontextprotocol/server";

const ORG = "11111111-1111-4111-8111-111111111111";
const CP = "22222222-2222-4222-8222-222222222222";
const CTR = "33333333-3333-4333-8333-333333333333";
const ACCOUNT = "44444444-4444-4444-8444-444444444444";
const OTHER = "55555555-5555-4555-8555-555555555555";
const CURRENCY = "66666666-6666-4666-8666-666666666666";
const EMPTY = "00000000-0000-0000-0000-000000000000";
const REG = "InformationRegister_СчетаУчетаРасчетовСКонтрагентами";
const KIND = "РасчетыВВалютеРегламентированногоУчета";
const key = { Организация_Key: ORG, Контрагент_Key: CP, Договор_Key: CTR, ВидРасчетовПоДоговору: KIND };
const dimensions = Object.keys(key);
const resources = [
  "СчетУчетаРасчетовСПокупателем_Key",
  "СчетУчетаАвансовПолученных_Key",
  "СчетУчетаРасчетовСПоставщиком_Key",
  "СчетУчетаАвансовВыданных_Key",
];
const dirs: string[] = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(dirs.splice(0).map((p) => rm(p, { recursive: true, force: true })));
});

async function fixture(
  existing?: ODataEntity,
  options: {
    readOnly?: boolean;
    writable?: boolean;
    versionless?: boolean;
    currencyCode?: string;
    perBaseCurrency?: string;
    globalCurrency?: string;
  } = {},
) {
  const dir = await mkdtemp(join(tmpdir(), "settings-test-"));
  dirs.push(dir);
  const cfg = parseConfig({
    ODATA_BASE_URL: "https://example.invalid/odata/",
    ODATA_USERNAME: "synthetic",
    ODATA_PASSWORD: "synthetic",
    READ_ONLY: String(options.readOnly ?? false),
    ODATA_WRITABLE: String(options.writable ?? true),
    WRITE_JOURNAL_DIR: dir,
    ODATA_REGULATED_CURRENCY_CODE: options.globalCurrency ?? "933",
    ...(options.perBaseCurrency
      ? {
          ODATA_DB_DEFAULT_BASE_URL: "https://example.invalid/odata/",
          ODATA_DB_DEFAULT_USERNAME: "synthetic",
          ODATA_DB_DEFAULT_PASSWORD: "synthetic",
          ODATA_DB_DEFAULT_REGULATED_CURRENCY_CODE: options.perBaseCurrency,
        }
      : {}),
  });
  const conn = new Connection(
    { ...cfg.connections[0]!, writable: options.writable ?? true },
    { ...cfg.behavior, writeJournalDir: dir },
  );
  const props = [
    ...dimensions,
    ...resources,
    ...(options.versionless ? [] : ["DataVersion"]),
    "СчетПрочихРасчетов_Key",
  ].map((name) => ({
    name,
    type: name === "ВидРасчетовПоДоговору" || name === "DataVersion" ? "Edm.String" : "Edm.Guid",
    nullable: true,
  }));
  const meta: EntityMeta = {
    entitySet: REG,
    entityType: REG,
    class: "informationRegister",
    shortName: "СчетаУчетаРасчетовСКонтрагентами",
    keys: [...dimensions],
    properties: props,
    navigations: [],
  };
  const entities = new Map([
    [REG, meta],
    [
      "Catalog_Контрагенты",
      {
        ...meta,
        properties: ["Ref_Key", "Parent_Key", "IsFolder", "DeletionMark"].map((name) => ({
          name,
          type: name.endsWith("_Key") ? "Edm.Guid" : "Edm.Boolean",
          nullable: false,
        })),
      },
    ],
    ["Catalog_Организации", meta],
    ...["Catalog_ДоговорыКонтрагентов", "Catalog_Валюты", "ChartOfAccounts_Хозрасчетный"].map(
      (name) => [name, { ...meta, entitySet: name }] as const,
    ),
  ]);
  vi.spyOn(conn, "getMetadata").mockImplementation(async () => ({ entities, odataVersion: "3.0" }));
  let row = existing ? { ...existing } : undefined;
  let contract = {
    Ref_Key: CTR,
    Owner_Key: CP,
    Организация_Key: ORG,
    ВидДоговора: "Прочее",
    ВалютаВзаиморасчетов_Key: CURRENCY,
    РасчетыВУсловныхЕдиницах: false,
  };
  let subject = { Ref_Key: CP, IsFolder: false, DeletionMark: false, Parent_Key: EMPTY };
  let organization = { Ref_Key: ORG, DeletionMark: false };
  let corruptReadback = false;
  let readbackError: ODataError | undefined;
  let writeError: ODataError | undefined;
  const writes: Array<{ path: string; method: string; body: unknown; headers: unknown }> = [];
  vi.spyOn(conn.client, "request").mockImplementation(
    async (path: string, method = "GET", body?: unknown, headers?: unknown) => {
      if (method !== "GET") {
        writes.push({ path, method, body, headers });
        if (writeError) throw writeError;
        row = { ...row, ...(body as object), ...(options.versionless ? {} : { DataVersion: "v2" }) };
        if (corruptReadback) row[resources[2]!] = OTHER;
        return {} as never;
      }
      if (path.startsWith("Catalog_Контрагенты")) return subject as never;
      if (path.startsWith("Catalog_Организации")) return organization as never;
      if (path.startsWith("Catalog_Договоры")) return contract as never;
      if (path.startsWith("Catalog_Валюты")) return { Code: options.currencyCode ?? "933" } as never;
      if (path.startsWith("ChartOfAccounts_"))
        return { value: [{ Ref_Key: ACCOUNT, Code: "66.2", Description: "Synthetic" }] } as never;
      if (path.startsWith(REG)) {
        if (writes.length && readbackError) throw readbackError;
        if (!row) throw new ODataError({ kind: "not_found", message: "Absent" });
        return { ...row } as never;
      }
      throw new Error(`Unexpected synthetic path ${path}`);
    },
  );
  const tools = (
    createServer({ db: () => conn } as never) as unknown as {
      _registeredTools: Record<string, { handler: (args: object, extra: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  const run = async (args: object = {}, read = false, generic = false) => {
    const tool =
      tools[
        generic
          ? read
            ? "read.accounting.get_settlement_settings"
            : "write.accounting.set_settlement_settings"
          : read
            ? "read.accounting.get_contract_settlement_settings"
            : "write.accounting.set_contract_settlement_settings"
      ];
    expect(tool, "dedicated exact saved-settings tool must be registered").toBeDefined();
    return tool!.handler(
      {
        database: "default",
        organizationRef: ORG,
        counterpartyRef: CP,
        contractRef: CTR,
        settlementKind: KIND,
        mode: "create",
        accounts: { supplierSettlement: ACCOUNT },
        confirm: false,
        ...args,
      },
      {},
    );
  };
  return {
    run,
    setSubject: (v: Partial<typeof subject>) => {
      subject = { ...subject, ...v };
    },
    setOrganization: (v: Partial<typeof organization>) => {
      organization = { ...organization, ...v };
    },
    conn,
    meta,
    entities,
    writes,
    status: (operationId: unknown) =>
      tools["write.operation.status"]!.handler({ database: "default", operationId }, {}),
    failReadback: (error: ODataError) => {
      readbackError = error;
    },
    failWrite: (error: ODataError) => {
      writeError = error;
    },
    setRow: (r: ODataEntity | undefined) => {
      row = r;
    },
    setContract: (r: Partial<typeof contract>) => {
      contract = { ...contract, ...r };
    },
    corrupt: () => {
      corruptReadback = true;
    },
  };
}
const data = (r: CallToolResult) => r.structuredContent as Record<string, unknown>;
const saved = () => ({
  ...key,
  DataVersion: "v1",
  [resources[0]!]: OTHER,
  [resources[2]!]: OTHER,
  СчетПрочихРасчетов_Key: OTHER,
});

describe("exact contract saved settlement settings", () => {
  it("reads regulated BYN933 and RUB643 settings in the same process without changing global fallback", async () => {
    const by = await fixture(saved(), { globalCurrency: "643", perBaseCurrency: "933", currencyCode: "933" });
    const ru = await fixture(saved(), { globalCurrency: "643", currencyCode: "643" });
    for (const f of [by, ru]) {
      const result = await f.run({}, true);
      expect(result.isError, JSON.stringify(result)).toBeFalsy();
      expect(data(result).found).toBe(true);
      expect(f.writes).toEqual([]);
    }
    expect(by.conn.behavior.regulatedCurrencyCode).toBe("933");
    expect(ru.conn.behavior.regulatedCurrencyCode).toBe("643");
  });
  it("reads exact compound-key row without inherited defaults", async () => {
    const f = await fixture(saved());
    const r = await f.run({}, true);
    expect(r.isError).toBeFalsy();
    expect(data(r).found).toBe(true);
    expect(data(r).record).toMatchObject(saved());
    expect(f.writes).toEqual([]);
  });
  it("previews and creates once, using all dimensions and no Ref_Key; confirms by readback", async () => {
    const f = await fixture();
    const preview = await f.run();
    expect(preview.isError).toBeFalsy();
    expect(data(preview)).toMatchObject({ dryRun: true, method: "POST" });
    expect(f.writes).toEqual([]);
    const operationId = data(preview).operationId;
    const result = await f.run({ confirm: true, operationId });
    expect(result.isError).toBeFalsy();
    expect(data(result).verified).toBe(true);
    expect(f.writes).toEqual([
      {
        path: `${REG}?$format=json`,
        method: "POST",
        body: { ...key, [resources[2]!]: ACCOUNT },
        headers: {},
      },
    ]);
    const replay = await f.run({ confirm: true, operationId });
    expect(data(replay).replayed).toBe(true);
    expect(f.writes).toHaveLength(1);
  });
  it("updates only explicit resources with exact compound path and If-Match, preserving other resources", async () => {
    const f = await fixture(saved());
    const args = { mode: "update", expectedDataVersion: "v1" };
    const p = await f.run(args);
    const r = await f.run({ ...args, confirm: true, operationId: data(p).operationId });
    expect(r.isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
    expect(f.writes[0]).toMatchObject({
      method: "PATCH",
      body: { [resources[2]!]: ACCOUNT },
      headers: { "If-Match": "v1" },
    });
    for (const name of dimensions) expect(f.writes[0]!.path).toContain(`${name}=`);
    expect(data(r).record).toMatchObject({ [resources[0]!]: OTHER, СчетПрочихРасчетов_Key: OTHER });
  });
  it.each(["RecordSet", "Recorder_Key", "Period"])(
    "blocks unsupported schema %s before any write",
    async (field) => {
      const f = await fixture();
      f.meta.properties.push({ name: field, type: "Edm.String", nullable: true });
      const r = await f.run();
      expect(r.isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it("blocks missing register and extra compound key instead of broad writes", async () => {
    const f = await fixture();
    f.entities.delete(REG);
    const missing = await f.run();
    expect(missing.isError).toBe(true);
    expect(missing.content).toContainEqual({
      type: "text",
      text: "Регистр сохранённых счетов расчётов (InformationRegister_СчетаУчетаРасчетовСКонтрагентами) отсутствует в доступной OData-схеме; проверьте имя, состав публикации и права. Настройка не читается и не записывается.",
    });
    const g = await fixture();
    g.meta.keys.push("Unknown");
    expect((await g.run()).isError).toBe(true);
    expect(g.writes).toEqual([]);
  });
  it.each([
    { organizationRef: EMPTY },
    { counterpartyRef: EMPTY },
    { contractRef: EMPTY },
    { accounts: {} },
    { accounts: { arbitrary: ACCOUNT } },
  ])("rejects wildcard or unknown resource arguments %j", async (args) => {
    const f = await fixture();
    expect((await f.run(args)).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each([{ Owner_Key: OTHER }, { Организация_Key: OTHER }])(
    "rejects contract scope mismatch %j",
    async (contract) => {
      const f = await fixture();
      f.setContract(contract);
      expect((await f.run()).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it("blocks create over existing row and update without an exact current version", async () => {
    const f = await fixture(saved());
    expect((await f.run()).isError).toBe(true);
    expect((await f.run({ mode: "update" })).isError).toBe(true);
    expect((await f.run({ mode: "update", expectedDataVersion: "stale" })).isError).toBe(true);
    f.setRow({ ...key });
    expect((await f.run({ mode: "update", expectedDataVersion: "v1" })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("blocks changed confirmation arguments and stale version between preview and confirmation", async () => {
    const f = await fixture(saved());
    const p = await f.run({ mode: "update", expectedDataVersion: "v1" });
    f.setRow({ ...saved(), DataVersion: "v2" });
    expect(
      (
        await f.run({
          mode: "update",
          expectedDataVersion: "v1",
          operationId: data(p).operationId,
          confirm: true,
        })
      ).isError,
    ).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("marks readback mismatch uncertain and prevents re-send with the same operationId", async () => {
    const f = await fixture();
    const p = await f.run();
    f.corrupt();
    const args = { confirm: true, operationId: data(p).operationId };
    expect((await f.run(args)).isError).toBe(true);
    expect((await f.run(args)).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
    const entry = await f.conn.client.operationEntry(data(p).operationId as string);
    expect(entry?.state).toBe("outcome_unknown");
  });
  it.each(
    (["create", "update"] as const).flatMap((mode) =>
      (["auth", "not_found", "bad_request"] as const).map((kind) => ({ mode, kind })),
    ),
  )(
    "keeps successful $mode with $kind readback failure unknown and blocks replay",
    async ({ mode, kind }) => {
      const f = await fixture(mode === "update" ? saved() : undefined);
      const input = { mode, ...(mode === "update" ? { expectedDataVersion: "v1" } : {}) };
      const preview = await f.run(input);
      expect(preview.isError).toBeFalsy();
      const operationId = data(preview).operationId;
      const originalError = new ODataError({ kind, message: "Synthetic post-write readback failure" });
      f.failReadback(originalError);
      const execute = f.conn.client.writeRecordOnce.bind(f.conn.client);
      let failure: unknown;
      vi.spyOn(f.conn.client, "writeRecordOnce").mockImplementation(async (...args) => {
        try {
          return await execute(...args);
        } catch (error) {
          failure = error;
          throw error;
        }
      });

      const result = await f.run({ ...input, confirm: true, operationId });
      expect(result.isError).toBe(true);
      expect(f.writes).toHaveLength(1);
      expect(f.writes[0]!.method).toBe(mode === "create" ? "POST" : "PATCH");
      expect(await f.conn.client.operationEntry(operationId as string)).toMatchObject({
        state: "outcome_unknown",
        errorKind: "unknown",
      });
      expect(failure).toBeInstanceOf(ODataError);
      expect((failure as ODataError).cause).toBeInstanceOf(ODataError);
      const verificationError = (failure as ODataError).cause as ODataError;
      expect(verificationError.kind).toBe("unknown");
      expect(verificationError.cause).toBe(originalError);
      const status = await f.status(operationId);
      expect(data(status).status).toBe("unverifiable");
      expect((await f.run({ ...input, confirm: true, operationId })).isError).toBe(true);
      expect(f.writes).toHaveLength(1);
    },
  );
  it.each(
    (["create", "update"] as const).flatMap((mode) =>
      (["auth", "not_found", "bad_request"] as const).map((kind) => ({ mode, kind })),
    ),
  )("keeps $kind rejection from the $mode request definitive", async ({ mode, kind }) => {
    const f = await fixture(mode === "update" ? saved() : undefined);
    const input = { mode, ...(mode === "update" ? { expectedDataVersion: "v1" } : {}) };
    const preview = await f.run(input);
    expect(preview.isError).toBeFalsy();
    const operationId = data(preview).operationId;
    f.failWrite(new ODataError({ kind, message: "Synthetic mutation rejection" }));

    expect((await f.run({ ...input, confirm: true, operationId })).isError).toBe(true);
    expect(await f.conn.client.operationEntry(operationId as string)).toMatchObject({
      state: "rejected",
      errorKind: kind,
    });
    expect(data(await f.status(operationId)).status).toBe("rejected");
    expect((await f.run({ ...input, confirm: true, operationId })).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
  });
  it.each([{ readOnly: true }, { writable: false }])("honors existing write gates %j", async (options) => {
    const f = await fixture(undefined, options);
    const p = await f.run();
    expect(p.isError).toBeFalsy();
    expect((await f.run({ confirm: true, operationId: data(p).operationId })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("blocks duplicate dimensions and unpublished CAS types", async () => {
    const f = await fixture(saved());
    f.meta.keys[1] = f.meta.keys[0]!;
    expect((await f.run({ mode: "update", expectedDataVersion: "v1" })).isError).toBe(true);
    const g = await fixture(saved());
    g.meta.properties = g.meta.properties.filter((p) => p.name !== "DataVersion");
    expect((await g.run({ mode: "update", expectedDataVersion: "v1" })).isError).toBe(true);
    expect(g.writes).toEqual([]);
  });
  it("does not send a write without preview or after changing confirmation input", async () => {
    const f = await fixture();
    expect((await f.run({ confirm: true })).isError).toBe(true);
    const p = await f.run();
    const r = await f.run({
      confirm: true,
      operationId: data(p).operationId,
      accounts: { customerSettlement: ACCOUNT },
    });
    expect(r.isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("blocks missing resources and account references not found in configured chart", async () => {
    const f = await fixture();
    f.meta.properties = f.meta.properties.filter((p) => p.name !== resources[2]);
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
    const g = await fixture();
    expect((await g.run({ accounts: { supplierSettlement: OTHER } })).isError).toBe(true);
    expect(g.writes).toEqual([]);
  });
  it("reports uncertain register operations without document/comment reconciliation", async () => {
    const f = await fixture();
    const p = await f.run();
    f.corrupt();
    const operationId = data(p).operationId;
    await f.run({ confirm: true, operationId });
    const tools = (
      createServer({ db: () => f.conn } as never) as unknown as {
        _registeredTools: Record<
          string,
          { handler: (args: object, extra: object) => Promise<CallToolResult> }
        >;
      }
    )._registeredTools;
    const status = await tools["write.operation.status"]!.handler({ database: "default", operationId }, {});
    expect(data(status).status).toBe("unverifiable");
    expect((await f.conn.client.operationEntry(operationId as string))?.state).toBe("outcome_unknown");
  });
  it("sends If-Match through fetch and never retries a failed PATCH body", async () => {
    const f = await fixture();
    vi.mocked(f.conn.client.request).mockRestore();
    const fetchMock = vi.fn(async (_url: unknown, _init: unknown) => new Response("", { status: 412 }));
    vi.stubGlobal("fetch", fetchMock);
    try {
      await expect(
        f.conn.client.request("Synthetic(key='x')", "PATCH", { value: 1 }, { "If-Match": "v1" }),
      ).rejects.toBeInstanceOf(ODataError);
      expect(fetchMock).toHaveBeenCalledOnce();
      expect(fetchMock.mock.calls[0]![1]).toMatchObject({
        method: "PATCH",
        headers: { "If-Match": "v1" },
        body: JSON.stringify({ value: 1 }),
      });
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe("counterparty and group saved settlement scope", () => {
  it.each(["counterparty", "group"])("saves and reads %s all-organizations scope", async (scope) => {
    const f = await fixture();
    f.setSubject({ IsFolder: scope === "group" });
    const args = { scope, contractRef: undefined, organizationRef: EMPTY };
    const preview = await f.run(args, false, true);
    expect(preview.isError).toBeFalsy();
    expect(data(preview).payload).toMatchObject({
      Организация_Key: EMPTY,
      Контрагент_Key: CP,
      Договор_Key: EMPTY,
    });
    expect(
      (await f.run({ ...args, confirm: true, operationId: data(preview).operationId }, false, true)).isError,
    ).toBeFalsy();
    expect(data(await f.run(args, true, true)).found).toBe(true);
  });

  it.each(["counterparty", "group"])(
    "creates %s defaults with empty contract and exact organization/subject",
    async (scope) => {
      const f = await fixture();
      f.setSubject({ IsFolder: scope === "group" });
      const args = { scope, contractRef: undefined };
      const preview = await f.run(args, false, true);
      expect(preview.isError).toBeFalsy();
      expect(data(preview).payload).toEqual({ ...key, Договор_Key: EMPTY, [resources[2]!]: ACCOUNT });
      const result = await f.run(
        { ...args, confirm: true, operationId: data(preview).operationId },
        false,
        true,
      );
      expect(result.isError).toBeFalsy();
      expect(f.writes).toHaveLength(1);
    },
  );
  it.each([
    { scope: "group", contractRef: undefined },
    { scope: "counterparty", contractRef: CTR },
    { scope: "contract", contractRef: undefined },
    { scope: "counterparty", contractRef: undefined, counterpartyRef: EMPTY },
  ])("rejects inconsistent or wildcard scope %j", async (args) => {
    const f = await fixture();
    expect((await f.run(args, false, true)).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("requires generic tool explicit scope and keeps legacy tool contract-only", async () => {
    const f = await fixture();
    expect((await f.run({ contractRef: undefined }, false, true)).isError).toBe(true);
    expect((await f.run({ scope: "group", contractRef: undefined })).isError).toBe(true);
  });
  it.each(["subject", "organization"])("rejects deleted %s before write", async (which) => {
    const f = await fixture();
    if (which === "subject") f.setSubject({ DeletionMark: true });
    else f.setOrganization({ DeletionMark: true });
    expect((await f.run()).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it("updates group resources with CAS and preserves unspecified fields", async () => {
    const f = await fixture({ ...saved(), Договор_Key: EMPTY });
    f.setSubject({ IsFolder: true });
    const args = { scope: "group", contractRef: undefined, mode: "update", expectedDataVersion: "v1" };
    const preview = await f.run(args, false, true);
    const result = await f.run(
      { ...args, confirm: true, operationId: data(preview).operationId },
      false,
      true,
    );
    expect(result.isError).toBeFalsy();
    expect(f.writes[0]).toMatchObject({
      method: "PATCH",
      headers: { "If-Match": "v1" },
      body: { [resources[2]!]: ACCOUNT },
    });
    expect(data(result).record).toMatchObject({ СчетПрочихРасчетов_Key: OTHER });
  });
});

const versionlessSaved = () => {
  const { DataVersion: _version, ...record } = saved();
  return record;
};

describe("saved settings without metadata DataVersion", () => {
  it.each(["contract", "counterparty", "group"])(
    "updates %s settings sequentially without If-Match and preserves other resources",
    async (scope) => {
      const row = { ...versionlessSaved(), Договор_Key: scope === "contract" ? CTR : EMPTY };
      const f = await fixture(row, { versionless: true });
      f.setSubject({ IsFolder: scope === "group" });
      const args = {
        mode: "update",
        scope,
        contractRef: scope === "contract" ? CTR : undefined,
        singleWriter: true,
      };
      const preview = await f.run(args, false, true);
      expect(preview.isError).toBeFalsy();
      expect(data(preview)).toMatchObject({ concurrency: "singleWriter", headers: {} });
      expect(JSON.stringify(data(preview).notes)).toContain("DataVersion");
      expect(f.writes).toEqual([]);
      const result = await f.run(
        { ...args, confirm: true, operationId: data(preview).operationId },
        false,
        true,
      );
      expect(result.isError).toBeFalsy();
      expect(data(result)).toMatchObject({ verified: true, concurrency: "singleWriter" });
      expect(f.writes[0]).toMatchObject({ method: "PATCH", headers: {}, body: { [resources[2]!]: ACCOUNT } });
      expect(data(result).record).toMatchObject({ [resources[0]!]: OTHER, СчетПрочихРасчетов_Key: OTHER });
      const secondArgs = { ...args, accounts: { customerSettlement: ACCOUNT } };
      const second = await f.run(secondArgs, false, true);
      expect(second.isError).toBeFalsy();
      expect(
        (await f.run({ ...secondArgs, confirm: true, operationId: data(second).operationId }, false, true))
          .isError,
      ).toBeFalsy();
      expect(f.writes).toHaveLength(2);
      const replay = await f.run(
        { ...args, confirm: true, operationId: data(preview).operationId },
        false,
        true,
      );
      expect(data(replay).replayed).toBe(true);
      expect(f.writes).toHaveLength(2);
    },
  );
  it("requires explicit single-writer confirmation only for versionless updates", async () => {
    const f = await fixture(versionlessSaved(), { versionless: true });
    for (const singleWriter of [undefined, false]) {
      expect((await f.run({ mode: "update", singleWriter })).isError).toBe(true);
    }
    expect(f.writes).toEqual([]);
    const g = await fixture(undefined, { versionless: true });
    const preview = await g.run();
    expect(preview.isError).toBeFalsy();
    expect((await g.run({ confirm: true, operationId: data(preview).operationId })).isError).toBeFalsy();
  });
  it.each(["stale", "*", "v1\r\nInjected: true"])(
    "rejects invented version %j",
    async (expectedDataVersion) => {
      const f = await fixture(versionlessSaved(), { versionless: true });
      expect((await f.run({ mode: "update", singleWriter: true, expectedDataVersion })).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it.each([{ DataVersion: "v1" }, { DataVersion: null }])(
    "rejects unexpected returned version %j",
    async (extra) => {
      const f = await fixture({ ...versionlessSaved(), ...extra }, { versionless: true });
      expect((await f.run({ mode: "update", singleWriter: true })).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it("never downgrades published DataVersion even with singleWriter=true", async () => {
    const f = await fixture(saved());
    expect((await f.run({ mode: "update", singleWriter: true })).isError).toBe(true);
    f.setRow(versionlessSaved());
    expect((await f.run({ mode: "update", singleWriter: true, expectedDataVersion: "v1" })).isError).toBe(
      true,
    );
    expect(f.writes).toEqual([]);
    f.setRow(saved());
    const args = { mode: "update", singleWriter: true, expectedDataVersion: "v1" };
    const preview = await f.run(args);
    expect(data(preview).concurrency).toBe("dataVersion");
    expect(
      (await f.run({ ...args, confirm: true, operationId: data(preview).operationId })).isError,
    ).toBeFalsy();
    expect(f.writes[0]!.headers).toEqual({ "If-Match": "v1" });
  });
  it.each(["changed-resource", "changed-unpatched", "missing-resource", "deleted"])(
    "rejects stale preview: %s",
    async (change) => {
      const f = await fixture(versionlessSaved(), { versionless: true });
      const args = { mode: "update", singleWriter: true };
      const preview = await f.run(args);
      expect(preview.isError).toBeFalsy();
      const changed: ODataEntity = { ...versionlessSaved() };
      if (change === "changed-resource") changed[resources[2]!] = ACCOUNT;
      if (change === "changed-unpatched") changed.СчетПрочихРасчетов_Key = ACCOUNT;
      if (change === "missing-resource") delete changed.СчетПрочихРасчетов_Key;
      f.setRow(change === "deleted" ? undefined : changed);
      expect((await f.run({ ...args, confirm: true, operationId: data(preview).operationId })).isError).toBe(
        true,
      );
      expect(f.writes).toEqual([]);
    },
  );
  it("rejects update of an absent record without creating it", async () => {
    const f = await fixture(undefined, { versionless: true });
    expect((await f.run({ mode: "update", singleWriter: true })).isError).toBe(true);
    expect(f.writes).toEqual([]);
  });
  it.each(["write", "readback"])("holds unknown %s outcomes and prevents replay", async (stage) => {
    const f = await fixture(versionlessSaved(), { versionless: true });
    const args = { mode: "update", singleWriter: true };
    const preview = await f.run(args);
    expect(preview.isError).toBeFalsy();
    if (stage === "write") f.failWrite(new ODataError({ kind: "network", message: "Synthetic timeout" }));
    else f.failReadback(new ODataError({ kind: "auth", message: "Synthetic readback failure" }));
    const input = { ...args, confirm: true, operationId: data(preview).operationId };
    expect((await f.run(input)).isError).toBe(true);
    expect((await f.run(input)).isError).toBe(true);
    expect((await f.conn.client.operationEntry(input.operationId as string))?.state).toBe("outcome_unknown");
    expect(f.writes).toHaveLength(1);
  });
  it("rejects metadata DataVersion with unsupported type", async () => {
    const f = await fixture(saved());
    f.meta.properties.find((p) => p.name === "DataVersion")!.type = "Edm.Guid";
    expect((await f.run({ mode: "update", singleWriter: true, expectedDataVersion: "v1" })).isError).toBe(
      true,
    );
    expect(f.writes).toEqual([]);
  });
});

it("actual Belarus register schema requires honest singleWriter update without If-Match", async () => {
  const f = await fixture({ ...key, [resources[0]!]: OTHER, [resources[2]!]: OTHER }, { versionless: true });
  const schema = JSON.parse(
    readFileSync(new URL("./fixtures/metadata/belarus-2.1.json", import.meta.url), "utf8"),
  ) as { entities: Array<{ entitySet: string; keys: string[]; properties: EntityMeta["properties"] }> };
  const actual = schema.entities.find((e) => e.entitySet === REG)!;
  f.entities.set(REG, { ...f.meta, ...actual });
  const blocked = await f.run({ mode: "update" });
  expect(blocked.isError).toBe(true);
  const p = await f.run({ mode: "update", singleWriter: true });
  expect(p.isError, JSON.stringify(p)).toBeFalsy();
  expect(data(p).concurrency).toBe("singleWriter");
  expect(
    (await f.run({ mode: "update", singleWriter: true, confirm: true, operationId: data(p).operationId }))
      .isError,
  ).toBeFalsy();
  expect(f.writes[0]?.headers).toEqual({});
});

it.each(["create", "update"] as const)(
  "reconciles lost register %s by exact compound key and normalized values without resend",
  async (mode) => {
    const before = { ...saved(), СчетПрочихРасчетов_Key: "abcdefab-cdef-4abc-8def-abcdefabcdef" };
    const f = await fixture(mode === "update" ? before : undefined);
    const args = { mode, ...(mode === "update" ? { expectedDataVersion: "v1" } : {}) };
    const p = await f.run(args);
    const operationId = String(data(p).operationId);
    f.failWrite(new ODataError({ kind: "timeout", message: "Synthetic lost response" }));
    expect((await f.run({ ...args, confirm: true, operationId })).isError).toBe(true);
    const expected = {
      ...(mode === "update"
        ? { ...before, СчетПрочихРасчетов_Key: before.СчетПрочихРасчетов_Key.toUpperCase() }
        : key),
      [resources[2]!]: ACCOUNT,
    };
    f.setRow(expected);
    const status = await f.status(operationId);
    expect(data(status).status).toBe("found_reconciled");
    expect(data(status).recordPath).toContain("Организация_Key=guid'");
    expect((await f.conn.client.operationEntry(operationId))?.state).toBe("succeeded");
    expect((await f.run({ ...args, confirm: true, operationId })).isError).toBeFalsy();
    expect(f.writes).toHaveLength(1);
  },
);
it.each(["missing", "different"])("keeps %s register result held after a lost response", async (reason) => {
  const f = await fixture();
  const p = await f.run();
  const operationId = String(data(p).operationId);
  f.failWrite(new ODataError({ kind: "timeout", message: "Synthetic lost response" }));
  await f.run({ confirm: true, operationId });
  if (reason === "different") f.setRow({ ...key, [resources[2]!]: OTHER });
  const status = await f.status(operationId);
  expect(data(status).status).toBe("unverifiable");
  expect((await f.conn.client.operationEntry(operationId))?.state).toBe("outcome_unknown");
  expect((await f.run({ confirm: true, operationId })).isError).toBe(true);
  expect(f.writes).toHaveLength(1);
});
