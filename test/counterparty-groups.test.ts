import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Connection } from "../src/context.js";
import { parseConfig } from "../src/config/env.js";
import { createServer } from "../src/mcp/server.js";
import { ODataError } from "../src/odata/errors.js";
import { counterpartyAncestors } from "../src/odata/counterparty-hierarchy.js";
const CP = "11111111-1111-4111-8111-111111111111";
const G = "22222222-2222-4222-8222-222222222222";
const E = "00000000-0000-0000-0000-000000000000";
const CAT = "Catalog_Контрагенты";
const dirs: string[] = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(dirs.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});
async function fixture() {
  const dir = await mkdtemp(join(tmpdir(), "group-tools-"));
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
    { ...cfg.behavior, writeJournalDir: dir },
  );
  const rows = new Map<string, Record<string, unknown>>([
    [
      CP,
      {
        Ref_Key: CP,
        Parent_Key: E,
        IsFolder: false,
        DeletionMark: false,
        DataVersion: "v1",
        Description: "Synthetic",
        Extra: "preserve",
      },
    ],
    [G, { Ref_Key: G, Parent_Key: E, IsFolder: true, DeletionMark: false, DataVersion: "g1" }],
  ]);
  const properties = [
    "Ref_Key",
    "Parent_Key",
    "IsFolder",
    "DeletionMark",
    "DataVersion",
    "Description",
    "Extra",
  ].map((name) => ({
    name,
    type: name.endsWith("_Key")
      ? "Edm.Guid"
      : ["IsFolder", "DeletionMark"].includes(name)
        ? "Edm.Boolean"
        : "Edm.String",
  }));
  vi.spyOn(conn, "getMetadata").mockResolvedValue({
    entities: new Map([[CAT, { entitySet: CAT, keys: ["Ref_Key"], properties }]]),
  } as never);
  const writes: { method: string; body: Record<string, unknown>; headers: unknown }[] = [];
  let corrupt = false;
  let writeError = false;
  vi.spyOn(conn.client, "request").mockImplementation(async (path, method = "GET", payload, headers) => {
    const ref = path.match(/guid'([^']+)'/)?.[1];
    if (method === "GET") {
      if (!ref || !rows.has(ref)) throw new ODataError({ kind: "not_found", message: "Missing" });
      return { ...rows.get(ref) } as never;
    }
    const body = payload as Record<string, unknown>;
    writes.push({ method, body, headers });
    if (writeError) throw new ODataError({ kind: "unknown", message: "Unknown effect" });
    const target = ref ?? String(body.Ref_Key);
    rows.set(target, {
      ...rows.get(target),
      ...body,
      DataVersion: "v2",
      DeletionMark: false,
      ...(corrupt ? { Extra: "lost" } : {}),
    });
    return {} as never;
  });
  const server = createServer({ db: () => conn } as never);
  const tools = (
    server as unknown as {
      _registeredTools: Record<
        string,
        {
          handler: (
            args: object,
            extra: object,
          ) => Promise<{ isError?: boolean; structuredContent?: Record<string, unknown> }>;
        }
      >;
    }
  )._registeredTools;
  const run = (name: string, args: object) => {
    const tool = tools[`write.counterparty.${name}`];
    expect(tool).toBeDefined();
    return tool!.handler({ database: "default", confirm: false, ...args }, {});
  };
  return {
    run,
    server,
    conn,
    writes,
    rows,
    properties,
    corrupt: () => {
      corrupt = true;
    },
    failWrite: () => {
      writeError = true;
    },
  };
}
describe("explicit counterparty group operations", () => {
  it.each(["group", "counterparty"])(
    "rejects relocating a group before preview even with subjectKind=%s through MCP",
    async (subjectKind) => {
      const f = await fixture();
      f.rows.get(CP)!.Parent_Key = G;
      let parent = E;
      for (let i = 0; i < 32; i++) {
        const ref = `${String(i).padStart(8, "0")}-aaaa-4aaa-8aaa-aaaaaaaaaaaa`;
        f.rows.set(ref, { Ref_Key: ref, Parent_Key: parent, IsFolder: true, DeletionMark: false });
        parent = ref;
      }
      const prepare = vi.spyOn(f.conn.client, "prepareCreate");
      const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair();
      const client = new Client({ name: "group-regression", version: "1" });
      await f.server.connect(serverTransport);
      await client.connect(clientTransport);
      try {
        const result = await client.callTool({
          name: "write.counterparty.set_group",
          arguments: { ref: G, subjectKind, parentRef: parent, expectedDataVersion: "g1" },
        });
        expect(result.isError).toBe(true);
        expect(prepare).not.toHaveBeenCalled();
        expect(f.writes).toEqual([]);
        expect(f.rows.get(G)!.Parent_Key).toBe(E);
        expect((await counterpartyAncestors(f.conn, CP)).groupRefs).toEqual([G]);
      } finally {
        await client.close();
        await f.server.close();
      }
    },
  );
  it.each([31, 32])("checks new child depth with parent having %i ancestors", async (count) => {
    const f = await fixture();
    let previous = G;
    for (let i = 0; i < count; i++) {
      const ref = `${String(i).padStart(8, "0")}-aaaa-4aaa-8aaa-aaaaaaaaaaaa`;
      f.rows.get(previous)!.Parent_Key = ref;
      f.rows.set(ref, { Ref_Key: ref, Parent_Key: E, IsFolder: true, DeletionMark: false });
      previous = ref;
    }
    const created = await f.run("create_group", { name: "Synthetic", parentRef: G });
    const moved = await f.run("set_group", {
      ref: CP,
      subjectKind: "counterparty",
      parentRef: G,
      expectedDataVersion: "v1",
    });
    expect(!!created.isError).toBe(count === 32);
    expect(!!moved.isError).toBe(count === 32);
    expect(f.writes).toEqual([]);
  });
  it("creates a validated group under explicit parent with preview/confirm/readback/replay", async () => {
    const f = await fixture();
    const args = { name: "Synthetic group", parentRef: G };
    const p = await f.run("create_group", args);
    expect(p.isError).toBeFalsy();
    expect(f.writes).toEqual([]);
    const r = await f.run("create_group", {
      ...args,
      confirm: true,
      operationId: p.structuredContent!.operationId,
    });
    expect(r.isError).toBeFalsy();
    expect(f.writes[0]!.body).toMatchObject({
      Description: "Synthetic group",
      Parent_Key: G,
      IsFolder: true,
    });
    const replay = await f.run("create_group", {
      ...args,
      confirm: true,
      operationId: p.structuredContent!.operationId,
    });
    expect(replay.structuredContent!.replayed).toBe(true);
    expect(f.writes).toHaveLength(1);
  });
  it("moves only Parent_Key with exact version and preserves all other fields", async () => {
    const f = await fixture();
    const args = { ref: CP, subjectKind: "counterparty", parentRef: G, expectedDataVersion: "v1" };
    const p = await f.run("set_group", args);
    const r = await f.run("set_group", {
      ...args,
      confirm: true,
      operationId: p.structuredContent!.operationId,
    });
    expect(r.isError).toBeFalsy();
    expect(f.writes).toEqual([{ method: "PATCH", body: { Parent_Key: G }, headers: { "If-Match": "v1" } }]);
    expect(f.rows.get(CP)!.Extra).toBe("preserve");
  });
  it.each(["stale", "individual parent", "deleted parent", "cycle", "missing version"])(
    "blocks %s",
    async (reason) => {
      const f = await fixture();
      const args: Record<string, unknown> = {
        ref: CP,
        subjectKind: "counterparty",
        parentRef: G,
        expectedDataVersion: "v1",
      };
      if (reason === "stale") args.expectedDataVersion = "old";
      if (reason === "missing version") delete args.expectedDataVersion;
      if (reason === "individual parent") args.parentRef = CP;
      if (reason === "deleted parent") f.rows.get(G)!.DeletionMark = true;
      if (reason === "cycle") {
        args.ref = G;
        args.subjectKind = "group";
        args.expectedDataVersion = "g1";
      }
      expect((await f.run("set_group", args)).isError).toBe(true);
      expect(f.writes).toEqual([]);
    },
  );
  it.each(["readback", "unknown"])("holds %s failure without second mutation", async (reason) => {
    const f = await fixture();
    const args = { ref: CP, subjectKind: "counterparty", parentRef: G, expectedDataVersion: "v1" };
    const p = await f.run("set_group", args);
    if (reason === "readback") f.corrupt();
    else f.failWrite();
    const confirmed = { ...args, confirm: true, operationId: p.structuredContent!.operationId };
    expect((await f.run("set_group", confirmed)).isError).toBe(true);
    expect((await f.run("set_group", confirmed)).isError).toBe(true);
    expect(f.writes).toHaveLength(1);
  });
});

it("lost create_group reply reconciles its exact assigned ref without marker search", async () => {
  const f = await fixture();
  const p = await f.run("create_group", { name: "Synthetic", parentRef: E });
  const operationId = String(p.structuredContent?.operationId);
  f.failWrite();
  expect(
    (await f.run("create_group", { name: "Synthetic", parentRef: E, confirm: true, operationId })).isError,
  ).toBe(true);
  f.rows.set(operationId, { Ref_Key: operationId, IsFolder: true, DeletionMark: false, Parent_Key: E });
  const tools = (
    f.server as unknown as {
      _registeredTools: Record<
        string,
        { handler: (args: object, extra: object) => Promise<{ structuredContent?: Record<string, unknown> }> }
      >;
    }
  )._registeredTools;
  const status = await tools["write.operation.status"]!.handler({ operationId }, {});
  expect(status.structuredContent?.status).toBe("found_reconciled");
  expect(f.writes).toHaveLength(1);
});
