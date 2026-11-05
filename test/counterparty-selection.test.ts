import { describe, expect, it, vi } from "vitest";
import { createServer } from "../src/mcp/server.js";
const G = "22222222-2222-4222-8222-222222222222";
const P = "33333333-3333-4333-8333-333333333333";
function fixture() {
  const row = {
    Ref_Key: G,
    Description: "Synthetic group",
    IsFolder: true,
    Parent_Key: P,
    DeletionMark: false,
    DataVersion: "v1",
  };
  const getCollection = vi.fn(async () => ({ value: [row] }));
  const conn = {
    cfg: { name: "test" },
    available: async () => new Set(["Catalog_Контрагенты"]),
    behavior: { pageSize: 100, maxRows: 100 },
    client: { getCollection, getEntity: async () => row },
  };
  const tools = (
    createServer({ db: () => conn, databases: () => [] } as never) as unknown as {
      _registeredTools: Record<
        string,
        { handler: (args: object, extra: object) => Promise<{ structuredContent?: Record<string, unknown> }> }
      >;
    }
  )._registeredTools;
  return { tools, getCollection };
}
describe("counterparty/group selection", () => {
  it("finds groups explicitly in the same catalog and exposes parent/version", async () => {
    const f = fixture();
    const result = await f.tools["read.counterparty.find_counterparty"]!.handler(
      { query: "Synthetic", subjectKind: "group", limit: 20 },
      {},
    );
    expect(result.structuredContent!.rows).toMatchObject([
      { ref: G, isFolder: true, parentRef: P, dataVersion: "v1", deletionMark: false },
    ]);
    expect(decodeURIComponent(f.getCollection.mock.calls[0]![0] ?? "")).toContain("IsFolder eq true");
  });
  it("card exposes membership and version for safe set_group", async () => {
    const f = fixture();
    const result = await f.tools["read.counterparty.get_counterparty"]!.handler({ ref: G }, {});
    expect(result.structuredContent).toMatchObject({ parentRef: P, dataVersion: "v1", isFolder: true });
  });
});

const CP = "11111111-1111-4111-8111-111111111111";
const CAT = "Catalog_Контрагенты";
const CONTRACTS = "Catalog_ДоговорыКонтрагентов";
function combinedFixture() {
  const counterparty = {
    Ref_Key: CP,
    Description: "Synthetic counterparty",
    НаименованиеПолное: "Synthetic counterparty full name",
    IsFolder: false,
    Parent_Key: G,
    DeletionMark: false,
    DataVersion: "v2",
  };
  const group = { Ref_Key: G, Description: "Synthetic group", IsFolder: true, Parent_Key: P };
  const getCollection = vi.fn(async (path: string) => {
    if (path.startsWith(CONTRACTS)) return { value: [{ Owner_Key: CP, ВидДоговора: "СПоставщиком" }] };
    const query = decodeURIComponent(path);
    return {
      value: query.includes("IsFolder eq true")
        ? [group]
        : query.includes("IsFolder eq false")
          ? [counterparty]
          : [group, counterparty],
    };
  });
  const conn = {
    cfg: { name: "test" },
    available: async () => new Set([CAT, CONTRACTS]),
    behavior: { pageSize: 100, maxRows: 100 },
    client: { getCollection },
  };
  const tools = (
    createServer({ db: () => conn, databases: () => [] } as never) as unknown as {
      _registeredTools: Record<
        string,
        { handler: (args: object, extra: object) => Promise<{ structuredContent?: Record<string, unknown> }> }
      >;
    }
  )._registeredTools;
  return { tools, getCollection };
}
describe("combined counterparty selection and contract roles", () => {
  it("preserves full name, membership and version while enriching individual roles", async () => {
    const f = combinedFixture();
    const result = await f.tools["read.counterparty.find_counterparty"]!.handler(
      { query: "Synthetic", limit: 20 },
      {},
    );
    expect(result.structuredContent!.rows).toMatchObject([
      {
        ref: CP,
        fullName: "Synthetic counterparty full name",
        parentRef: G,
        dataVersion: "v2",
        deletionMark: false,
        roles: ["поставщик"],
        contracts: 1,
      },
    ]);
    expect(decodeURIComponent(f.getCollection.mock.calls[0]![0])).toContain("НаименованиеПолное");
    expect(decodeURIComponent(f.getCollection.mock.calls[1]![0])).toContain("DeletionMark eq false");
  });
  it("returns groups without querying contracts owned by groups", async () => {
    const f = combinedFixture();
    const result = await f.tools["read.counterparty.find_counterparty"]!.handler(
      { query: "Synthetic", subjectKind: "group", limit: 20 },
      {},
    );
    expect(result.structuredContent!.rows).toMatchObject([{ ref: G, isFolder: true, parentRef: P }]);
    expect(f.getCollection).toHaveBeenCalledTimes(1);
  });
  it("all returns both kinds and restricts contract enrichment to individual refs", async () => {
    const f = combinedFixture();
    const result = await f.tools["read.counterparty.find_counterparty"]!.handler(
      { query: "Synthetic", subjectKind: "all", limit: 20 },
      {},
    );
    expect(result.structuredContent!.rows).toMatchObject([
      { ref: G, isFolder: true },
      { ref: CP, roles: ["поставщик"], contracts: 1 },
    ]);
    const contractQuery = decodeURIComponent(f.getCollection.mock.calls[1]![0]);
    expect(contractQuery).toContain(CP);
    expect(contractQuery).not.toContain(G);
  });
});
