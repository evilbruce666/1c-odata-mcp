import { describe, expect, it, vi } from "vitest";
import { counterpartyAncestors, MAX_GROUP_DEPTH } from "../src/odata/counterparty-hierarchy.js";
import { settlementAccountsFor } from "../src/odata/settlement-accounts.js";
const E = "00000000-0000-0000-0000-000000000000";
const CP = "11111111-1111-4111-8111-111111111111";
const G = "22222222-2222-4222-8222-222222222222";
const FAR = "33333333-3333-4333-8333-333333333333";
const REG = "InformationRegister_СчетаУчетаРасчетовСКонтрагентами";
const CAT = "Catalog_Контрагенты";
function fixture() {
  const rows = new Map<string, Record<string, unknown>>([
    [CP, { Ref_Key: CP, IsFolder: false, DeletionMark: false, Parent_Key: G }],
    [G, { Ref_Key: G, IsFolder: true, DeletionMark: false, Parent_Key: FAR }],
    [FAR, { Ref_Key: FAR, IsFolder: true, DeletionMark: false, Parent_Key: E }],
  ]);
  const properties = ["Ref_Key", "Parent_Key", "IsFolder", "DeletionMark"].map((name) => ({
    name,
    type: name.endsWith("_Key") ? "Edm.Guid" : "Edm.Boolean",
  }));
  const entities = new Map([[CAT, { properties }]]);
  const contract = {
    Owner_Key: CP,
    Организация_Key: CP,
    ВидДоговора: "СПокупателем",
    ВалютаВзаиморасчетов_Key: CP,
    РасчетыВУсловныхЕдиницах: false,
  };
  const getEntity = vi.fn(async (path: string) => {
    if (path.startsWith(CAT)) {
      const ref = path.match(/guid'([^']+)'/)![1]!;
      if (!rows.has(ref)) throw new Error("missing parent");
      return rows.get(ref)!;
    }
    if (path.startsWith("Catalog_Валюты")) return { Code: "643" };
    return contract;
  });
  const getCollection = vi.fn(async () => ({ value: [] as object[] }));
  const conn = {
    getMetadata: async () => ({ entities }),
    available: async () => new Set([CAT, REG, "Catalog_ДоговорыКонтрагентов", "Catalog_Валюты"]),
    behavior: { pageSize: 100 },
    client: { getEntity, getCollection },
  } as never;
  return { rows, properties, conn, contract, getEntity, getCollection };
}
describe("current bounded counterparty ancestors", () => {
  it("reads current parent each time and never caches old membership", async () => {
    const f = fixture();
    expect((await counterpartyAncestors(f.conn, CP)).groupRefs).toEqual([G, FAR]);
    f.rows.get(CP)!.Parent_Key = FAR;
    expect((await counterpartyAncestors(f.conn, CP)).groupRefs).toEqual([FAR]);
  });
  it.each(["cycle", "missing", "deleted", "individual parent", "unknown parent", "schema"])(
    "blocks %s instead of wrong/default fallback",
    async (reason) => {
      const f = fixture();
      if (reason === "cycle") f.rows.get(FAR)!.Parent_Key = CP;
      if (reason === "missing") f.rows.delete(G);
      if (reason === "deleted") f.rows.get(G)!.DeletionMark = true;
      if (reason === "individual parent") f.rows.get(G)!.IsFolder = false;
      if (reason === "unknown parent") delete f.rows.get(CP)!.Parent_Key;
      if (reason === "schema") f.properties.pop();
      await expect(
        settlementAccountsFor(f.conn, { contractRef: CP }, { contract: f.contract }),
      ).rejects.toThrow();
      expect(f.getCollection).not.toHaveBeenCalled();
    },
  );
  it("blocks over-depth hierarchy", async () => {
    const f = fixture();
    let previous = CP;
    for (let i = 0; i <= MAX_GROUP_DEPTH; i++) {
      const ref = `${String(i).padStart(8, "0")}-aaaa-4aaa-8aaa-aaaaaaaaaaaa`;
      f.rows.get(previous)!.Parent_Key = ref;
      f.rows.set(ref, { Ref_Key: ref, Parent_Key: E, IsFolder: true, DeletionMark: false });
      previous = ref;
    }
    await expect(counterpartyAncestors(f.conn, CP)).rejects.toThrow("глубже");
  });
  it("queries only current ancestors and refuses truncated register rows", async () => {
    const f = fixture();
    f.getCollection.mockResolvedValue({ value: Array.from({ length: 500 }, () => ({})) });
    await expect(
      settlementAccountsFor(f.conn, { contractRef: CP }, { contract: f.contract }),
    ).rejects.toThrow("усечён");
    const query = decodeURIComponent(f.getCollection.mock.calls[0]![0] ?? "");
    expect(query).toContain(G);
    expect(query).toContain(FAR);
    expect(f.getCollection).toHaveBeenCalledTimes(1);
  });
});
