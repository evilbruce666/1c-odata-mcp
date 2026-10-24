import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { sameStoredValue } from "../src/tools/write.js";
import { fake1C } from "./helpers/fake-1c.js";

// #60: изменение — успех только по фактически сохранённым значениям, а не по ответу на PATCH.
const BANK_IN = "Document_ПоступлениеНаРасчетныйСчет";
const REF = "33333333-3333-4333-8333-333333333333";
const CONTRACT_OLD = "aaaaaaaa-0000-4000-8000-000000000001";
const CONTRACT_NEW = "aaaaaaaa-0000-4000-8000-000000000002";
const ACCOUNT_NEW = "bbbbbbbb-0000-4000-8000-000000000002";

function tool(ctx: ReturnType<typeof fake1C>["ctx"]) {
  const tools = (
    createServer(ctx) as unknown as {
      _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  return (fields: Record<string, unknown>, entitySet = BANK_IN) =>
    tools["write.entity.update_entity"]!.handler({ entitySet, ref: REF, fields, confirm: true }, {});
}
const writable = () => fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
const text = (r: CallToolResult) => JSON.stringify(r.content);

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("update_entity: сверка сохранённых значений после PATCH (#60)", () => {
  it("все поля сохранились — updated и verified", async () => {
    const f = writable();
    const r = await tool(f.ctx)({ ДоговорКонтрагента_Key: CONTRACT_NEW, Комментарий: "новый" });
    expect(r.isError).toBeFalsy();
    expect(r.structuredContent).toMatchObject({ updated: true, verified: true });
    expect(f.violations).toEqual([]);
  });

  it("1С сохранила счёт, но оставила договор — ошибка с именем поля и подсказкой для банка", async () => {
    const f = writable();
    const getEntity = f.conn.client.getEntity.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getEntity").mockImplementation(async (path: string) => ({
      ...(await getEntity(path)),
      ДоговорКонтрагента_Key: CONTRACT_OLD,
    }));
    const patch = vi.spyOn(f.conn.client, "patch");
    const r = await tool(f.ctx)({
      ДоговорКонтрагента_Key: CONTRACT_NEW,
      СчетУчетаРасчетовСКонтрагентом_Key: ACCOUNT_NEW,
    });
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("ДоговорКонтрагента_Key");
    expect(text(r)).not.toContain("СчетУчетаРасчетовСКонтрагентом_Key,");
    expect(text(r)).toContain("остальные сохранены");
    expect(text(r)).toContain("apply_bank_settlement_settings");
    expect(patch).toHaveBeenCalledOnce(); // без автоматического повтора
  });

  it("не банковский документ — без банковской подсказки", async () => {
    const f = writable();
    const getEntity = f.conn.client.getEntity.bind(f.conn.client);
    vi.spyOn(f.conn.client, "getEntity").mockImplementation(async (path: string) => ({
      ...(await getEntity(path)),
      Комментарий: "старый",
    }));
    const r = await tool(f.ctx)({ Комментарий: "новый" }, "Catalog_Контрагенты");
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("Комментарий");
    expect(text(r)).not.toContain("apply_bank_settlement_settings");
  });

  it("табличная часть не сверяется — updated без verified, с пометкой", async () => {
    const f = writable();
    const r = await tool(f.ctx)({ Комментарий: "новый", РасшифровкаПлатежа: [] });
    expect(r.isError).toBeFalsy();
    expect(r.structuredContent).toMatchObject({ updated: true, verified: false });
    expect(String((r.structuredContent as { readbackNote: string }).readbackNote)).toContain(
      "РасшифровкаПлатежа",
    );
  });

  it("перечитать не удалось — просьба сверить в 1С", async () => {
    const f = writable();
    vi.spyOn(f.conn.client, "getEntity").mockRejectedValue(new Error("сеть"));
    const r = await tool(f.ctx)({ Комментарий: "новый" });
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("Сверьте поля в 1С");
  });
});

describe("sameStoredValue — как 1С возвращает сохранённое", () => {
  it.each([
    ["AAAAAAAA-0000-4000-8000-000000000001", "aaaaaaaa-0000-4000-8000-000000000001", "Edm.Guid", true],
    [null, "00000000-0000-0000-0000-000000000000", "Edm.Guid", true],
    ["2026-10-24", "2026-10-24T00:00:00", "Edm.DateTime", true],
    ["2026-10-24", "2026-10-25T00:00:00", "Edm.DateTime", false],
    [100, 100.0, "Edm.Double", true],
    ["100", 100, "Edm.Decimal", true],
    [100, 99.99, "Edm.Double", false],
    [true, false, "Edm.Boolean", false],
    ["Текст", "Текст  ", "Edm.String", true],
    ["", null, "Edm.String", true],
    ["a", "", "Edm.String", false],
  ])("%j ↔ %j (%s) → %s", (sent, stored, type, same) =>
    expect(sameStoredValue(sent, stored, type)).toBe(same),
  );
});
