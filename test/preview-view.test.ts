import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { buildPreviewView, fieldLabel, tryPreviewView } from "../src/tools/preview-view.js";
import { createServer } from "../src/mcp/server.js";
import { bp30Metadata } from "./helpers/metadata-snapshot.js";
import { fake1C } from "./helpers/fake-1c.js";

const SALE = "Document_РеализацияТоваровУслуг";
const CP = "11111111-1111-1111-1111-111111111111";
const CONTRACT = "22222222-2222-2222-2222-222222222222";
const ITEM = "33333333-3333-3333-3333-333333333333";
const ACCOUNT = "44444444-4444-4444-4444-444444444444";
const ORG = "55555555-5555-5555-5555-555555555555";
const EMPTY = "00000000-0000-0000-0000-000000000000";

const NAMES: Record<string, Record<string, unknown>> = {
  [CP]: { Description: "ООО Ромашка" },
  [CONTRACT]: { Description: "Основной договор" },
  [ITEM]: { Description: "Консультация" },
  [ACCOUNT]: { Code: "90.01.1", Description: "Выручка" },
  [ORG]: { Description: "ООО Свои" },
};

/** Подключение на снимке схемы БП 3.0; названия — из NAMES по Ref_Key в отборе. */
function conn(failing: string[] = []) {
  const getCollection = vi.fn(async (path: string) => {
    const set = /^[^?]+/.exec(path)![0];
    if (failing.includes(set)) throw new Error("нет прав");
    const refs = [...decodeURIComponent(path).matchAll(/guid'([0-9a-f-]{36})'/g)].map((m) => m[1]!);
    return { value: refs.filter((r) => NAMES[r]).map((r) => ({ Ref_Key: r, ...NAMES[r] })) };
  });
  const meta = bp30Metadata();
  return {
    getCollection,
    conn: {
      getMetadata: async () => meta,
      available: async () => new Set(meta.entities.keys()),
      client: { getCollection },
    } as never,
  };
}

describe("читаемый предпросмотр записи", () => {
  it("подписи полей по-русски", () => {
    expect(fieldLabel("ДоговорКонтрагента_Key")).toBe("Договор контрагента");
    expect(fieldLabel("СчетУчетаНДСПоРеализации_Key")).toBe("Счет учета НДС по реализации");
  });

  it("названия вместо GUID, пустые и устаревшие «Удалить…» скрыты, подписи и даты по-русски", async () => {
    const { conn: c } = conn();
    const view = await buildPreviewView(c, SALE, {
      Date: "2026-10-24T05:55:43",
      Posted: false,
      РучнаяКорректировка: false,
      СуммаВключаетНДС: false,
      УдалитьУчитыватьНДС: true,
      Комментарий: "",
      АдресДоставкиЗначение: '{"version": 4, "value": "г. Златоуст"}',
      ДатаВходящегоДокумента: "0001-01-01T00:00:00",
      Контрагент_Key: CP,
      ДоговорКонтрагента_Key: CONTRACT,
      Склад_Key: EMPTY,
      СуммаДокумента: 12000,
      Услуги: [{ LineNumber: "1", Номенклатура_Key: ITEM, Сумма: 12000, СчетДоходов_Key: ACCOUNT }],
    });
    expect(view.title).toBe("Реализация товаров услуг");
    expect(view.fields).toEqual({
      Дата: "2026-10-24 05:55",
      "Сумма включает НДС": "нет",
      Контрагент: "ООО Ромашка",
      "Договор контрагента": "Основной договор",
      "Сумма документа": 12000,
    });
    expect(view.tables).toEqual({
      Услуги: [{ Номенклатура: "Консультация", Сумма: 12000, "Счет доходов": "90.01.1 Выручка" }],
    });
  });

  it("составное поле — цель по *_Type", async () => {
    const { conn: c } = conn();
    const view = await buildPreviewView(c, "Document_ПоступлениеНаРасчетныйСчет", {
      Контрагент: CP,
      Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
    });
    expect(view.fields).toEqual({ Контрагент: "ООО Ромашка" });
  });

  it("не больше 4 справочников, по важности; остальные ссылки — GUID", async () => {
    const { conn: c, getCollection } = conn();
    const view = await buildPreviewView(c, SALE, {
      Организация_Key: ORG,
      Контрагент_Key: CP,
      ДоговорКонтрагента_Key: CONTRACT,
      Товары: [{ Номенклатура_Key: ITEM, СчетУчета_Key: ACCOUNT }],
    });
    const sets = getCollection.mock.calls.map((call) => /^[^?]+/.exec(call[0])![0]);
    expect(sets).toEqual([
      "Catalog_Контрагенты",
      "Catalog_ДоговорыКонтрагентов",
      "Catalog_Номенклатура",
      "ChartOfAccounts_Хозрасчетный",
    ]);
    expect(view.fields["Организация"]).toBe(ORG);
  });

  it("изменение — «было → станет», очистка ссылки видна", async () => {
    const { conn: c } = conn();
    const view = await buildPreviewView(
      c,
      SALE,
      { ДоговорКонтрагента_Key: EMPTY, Комментарий: "новый" },
      { ДоговорКонтрагента_Key: CONTRACT, Комментарий: "старый" },
    );
    expect(view.changes).toEqual([
      { field: "Договор контрагента", from: "Основной договор", to: "—" },
      { field: "Комментарий", from: "старый", to: "новый" },
    ]);
  });

  it("названия не прочитались — GUID; метаданные не прочитались — предпросмотр без вида", async () => {
    const { conn: c } = conn(["Catalog_Контрагенты"]);
    expect((await buildPreviewView(c, SALE, { Контрагент_Key: CP })).fields).toEqual({ Контрагент: CP });
    const broken = {
      getMetadata: async () => {
        throw new Error("сбой");
      },
    } as never;
    expect(await tryPreviewView(broken, SALE, { Контрагент_Key: CP })).toEqual({});
  });
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("предпросмотр инструментов записи несёт view, payload прежний", () => {
  const tools = (ctx: ReturnType<typeof fake1C>["ctx"]) =>
    (
      createServer(ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools;

  it("create_invoice: view с заголовком и строками рядом с payload", async () => {
    const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const r = await tools(f.ctx)["write.sales.create_invoice"]!.handler(
      {
        counterpartyRef: CP,
        sumIncludesVat: true,
        lines: [{ nomenclatureRef: ITEM, quantity: 1, price: 100 }],
        confirm: false,
      },
      {},
    );
    expect(r.isError).toBeFalsy();
    const sc = r.structuredContent as { view: { title: string; tables: object }; payload: object };
    expect(sc.view.title).toBe("Счет на оплату покупателю");
    expect(Object.keys(sc.view.tables)).toEqual(["Товары"]);
    expect(sc.payload).toHaveProperty("Контрагент_Key", CP);
    expect(f.violations).toEqual([]);
  });

  it("update_entity: view.changes по текущим значениям", async () => {
    const f = fake1C();
    const r = await tools(f.ctx)["write.entity.update_entity"]!.handler(
      { entitySet: "Catalog_Контрагенты", ref: CP, fields: { Комментарий: "новый" }, confirm: false },
      {},
    );
    const sc = r.structuredContent as { view: { changes: unknown[] }; fields: object };
    expect(sc.view.changes).toEqual([{ field: "Комментарий", from: "—", to: "новый" }]);
    expect(sc.fields).toEqual({ Комментарий: "новый" });
    expect(f.violations).toEqual([]);
  });
});
