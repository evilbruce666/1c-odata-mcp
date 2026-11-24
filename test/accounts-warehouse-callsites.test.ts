import { describe, expect, it } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";
import { SNAPSHOT_REF as id } from "./helpers/metadata-snapshot.js";

/**
 * Склад документа доходит до подбора счетов строк (регистр «Счета учёта номенклатуры») у каждого инструмента БП со
 * складом: при общей записи и записи этого склада строка получает счёт склада при любом порядке записей. Ревью
 * 501c7b6: инвентаризация склад не передавала, и юнит-тесты подбора этого не видели — они проверяют функцию, а не
 * места вызова.
 */

const E = "00000000-0000-0000-0000-000000000000";
const GENERAL = "aaaaaaaa-0000-4000-8000-000000000001";
const OWN = "aaaaaaaa-0000-4000-8000-000000000002";
const row = (warehouse: string, account: string) => ({
  Организация_Key: E,
  Номенклатура_Key: E,
  ВидНоменклатуры_Key: E,
  Склад_Key: warehouse,
  ТипСклада: "",
  СчетУчета_Key: account,
  СчетУчетаНДСПоПриобретеннымЦенностям_Key: E,
  СчетУчетаНДСПоРеализации_Key: E,
  СчетДоходовОтРеализации_Key: E,
  СчетРасходовОтРеализации_Key: E,
});

type Tools = Record<
  string,
  {
    handler: (
      a: unknown,
      e: unknown,
    ) => Promise<{
      structuredContent?: Record<string, unknown>;
      isError?: boolean;
      content?: { text?: string }[];
    }>;
    inputSchema: { parse: (x: unknown) => unknown };
  }
>;

const line = { nomenclatureRef: id, quantity: 1, price: 100 };
const CASES: Array<[string, Record<string, unknown>]> = [
  ["write.sales.create_shipment", { counterpartyRef: id, contractRef: id, warehouse: id, lines: [line] }],
  ["write.purchase.create_purchase", { counterpartyRef: id, contractRef: id, warehouse: id, lines: [line] }],
  ["write.warehouse.create_return_from_customer", { counterpartyRef: id, warehouse: id, lines: [line] }],
  ["write.warehouse.create_return_to_supplier", { counterpartyRef: id, warehouse: id, lines: [line] }],
  ["write.warehouse.create_surplus", { warehouse: id, lines: [line] }],
  ["write.warehouse.create_writeoff", { warehouse: id, lines: [{ nomenclatureRef: id, quantity: 1 }] }],
  [
    "write.warehouse.create_inventory",
    { warehouse: id, lines: [{ nomenclatureRef: id, accountingQuantity: 0, factQuantity: 1, price: 100 }] },
  ],
  [
    "write.warehouse.create_transfer",
    { fromWarehouse: id, toWarehouse: id, lines: [{ nomenclatureRef: id, quantity: 1 }] },
  ],
];

describe("склад документа доходит до подбора счетов строк", () => {
  for (const order of ["общая первой", "склада первой"] as const)
    it.each(CASES)(`%s (${order})`, async (tool, args) => {
      const register =
        order === "общая первой" ? [row(E, GENERAL), row(id, OWN)] : [row(id, OWN), row(E, GENERAL)];
      const fake = fake1C(
        { READ_ONLY: "false", ODATA_WRITABLE: "true" },
        { data: (q) => (q.base === "InformationRegister_СчетаУчетаНоменклатуры" ? register : undefined) },
      );
      const t = (createServer(fake.ctx) as unknown as { _registeredTools: Tools })._registeredTools[tool]!;
      const r = await t.handler(t.inputSchema.parse(args), {});
      const text = (r.content ?? []).map((c) => c.text ?? "").join("");
      expect(r.isError, text.slice(0, 300)).not.toBe(true);
      expect(text, tool).toContain(OWN);
      expect(text, tool).not.toContain(GENERAL);
    });
});
