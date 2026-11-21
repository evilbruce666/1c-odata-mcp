import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { SALES_DOCS, tradeSalesSet } from "../src/tools/trade-sales.js";
import { MONEY_DOCS, tradeMoneySet } from "../src/tools/trade-money.js";
import { STOCK_DOCS, tradeStockSet } from "../src/tools/trade-warehouse.js";
import { fake1C } from "./helpers/fake-1c.js";
import { SNAPSHOT_REF as ID } from "./helpers/metadata-snapshot.js";

/**
 * Бухгалтерия, где регистр бухгалтерии не опубликован: защита записи обязана отказать, а не принять документ
 * Бухгалтерии за документ УТ. Ревью выпуска 5: сигнатура счёта-фактуры выданного УТ совпала со схемой БП 3.0, и
 * create_issued_invoice, update_entity, post_document и mark_for_deletion уходили в 1С.
 */
function noAccounting(data?: Parameters<typeof fake1C>[1]) {
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-gate-")),
    },
    data,
  );
  const available = fake.conn.available.bind(fake.conn);
  vi.spyOn(fake.conn, "available").mockImplementation(
    async () => new Set([...(await available())].filter((s) => !s.startsWith("AccountingRegister_"))),
  );
  return fake;
}
afterEach(() => vi.restoreAllMocks());

describe("Бухгалтерия без регистра бухгалтерии: документы Бухгалтерии не принимаются за документы УТ", () => {
  it("ни одна сигнатура склада, денег, продаж и закупок УТ не совпадает со схемой БП 3.0", async () => {
    const fake = noAccounting();
    const hits: string[] = [];
    for (const k of Object.keys(SALES_DOCS) as (keyof typeof SALES_DOCS)[])
      if (await tradeSalesSet(fake.conn, k)) hits.push(SALES_DOCS[k]);
    for (const k of Object.keys(MONEY_DOCS) as (keyof typeof MONEY_DOCS)[])
      if (await tradeMoneySet(fake.conn, k)) hits.push(MONEY_DOCS[k]);
    for (const k of Object.keys(STOCK_DOCS) as (keyof typeof STOCK_DOCS)[])
      if (await tradeStockSet(fake.conn, k)) hits.push(STOCK_DOCS[k]);
    expect(hits).toEqual([]);
  });

  it.each([
    ["write.sales.create_issued_invoice", { baseDocumentRef: ID }],
    [
      "write.entity.update_entity",
      { entitySet: "Document_СчетФактураВыданный", ref: ID, fields: { Комментарий: "x" } },
    ],
    ["write.document.post_document", { entitySet: "Document_СчетФактураВыданный", ref: ID, post: true }],
    ["write.entity.mark_for_deletion", { entitySet: "Document_СчетФактураВыданный", ref: ID, mark: true }],
  ] as const)("%s — отказ, ни одной записи", async (name, args) => {
    const fake = noAccounting({
      data: (q) =>
        q.base === "Document_РеализацияТоваровУслуг" && !q.fn
          ? [
              {
                Ref_Key: ID,
                Posted: true,
                Организация_Key: ID,
                Контрагент_Key: ID,
                ДоговорКонтрагента_Key: ID,
              },
            ]
          : undefined,
    });
    const tool = (
      createServer(fake.ctx) as unknown as {
        _registeredTools: Record<
          string,
          {
            inputSchema: { parse: (v: unknown) => Record<string, unknown> };
            handler: (a: unknown, e: unknown) => Promise<{ isError?: boolean; content?: { text: string }[] }>;
          }
        >;
      }
    )._registeredTools[name]!;
    const parsed = tool.inputSchema.parse(args);
    let r = (await tool.handler(parsed, {})) as {
      isError?: boolean;
      content?: { text: string }[];
      structuredContent?: { operationId?: string };
    };
    if (!r.isError)
      r = await tool.handler({ ...parsed, confirm: true, operationId: r.structuredContent?.operationId }, {});
    expect(r.isError).toBe(true);
    expect(r.content?.map((c) => c.text).join("")).toContain("только для Бухгалтерии");
    expect(fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });
});
