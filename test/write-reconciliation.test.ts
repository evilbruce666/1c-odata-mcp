import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import {
  CREATE_CASES,
  SNAPSHOT_REF as id,
  snapshotConnection,
  snapshotTools,
} from "./helpers/metadata-snapshot.js";

/**
 * Любая запись с operationId должна оставлять в журнале способ сверки: назначенный Ref_Key
 * (созданный объект ищется простым GET) или проверку (`check`, например число строк документа).
 * Иначе после потерянного ответа 1С `write.operation.status` не отличит «не дошло» от «создано»
 * и посоветует повторить — так получаются дубли. Новый инструмент записи добавьте в ARGS.
 */

const line = { nomenclatureRef: id, quantity: 1, price: 100, vatRate: "БезНДС" };
const doc = { entitySet: "Document_СчетНаОплатуПокупателю", ref: id };
const ARGS: Record<string, Record<string, unknown>> = {
  ...Object.fromEntries(CREATE_CASES),
  "write.counterparty.create_group": { name: "Synthetic", parentRef: "00000000-0000-0000-0000-000000000000" },
  "write.counterparty.set_group": {
    ref: id,
    subjectKind: "counterparty",
    parentRef: "00000000-0000-0000-0000-000000000000",
    expectedDataVersion: "v1",
  },
  "write.accounting.set_contract_settlement_settings": {
    organizationRef: id,
    counterpartyRef: id,
    contractRef: id,
    settlementKind: "РасчетыВВалютеРегламентированногоУчета",
    mode: "create",
    accounts: { supplierSettlement: id },
  },
  "write.accounting.set_settlement_settings": {
    scope: "contract",
    organizationRef: id,
    counterpartyRef: id,
    contractRef: id,
    settlementKind: "РасчетыВВалютеРегламентированногоУчета",
    mode: "create",
    accounts: { supplierSettlement: id },
  },
  "write.document.apply_bank_settlement_settings": {
    entitySet: "Document_СписаниеСРасчетногоСчета",
    ref: id,
    contractRef: id,
    side: "supplier",
    expectedDataVersion: "v1",
  },
  "write.document.align_bank_payment_amounts": {
    entitySet: "Document_ПоступлениеНаРасчетныйСчет",
    ref: id,
    expectedDataVersion: "v1",
  },
  "write.document.post_document": { ...doc, post: true, expectedDataVersion: "v1" },
  "write.sales.create_services_act": {
    counterpartyRef: id,
    contractRef: id,
    nomenclatureGroupRef: id,
    lines: [{ name: "Тест", quantity: 1, price: 100, nomenclatureRef: id, vatRate: "БезНДС" }],
  },
  "write.warehouse.create_transfer": {
    fromWarehouse: "Synthetic",
    toWarehouse: "Synthetic",
    lines: [{ nomenclatureRef: id, quantity: 1, price: 100 }],
  },
  "write.sales.create_issued_invoice": { baseDocumentRef: id },
  "write.purchase.create_received_invoice": { baseDocumentRef: id },
  "write.counterparty.create_bank_account": {
    ownerRef: id,
    accountNumber: "40702810000000000001",
    bik: "044525225",
  },
  "write.counterparty.create_contact_person": { ownerRef: id, name: "Тест" },
  "write.entity.create_folder": { entitySet: "Catalog_Контрагенты", name: "Тест" },
  "write.document.copy_document": doc,
  "write.document.add_document_line": { ...doc, line },
  "write.document.remove_document_line": { ...doc, lineNumber: 1 },
};

function setup() {
  const conn = snapshotConnection({
    READ_ONLY: "false",
    ODATA_WRITABLE: "true",
    ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-reconcile-")),
  });
  const originalGet = vi.mocked(conn.client.getEntity).getMockImplementation()!;
  vi.mocked(conn.client.getEntity).mockImplementation(async (path) => {
    if (path.startsWith("InformationRegister_СчетаУчетаРасчетовСКонтрагентами("))
      throw new ODataError({ kind: "not_found", message: "Synthetic absent setting" });
    const record = await originalGet(path);
    // Выравнивание сумм: рублёвое поступление с одной строкой расшифровки без сумм.
    const unalignedReceipt = path.startsWith("Document_ПоступлениеНаРасчетныйСчет(")
      ? {
          СуммаДокумента: 100,
          РасшифровкаПлатежа: [
            {
              ...(record.РасшифровкаПлатежа as Array<Record<string, unknown>>)?.[0],
              ДоговорКонтрагента_Key: id,
              СуммаПлатежа: 0,
              СуммаВзаиморасчетов: 0,
              СтавкаНДС: "БезНДС",
            },
          ],
        }
      : {};
    if (path.startsWith("Catalog_Валюты(")) return { ...record, Code: "643" };
    return {
      ...record,
      Ref_Key: path.match(/guid'([^']+)'/)?.[1] ?? id,
      IsFolder: false,
      Parent_Key: "00000000-0000-0000-0000-000000000000",
      DataVersion: "v1",
      ВидДоговора: "Прочее",
      РасчетыВУсловныхЕдиницах: false,
      ЗапретитьИспользоватьВПроводках: false,
      Контрагент: id,
      Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
      ВидОперации: "ОплатаПоставщику",
      ВалютаДокумента_Key: id,
      РасшифровкаПлатежа: [],
      ...unalignedReceipt,
    };
  });
  const originalCollection = vi.mocked(conn.client.getCollection).getMockImplementation()!;
  vi.mocked(conn.client.getCollection).mockImplementation(async (path) =>
    path.startsWith("InformationRegister_СчетаУчетаРасчетовСКонтрагентами?")
      ? {
          value: [
            {
              Организация_Key: id,
              Контрагент_Key: id,
              Договор_Key: id,
              ВидРасчетовПоДоговору: "РасчетыВВалютеРегламентированногоУчета",
              СчетУчетаРасчетовСПоставщиком_Key: id,
              СчетУчетаАвансовВыданных_Key: id,
            },
          ],
        }
      : originalCollection(path),
  );
  // Запрос в 1С ушёл, ответ потерян: исход неизвестен, как при таймауте сети.
  vi.spyOn(conn.client, "request").mockImplementation(async (_path, method = "GET") => {
    if (method === "GET") return {} as never;
    throw new ODataError({ kind: "timeout", message: "ответ потерян (тест)" });
  });
  return { conn, tools: snapshotTools(conn) };
}

const journaled = Object.entries(setup().tools)
  .filter(
    ([name, t]) => name !== "write.operation.status" && t.inputSchema?.shape?.["operationId"] !== undefined,
  )
  .map(([name]) => name)
  .sort();

describe("сверка записи после потерянного ответа", () => {
  it("у каждого инструмента с operationId есть аргументы в тесте", () => {
    expect(journaled.filter((name) => !ARGS[name])).toEqual([]);
  });

  it.each(journaled.filter((name) => ARGS[name]))("%s", async (name) => {
    const { conn, tools } = setup();
    const args = ARGS[name]!;
    const preview = await tools[name]!.handler({ ...args, confirm: false }, {});
    expect(preview.isError, preview.content?.[0]?.text).not.toBe(true);
    const operationId = preview.structuredContent?.["operationId"];
    expect(typeof operationId).toBe("string");

    const confirmed = await tools[name]!.handler({ ...args, confirm: true, operationId }, {});
    expect(confirmed.isError).toBe(true);

    const entry = await conn.client.operationEntry(String(operationId));
    expect(entry?.state).toBe("outcome_unknown");
    expect(
      entry?.refKey !== undefined || entry?.check !== undefined,
      "в журнале нет ни назначенного Ref_Key, ни проверки — после таймаута операцию не сверить",
    ).toBe(true);

    // Объект в 1С есть (GET отвечает): созданное по назначенному Ref_Key статус должен найти.
    const status = await tools["write.operation.status"]!.handler({ operationId }, {});
    const state = status.structuredContent?.["status"];
    if (entry?.refKey) expect(state).toBe("found_reconciled");
    else expect(["found_reconciled", "not_found", "unverifiable"]).toContain(state);
  });

  it("самопроверка: запись без назначенного Ref_Key и без проверки видна как несверяемая", async () => {
    const { conn, tools } = setup();
    // Так журналировали группы в PR #43: prepareCreate с withRefKey: false и без check.
    const prepare = conn.client.prepareCreate.bind(conn.client);
    vi.spyOn(conn.client, "prepareCreate").mockImplementation((set, payload) =>
      prepare(set, payload, { withRefKey: false }),
    );
    const args = ARGS["write.counterparty.create_counterparty"]!;
    const preview = await tools["write.counterparty.create_counterparty"]!.handler(
      { ...args, confirm: false },
      {},
    );
    const operationId = String(preview.structuredContent?.["operationId"]);
    await tools["write.counterparty.create_counterparty"]!.handler(
      { ...args, confirm: true, operationId },
      {},
    );
    const entry = await conn.client.operationEntry(operationId);
    expect(entry?.refKey === undefined && entry?.check === undefined).toBe(true);
  });
});
