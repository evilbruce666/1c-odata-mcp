import { readFileSync } from "node:fs";
import { vi } from "vitest";
import { parseConfig } from "../../src/config/env.js";
import { Connection } from "../../src/context.js";
import { createServer } from "../../src/mcp/server.js";
import { parseMetadata } from "../../src/odata/metadata.js";
import type { MetadataMap } from "../../src/types/odata.js";

/**
 * Снимок $metadata типовой Бухгалтерии 3.0 (облако, октябрь 2026): только объекты, которые использует
 * коннектор, их табличные части и все перечисления. Данных базы в нём нет — только схема.
 * Обновить: снять $metadata с базы и отфильтровать по именам объектов из src (см. CONTRIBUTING).
 */
export const BP30_METADATA_XML = readFileSync(
  new URL("../fixtures/metadata/bp30.xml", import.meta.url),
  "utf8",
);

/**
 * Снимок $metadata УТ 11.5 (ноябрь 2026): регистры «Товары на складах» и «Выручка и себестоимость продаж», номенклатура, склады, ключи аналитики учёта номенклатуры, отчёт о розничных продажах. Поля доработок базы
 * (Б_*, Операции0, …ДляЗП) убраны: в снимке только типовая схема, данных базы нет.
 */
export const UT11_METADATA_XML = readFileSync(
  new URL("../fixtures/metadata/ut11.xml", import.meta.url),
  "utf8",
);

let parsedUt: MetadataMap | undefined;
export function ut11Metadata(): MetadataMap {
  return (parsedUt ??= parseMetadata(UT11_METADATA_XML));
}

let parsed: MetadataMap | undefined;
export function bp30Metadata(): MetadataMap {
  return (parsed ??= parseMetadata(BP30_METADATA_XML));
}

/** Значения перечислений снимка: имя перечисления → множество значений. */
export function bp30Enums(): Map<string, Set<string>> {
  const enums = new Map<string, Set<string>>();
  for (const m of BP30_METADATA_XML.matchAll(/<EnumType Name="([^"]+)"[\s\S]*?<\/EnumType>/g))
    enums.set(m[1]!, new Set([...m[0].matchAll(/<Member Name="([^"]+)"/g)].map((x) => x[1]!)));
  return enums;
}

/** Поля сущности снимка: имя → тип Edm. */
export function bp30Fields(entitySet: string): Map<string, string> | undefined {
  const meta = bp30Metadata().entities.get(entitySet);
  return meta && new Map(meta.properties.map((p) => [p.name, p.type]));
}

export const SNAPSHOT_REF = "11111111-1111-4111-8111-111111111111";

/**
 * Подключение к «базе» по снимку схемы: метаданные — из снимка, чтения отвечают типовыми
 * синтетическими объектами. Запись идёт через настоящий клиент и журнал (каталог журнала задаёт
 * вызывающий через env), поэтому `send` можно подменить, а журнал проверить.
 */
export function snapshotConnection(env: Record<string, string> = {}): Connection {
  const parsed = parseConfig({
    ODATA_BASE_URL: "https://example.invalid/odata/",
    ODATA_USERNAME: "u",
    ODATA_PASSWORD: "p",
    ...env,
  });
  const conn = new Connection(parsed.connections[0]!, parsed.behavior);
  const meta = bp30Metadata();
  const id = SNAPSHOT_REF;
  const row = { LineNumber: "1", Номенклатура_Key: id, Количество: 1, Цена: 100, Сумма: 100 };
  vi.spyOn(conn, "available").mockResolvedValue(new Set(meta.entities.keys()));
  vi.spyOn(conn, "getMetadata").mockResolvedValue(meta);
  vi.spyOn(conn.client, "getEntity").mockResolvedValue({
    Ref_Key: id,
    Description: "Synthetic",
    Owner_Key: id,
    Организация_Key: id,
    Контрагент_Key: id,
    ДоговорКонтрагента_Key: id,
    ВидДоговора: "СПокупателем",
    ВалютаВзаиморасчетов_Key: id,
    Code: "643",
    Number: "0000-000001",
    Date: "2026-10-01T00:00:00",
    Posted: false,
    DeletionMark: false,
    Товары: [row, { ...row, LineNumber: "2" }],
  });
  vi.spyOn(conn.client, "getCollection").mockImplementation(async (path: string) => ({
    value: path.startsWith("ChartOfAccounts_")
      ? ["41.01", "60.01", "60.02", "62.01", "62.02", "90.01.1", "90.02.1", "91.01", "91.02", "94"].map(
          (Code) => ({ Ref_Key: id, Code, Description: Code }),
        )
      : [{ Ref_Key: id, Description: "Synthetic", Code: "1", Owner_Key: id }],
  }));
  vi.spyOn(conn.client, "getText").mockResolvedValue("");
  return conn;
}

export type SnapshotTool = {
  inputSchema?: { shape?: Record<string, unknown> };
  handler: (
    a: Record<string, unknown>,
    e: unknown,
  ) => Promise<{
    isError?: boolean;
    structuredContent?: Record<string, unknown>;
    content?: { text?: string }[];
  }>;
};

export function snapshotTools(conn: Connection): Record<string, SnapshotTool> {
  return (
    createServer({ db: () => conn, databases: () => [], behavior: {} } as never) as unknown as {
      _registeredTools: Record<string, SnapshotTool>;
    }
  )._registeredTools;
}

const line = { nomenclatureRef: SNAPSHOT_REF, quantity: 1, price: 100, vatRate: "БезНДС" };
const stockLine = { nomenclatureRef: SNAPSHOT_REF, quantity: 1, price: 100 };
/** Аргументы предпросмотра инструментов создания (по ним тесты прогоняют запись по снимку). */
export const CREATE_CASES: Array<[string, Record<string, unknown>]> = [
  ["write.sales.create_invoice", { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] }],
  [
    "write.sales.create_shipment",
    { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] },
  ],
  ["write.sales.create_act", { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] }],
  [
    "write.purchase.create_purchase",
    { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] },
  ],
  [
    "write.purchase.create_supplier_invoice",
    { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] },
  ],
  [
    "write.warehouse.create_return_from_customer",
    { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] },
  ],
  [
    "write.warehouse.create_return_to_supplier",
    { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, lines: [line] },
  ],
  ["write.warehouse.create_surplus", { lines: [stockLine] }],
  ["write.warehouse.create_writeoff", { lines: [stockLine] }],
  [
    "write.warehouse.create_inventory",
    { lines: [{ nomenclatureRef: SNAPSHOT_REF, accountingQuantity: 0, factQuantity: 1, price: 100 }] },
  ],
  ["write.money.create_payment", { counterpartyRef: SNAPSHOT_REF, contractRef: SNAPSHOT_REF, amount: 1 }],
  [
    "write.money.create_bank_writeoff",
    {
      operationKind: "ОплатаПоставщику",
      amount: 1,
      counterpartyRef: SNAPSHOT_REF,
      contractRef: SNAPSHOT_REF,
    },
  ],
  [
    "write.money.create_cash_receipt",
    {
      operationKind: "ОплатаПокупателя",
      amount: 1,
      counterpartyRef: SNAPSHOT_REF,
      contractRef: SNAPSHOT_REF,
    },
  ],
  [
    "write.money.create_cash_payment",
    {
      operationKind: "ОплатаПоставщику",
      amount: 1,
      counterpartyRef: SNAPSHOT_REF,
      contractRef: SNAPSHOT_REF,
    },
  ],
  [
    "write.money.create_payout_order",
    { counterpartyRef: SNAPSHOT_REF, amount: 1, number: "1", purposeText: "тест", recipientText: "тест" },
  ],
  ["write.counterparty.create_counterparty", { name: "Тест" }],
  ["write.catalog.create_nomenclature", { name: "Тест" }],
  ["write.catalog.create_contract", { counterpartyRef: SNAPSHOT_REF, kind: "СПокупателем" }],
];
