import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";
import { SNAPSHOT_REF as id, bp30Metadata } from "./helpers/metadata-snapshot.js";
import { DOCUMENTS } from "../src/config/mapping.js";

/**
 * Все инструменты против «поддельной 1С» на снимке схемы БП 3.0 (test/helpers/fake-1c.ts):
 * чтение — как есть, запись — в режиме предпросмотра. Ни один запрос не должен обращаться к
 * полю, объекту или значению перечисления, которых нет в типовой конфигурации.
 *
 * Бюджет запросов: число обращений к 1С на каждый инструмент сверяется с
 * test/fixtures/request-budget.json. Больше нормы — тест падает (на облачной 1С каждый запрос —
 * около секунды). Если рост осознанный, обновите норму: UPDATE_REQUEST_BUDGET=1 npx vitest run test/fake-1c.test.ts
 */

const BUDGET_FILE = new URL("./fixtures/request-budget.json", import.meta.url);
const budget: Record<string, number> = existsSync(BUDGET_FILE)
  ? (JSON.parse(readFileSync(BUDGET_FILE, "utf8")) as Record<string, number>)
  : {};
const measured: Record<string, number> = {};

type Schema = {
  type?: string | string[];
  enum?: unknown[];
  properties?: Record<string, Schema>;
  required?: string[];
  items?: Schema;
  default?: unknown;
};
const manifest = JSON.parse(readFileSync(new URL("../manifest.json", import.meta.url), "utf8")) as {
  tools: Array<{ name: string; inputSchema: Schema }>;
};

const line = { nomenclatureRef: id, quantity: 1, price: 100, vatRate: "БезНДС" };
const SETTINGS_SCOPE = {
  organizationRef: id,
  counterpartyRef: id,
  contractRef: id,
  settlementKind: "РасчетыВВалютеРегламентированногоУчета",
};
// Запись регистра в подделке уже есть, версии у регистра нет — правка «одним пользователем».
const SETTINGS_UPDATE = { mode: "update", singleWriter: true, accounts: { customerSettlement: id } };

/** Аргументы, которые не подобрать по имени поля. Новый инструмент обычно сюда не нужен. */
const OVERRIDES: Record<string, Record<string, unknown>> = {
  "read.document.get_document": { entitySet: "Document_РеализацияТоваровУслуг", ref: id },
  "read.document.search_documents": { entitySet: "Document_РеализацияТоваровУслуг" },
  "read.accounting.get_document_postings": {
    documentEntity: "Document_РеализацияТоваровУслуг",
    documentRef: id,
  },
  "read.audit.get_document_history": { documentEntity: "Document_РеализацияТоваровУслуг", documentRef: id },
  "read.analytics.get_deal_history": { contractRef: id },
  "write.counterparty.update_counterparty": { ref: id, name: "Тест" },
  "write.catalog.update_nomenclature": { ref: id, name: "Тест" },
  "read.schema.describe_entity": { entitySet: "Catalog_Контрагенты" },
  "write.document.post_document": { entitySet: "Document_РеализацияТоваровУслуг", ref: id },
  "write.entity.mark_for_deletion": { entitySet: "Catalog_Контрагенты", ref: id },
  "write.entity.update_entity": {
    entitySet: "Catalog_Контрагенты",
    ref: id,
    fields: { Комментарий: "тест" },
  },
  "write.entity.create_folder": { entitySet: "Catalog_Контрагенты", name: "Тест" },
  "write.entity.move_to_folder": { entitySet: "Catalog_Контрагенты", ref: id, folder: id },
  "write.document.update_document_lines": {
    entitySet: "Document_РеализацияТоваровУслуг",
    ref: id,
    lines: [line],
  },
  "write.document.add_document_line": { entitySet: "Document_РеализацияТоваровУслуг", ref: id, line },
  "write.document.remove_document_line": {
    entitySet: "Document_РеализацияТоваровУслуг",
    ref: id,
    lineNumber: 1,
  },
  "write.document.copy_document": { entitySet: "Document_СчетНаОплатуПокупателю", ref: id },
  "write.warehouse.create_transfer": {
    fromWarehouse: "Synthetic",
    toWarehouse: "Synthetic",
    lines: [{ nomenclatureRef: id, quantity: 1, price: 100 }],
  },
  "write.warehouse.create_inventory": {
    lines: [{ nomenclatureRef: id, accountingQuantity: 0, factQuantity: 1, price: 100 }],
  },
  "write.sales.create_services_act": {
    counterpartyRef: id,
    nomenclatureGroupRef: id,
    lines: [{ name: "Тест", quantity: 1, price: 100, nomenclatureRef: id, vatRate: "БезНДС" }],
  },
  "write.money.create_bank_writeoff": { operationKind: "ОплатаПоставщику", amount: 1 },
  "write.money.create_cash_payment": { operationKind: "ОплатаПоставщику", amount: 1 },
  "write.document.apply_bank_settlement_settings": {
    entitySet: "Document_ПоступлениеНаРасчетныйСчет",
    ref: id,
    contractRef: id,
    side: "customer",
    expectedDataVersion: "AAAAAAAAAAA=",
  },
  "write.document.align_bank_payment_amounts": {
    entitySet: "Document_ПоступлениеНаРасчетныйСчет",
    ref: id,
    expectedDataVersion: "AAAAAAAAAAA=",
  },
  "write.counterparty.set_group": {
    ref: id,
    subjectKind: "counterparty",
    parentRef: "00000000-0000-0000-0000-000000000000",
    expectedDataVersion: "AAAAAAAAAAA=",
  },
  "write.counterparty.create_group": { name: "Тест", parentRef: "00000000-0000-0000-0000-000000000000" },
  "read.accounting.get_settlement_settings": { ...SETTINGS_SCOPE, scope: "contract" },
  "read.accounting.get_contract_settlement_settings": SETTINGS_SCOPE,
  "write.accounting.set_settlement_settings": { ...SETTINGS_SCOPE, ...SETTINGS_UPDATE, scope: "contract" },
  "write.accounting.set_contract_settlement_settings": { ...SETTINGS_SCOPE, ...SETTINGS_UPDATE },
  "write.operation.status": { operationId: "00000000-0000-4000-8000-000000000000" },
};

/** Значение аргумента по имени и схеме поля. */
function valueFor(name: string, s: Schema): unknown {
  if (s.enum?.length) return s.enum[0];
  if (name === "lines") return [line];
  if (name === "line") return line;
  if (/Ref$|^ref$/.test(name)) return id;
  if (name === "month") return "2026-09";
  if (/^(from|date|asOf|incomingDate)$/.test(name)) return "2026-09-01";
  if (name === "to") return "2026-09-30";
  if (name === "accountNumber") return "40702810000000000001";
  if (name === "bik") return "044525225";
  if (name === "accounts") return { customerSettlement: id };
  if (name === "expectedDataVersion") return "AAAAAAAAAAA=";
  if (name === "fields") return { Комментарий: "тест" };
  const type = Array.isArray(s.type) ? s.type[0] : s.type;
  if (type === "number" || type === "integer") return 1;
  if (type === "boolean") return false;
  if (type === "array") return [];
  if (type === "object") return {};
  return name === "account" ? "62" : "Тест";
}

/** `optional` — заполнить и необязательные аргументы (фильтры, суммы), кроме выбора базы и организации. */
function argsFor(tool: { name: string; inputSchema: Schema }, optional = false): Record<string, unknown> {
  const props = tool.inputSchema.properties ?? {};
  const args: Record<string, unknown> = {};
  const names = optional
    ? Object.keys(props).filter((n) => !["database", "organization", "operationId", "confirm"].includes(n))
    : (tool.inputSchema.required ?? []);
  for (const name of names) args[name] = valueFor(name, props[name] ?? {});
  Object.assign(args, OVERRIDES[tool.name]);
  if (tool.name.startsWith("write.") && "confirm" in props) args.confirm = false;
  return args;
}

/**
 * Потолок ответа инструмента в символах. Клиенты обрезают или отвергают большие ответы (Claude Code —
 * около 25 тыс. токенов; русский JSON — примерно 2–3 символа на токен). Большой отчёт — давайте сводку
 * и подробности по фильтру.
 */
const MAX_RESPONSE_CHARS = 50_000;

/**
 * Инструменты, которым синтетических данных подделки мало для успешного ответа (ошибка — по
 * бизнес-логике, а не из-за схемы). Нарушения схемы проверяются и для них.
 */
const EXPECTED_ERRORS: Record<string, string> = {
  // В синтетике сумма документа 0 — выравнивать не по чему; схема чтения документа проверяется.
  "write.document.align_bank_payment_amounts": "Сумма документа не заполнена",
};

afterEach(() => {
  vi.unstubAllGlobals();
});

type Tools = Record<
  string,
  {
    inputSchema?: { parse: (a: unknown) => Record<string, unknown> };
    handler: (
      a: Record<string, unknown>,
      e: unknown,
    ) => Promise<{ isError?: boolean; content?: { text?: string }[] }>;
  }
>;

/** Вызов инструмента так же, как SDK: аргументы проходят схему (значения по умолчанию, проверка типов). */
async function call(name: string, raw: Record<string, unknown>) {
  const fake = fake1C({ READ_ONLY: "true" });
  const registered = (createServer(fake.ctx) as unknown as { _registeredTools: Tools })._registeredTools[
    name
  ]!;
  const args = registered.inputSchema ? registered.inputSchema.parse(raw) : raw;
  const r = await registered.handler(args, {});
  const text = (r.content ?? []).map((c) => c.text ?? "").join("");
  expect(
    text.length,
    `${name}: ответ ${text.length} символов — больше ${MAX_RESPONSE_CHARS}, в контекст модели не поместится`,
  ).toBeLessThanOrEqual(MAX_RESPONSE_CHARS);
  return { fake, r, text };
}

describe("все инструменты против поддельной 1С на снимке схемы БП 3.0", () => {
  it.each(manifest.tools.map((t) => [t.name, t] as const))("%s", async (name, tool) => {
    const { fake, r, text } = await call(name, argsFor(tool));
    expect(fake.violations, "запросы к полям, которых нет в схеме БП 3.0").toEqual([]);
    if (!EXPECTED_ERRORS[name]) expect(r.isError, text).not.toBe(true);

    const count = fake.requests.length;
    measured[name] = count;
    if (process.env.UPDATE_REQUEST_BUDGET !== "1" && budget[name] !== undefined)
      expect(
        count,
        `${name} делает ${count} запросов к 1С вместо ${budget[name]}. Если так задумано — обновите норму (см. шапку теста).`,
      ).toBeLessThanOrEqual(budget[name]!);
  });

  it("у каждого инструмента есть норма запросов", () => {
    if (process.env.UPDATE_REQUEST_BUDGET === "1") {
      writeFileSync(BUDGET_FILE, `${JSON.stringify(measured, Object.keys(measured).sort(), 2)}\n`);
      return;
    }
    expect(manifest.tools.map((t) => t.name).filter((n) => budget[n] === undefined)).toEqual([]);
  });
});

describe("самопроверка: поддельная 1С отвечает на ошибки так же, как настоящая", () => {
  const expectRejected = async (fake: ReturnType<typeof fake1C>, run: () => Promise<unknown>) => {
    await expect(run()).rejects.toThrow();
    expect(fake.violations.length).toBe(1);
    return fake.violations[0]!;
  };

  it("поле, которого нет у объекта, в $select / $filter", async () => {
    const fake = fake1C({ READ_ONLY: "true" });
    const v = await expectRejected(fake, () =>
      fake.conn.client.getCollection(
        "Document_ПоступлениеНаРасчетныйСчет?$format=json&$select=Контрагент_Key",
      ),
    );
    expect(v).toContain("Контрагент_Key");
    const f2 = fake1C({ READ_ONLY: "true" });
    await expectRejected(f2, () =>
      f2.conn.client.getCollection(
        `Catalog_Контрагенты?$format=json&$filter=${encodeURIComponent("Выдуманное eq 'x'")}`,
      ),
    );
  });

  it("запись поля, которого нет в документе или в строке табличной части", async () => {
    // Ошибка из первой версии PR #43: у банковских документов контрагент составной, Контрагент_Key нет.
    const fake = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const v = await expectRejected(fake, () =>
      fake.conn.client.request(`Document_ПоступлениеНаРасчетныйСчет(guid'${id}')?$format=json`, "PATCH", {
        Контрагент_Key: id,
      }),
    );
    expect(v).toContain("Контрагент_Key");
    const f2 = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    await expectRejected(f2, () =>
      f2.conn.client.request(`Document_ПоступлениеНаРасчетныйСчет(guid'${id}')?$format=json`, "PATCH", {
        РасшифровкаПлатежа: [{ Выдуманное: 1 }],
      }),
    );
  });

  it("вид операции, которого нет в перечислениях", async () => {
    const fake = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const v = await expectRejected(fake, () =>
      fake.conn.client.request(`Document_ПоступлениеНаРасчетныйСчет(guid'${id}')?$format=json`, "PATCH", {
        ВидОперации: "ОплатаОтПокупателя",
      }),
    );
    expect(v).toContain("ОплатаОтПокупателя");
  });

  it("функция регистра, которой нет в $metadata, и её необъявленный параметр", async () => {
    const fake = fake1C({ READ_ONLY: "true" });
    const v = await expectRejected(fake, () =>
      fake.conn.client.getCollection("AccountingRegister_Хозрасчетный/Остатки?$format=json"),
    );
    expect(v).toContain("не объявлена");
    const f2 = fake1C({ READ_ONLY: "true" });
    const v2 = await expectRejected(f2, () =>
      f2.conn.client.getCollection(
        "AccountingRegister_Хозрасчетный/Balance(Период=datetime'2026-09-01T00:00:00')?$format=json",
      ),
    );
    expect(v2).toContain("Период");
  });

  it("объект, который не опубликован", async () => {
    const fake = fake1C({ READ_ONLY: "true" });
    await expect(fake.conn.client.getCollection("Catalog_Выдуманный?$format=json")).rejects.toThrow(
      /не найден/,
    );
  });
});

/** Необязательные аргументы, которые не заполнить по имени: взаимоисключающие или особой формы. */
const OPTIONAL_OMIT: Record<string, string[]> = {
  "read.analytics.get_deal_history": ["dealKey"], // взаимоисключающий с contractRef
  "write.document.copy_document": ["lines"], // правки строк по номеру, не новые строки
};

describe("с заполненными необязательными аргументами", () => {
  // Фильтры и параметры, которые не задаются в основном прогоне. Синтетические строки («Тест» вместо
  // имени счёта или объекта) дают отказы бизнес-логики — здесь проверяется только схема запросов.
  it.each(manifest.tools.map((t) => [t.name, t] as const))("%s", async (name, tool) => {
    const raw = { ...argsFor(tool, true), ...OVERRIDES[name] };
    for (const arg of OPTIONAL_OMIT[name] ?? []) delete raw[arg];
    if (name.startsWith("write.")) raw.confirm = false;
    const { fake } = await call(name, raw);
    expect(fake.violations, "запросы к полям, которых нет в схеме БП 3.0").toEqual([]);
  });
});

/** Инструменты, которые принимают любой документ: прогоняются по всем видам документов из маппинга. */
const ANY_DOCUMENT: Record<string, string> = {
  "read.document.search_documents": "entitySet",
  "read.document.get_document": "entitySet",
  "read.accounting.get_document_postings": "documentEntity",
  "read.audit.get_document_history": "documentEntity",
};
const DOCUMENT_SETS = [...new Set(Object.values(DOCUMENTS).flat())].filter((s) =>
  bp30Metadata().entities.has(s),
);

/**
 * Известные ошибки — тест ждёт, что нарушение ещё есть. Исправили — уберите строку отсюда.
 */
const KNOWN_VIOLATIONS: Record<string, string> = {};

describe("инструменты для любого документа — по всем видам документов БП 3.0", () => {
  const cases = Object.entries(ANY_DOCUMENT).flatMap(([name, field]) =>
    DOCUMENT_SETS.map((set) => [`${name} ${set}`, name, field, set] as const),
  );
  it.each(cases)("%s", async (label, name, field, set) => {
    const tool = manifest.tools.find((t) => t.name === name)!;
    const { fake } = await call(name, { ...argsFor(tool, true), ...OVERRIDES[name], [field]: set });
    const known = KNOWN_VIOLATIONS[label];
    if (known)
      expect(fake.violations.join("\n"), "ошибка исправлена — уберите её из KNOWN_VIOLATIONS").toContain(
        known,
      );
    else expect(fake.violations, "запросы к полям, которых нет в схеме БП 3.0").toEqual([]);
  });
});

// Снимок — схема реальной БП 3.0: в записях регистра (_RecordType) нет полей ExtDimension. Субконто
// проводок должны браться из RecordsWithExtDimensions; раньше тест проводок рисовал субконто прямо в
// записях, и то, что на живой базе их не было никогда, тесты не видели (#49).
describe("проводки документа на схеме БП 3.0 — субконто из «Движений с субконто»", () => {
  it("в записях регистра снимка субконто нет", () => {
    const props = bp30Metadata().entities.get("AccountingRegister_Хозрасчетный_RecordType")!.properties;
    expect(props.some((p) => p.name.startsWith("ExtDimension"))).toBe(false);
  });
  it("get_document_postings не отвечает «субконто недоступны»", async () => {
    const { r: result, fake } = await call("read.accounting.get_document_postings", {
      documentEntity: "Document_РеализацияТоваровУслуг",
      documentRef: id,
    });
    expect(result.isError).toBeFalsy();
    const sc = result.structuredContent as {
      postings: Array<{ debit: { dimensionsStatus: string } | null }>;
      dimensionsNote?: string;
    };
    expect(sc.postings[0]!.debit!.dimensionsStatus).not.toBe("not_in_schema");
    expect(sc.dimensionsNote).toBeUndefined();
    // «Движения с субконто» — функция самого регистра (поддельная 1С пишет в requests имя регистра).
    expect(fake.requests.some((r) => r.resource === "AccountingRegister_Хозрасчетный")).toBe(true);
    expect(fake.violations).toEqual([]);
  });
});
