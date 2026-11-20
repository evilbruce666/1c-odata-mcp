import { describe, expect, it } from "vitest";
import { CATALOGS, DOCUMENTS, REGISTERS } from "../src/config/mapping.js";
import {
  CREATE_CASES,
  SNAPSHOT_REF as id,
  bp30Enums,
  bp30Fields,
  bp30Metadata,
  snapshotConnection,
  snapshotTools,
} from "./helpers/metadata-snapshot.js";

/**
 * Код против настоящей схемы БП 3.0, а не против рукописных заглушек: каждое поле, которое пишет
 * инструмент, есть у объекта, а вид операции — среди значений перечислений. Так ловятся поля,
 * которых нет в типовой конфигурации (например, Контрагент_Key у банковских документов — там
 * составной Контрагент + Контрагент_Type), и несуществующие виды операций.
 */

const meta = bp30Metadata();
const operationKinds = new Set(
  [...bp30Enums()].filter(([name]) => name.startsWith("ВидыОпераций")).flatMap(([, v]) => [...v]),
);

/**
 * Не опубликованы в базе, с которой снят снимок; инструменты без них отвечают понятной ошибкой. Заказы клиентов и
 * поставщикам, заявки на расходование — документы Управления торговлей (в Бухгалтерии их нет).
 */
const NOT_IN_SNAPSHOT = new Set([
  "stock",
  "responsiblePersons",
  "customerOrder",
  "supplierOrder",
  "paymentRequest",
]);

describe("снимок схемы БП 3.0", () => {
  it("каждый объект из mapping.ts есть в снимке (хотя бы один из вариантов имени)", () => {
    const missing = Object.entries({ ...CATALOGS, ...DOCUMENTS, ...REGISTERS })
      .filter(([key, names]) => !NOT_IN_SNAPSHOT.has(key) && !names.some((n) => meta.entities.has(n)))
      .map(([key]) => key);
    expect(missing).toEqual([]);
  });

  it("проверка ловит поля и виды операций, которых нет в типовой конфигурации", () => {
    // Банковские документы: контрагент составной (Контрагент + Контрагент_Type), Контрагент_Key нет.
    expect(
      unknownFields("Document_ПоступлениеНаРасчетныйСчет", {
        Контрагент_Key: id,
        РасшифровкаПлатежа: [{ ДоговорКонтрагента_Key: id, СчетУчетаРасчетовПоАвансам_Key: id }],
      }),
    ).toEqual(["Контрагент_Key"]);
    expect(operationKinds.has("ОплатаПокупателя")).toBe(true);
    expect(operationKinds.has("ОплатаОтПокупателя")).toBe(false);
  });
});

/** Поля payload, которых нет в схеме: у шапки — в самом объекте, у строк — в табличной части. */
function unknownFields(entitySet: string, payload: Record<string, unknown>): string[] {
  const fields = bp30Fields(entitySet);
  if (!fields) return [`объект ${entitySet} отсутствует в снимке`];
  const bad: string[] = [];
  for (const [key, value] of Object.entries(payload)) {
    if (Array.isArray(value) && value.length && typeof value[0] === "object") {
      const section = bp30Fields(`${entitySet}_${key}`);
      if (!section) bad.push(`табличная часть ${key}`);
      else
        for (const row of value as Record<string, unknown>[])
          for (const k of Object.keys(row)) if (!section.has(k)) bad.push(`${key}.${k}`);
    } else if (!fields.has(key)) bad.push(key);
  }
  return [...new Set(bad)];
}

describe("инструменты создания пишут только поля типовой схемы", () => {
  it.each(CREATE_CASES)("%s", async (name, args) => {
    const tools = snapshotTools(snapshotConnection());
    const r = await tools[name]!.handler({ ...args, confirm: false }, {});
    expect(r.isError, JSON.stringify(r)).not.toBe(true);
    const sc = r.structuredContent ?? {};
    const entitySet = String(sc["entitySet"] ?? sc["willCreate"] ?? "");
    const payload = sc["payload"] as Record<string, unknown> | undefined;
    expect(payload, JSON.stringify(sc).slice(0, 300)).toBeDefined();
    expect(unknownFields(entitySet, payload!)).toEqual([]);
    if (typeof payload!["ВидОперации"] === "string")
      expect(operationKinds.has(payload!["ВидОперации"])).toBe(true);
  });
});
