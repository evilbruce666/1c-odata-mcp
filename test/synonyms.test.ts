import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { CATALOGS, DOCUMENTS, REGISTERS } from "../src/config/mapping.js";
import { SYNONYMS } from "../src/config/synonyms.js";
import { describeObject, publicationHelp } from "../src/odata/publication.js";

/**
 * Имена-кандидаты других релизов и конфигураций: ни в одной из проверенных баз их нет, поэтому и названия взять
 * неоткуда. Новый объект в коде без названия сюда не добавляют молча — название снимают из базы.
 */
const UNVERIFIED = new Set([
  "Catalog_Договор",
  "Catalog_Договоры",
  "Catalog_СкладыМеста",
  "Catalog_ТипыЦен",
  "Document_ЗаказПокупателя",
  "Document_ПоступлениеТоваровИУслуг",
  "Document_РеализацияТоваровИУслуг",
  "Document_СчетПокупателю",
  "InformationRegister_ОтветственныеЛицаОрганизации",
]);

const ENTITY =
  /["'`]((?:Catalog|Document|InformationRegister|AccumulationRegister|AccountingRegister|ChartOfAccounts|ChartOfCharacteristicTypes)_[А-Яа-яЁёA-Za-z0-9_]+)["'`]/g;

function sources(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const path = join(dir, e.name);
    if (e.isDirectory()) return e.name === "scripts" ? [] : sources(path);
    return path.endsWith(".ts") && !path.endsWith("synonyms.ts") ? [path] : [];
  });
}

describe("названия объектов для подсказки о «Составе OData»", () => {
  it("у каждого объекта, названного в коде, есть проверенное название или явная отметка «не проверено»", () => {
    const names = new Set(
      sources("src").flatMap((f) =>
        [...readFileSync(f, "utf8").matchAll(ENTITY)].map((m) => m[1]!.replace(/_RecordType$/, "")),
      ),
    );
    expect(names.size).toBeGreaterThan(100);
    expect([...names].filter((n) => !SYNONYMS[n] && !UNVERIFIED.has(n))).toEqual([]);
    expect([...UNVERIFIED].filter((n) => SYNONYMS[n] || !names.has(n))).toEqual([]);
  });

  it("в каждой группе кандидатов хотя бы одно имя с названием", () => {
    const groups = [CATALOGS, DOCUMENTS, REGISTERS].flatMap((g) => Object.entries(g));
    const bare = groups.filter(([, c]) => !c.some((n) => SYNONYMS[n.replace(/_RecordType$/, "")]));
    expect(bare.map(([k]) => k)).toEqual([]);
  });

  it("регистр называется как в «Составе», без таблицы записей и с отличием в УТ", () => {
    expect(describeObject("InformationRegister_РасчетНалогаУСН_RecordType")).not.toContain("_RecordType");
    expect(describeObject("InformationRegister_НастройкиСистемыНалогообложения")).toBe(
      "регистр сведений «Система налогообложения» (в Управлении торговлей — «Настройки системы налогообложения») " +
        "[InformationRegister_НастройкиСистемыНалогообложения]",
    );
  });

  it("без проверенного названия остаётся техническое имя", () => {
    expect(describeObject("Catalog_ТипыЦен")).toBe("Catalog_ТипыЦен");
    expect(describeObject("Catalog_НетТакого")).toBe("Catalog_НетТакого");
  });

  it("подсказка не повторяет объект дважды из-за таблицы записей", () => {
    const reg = "AccountingRegister_Хозрасчетный";
    const text = publicationHelp([{ label: "Проводки", candidates: [reg, `${reg}_RecordType`] }]);
    expect(text.match(/\[AccountingRegister_Хозрасчетный\]/g)).toHaveLength(1);
  });
});
