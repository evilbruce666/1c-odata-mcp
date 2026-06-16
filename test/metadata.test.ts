import { describe, it, expect } from "vitest";
import { classify } from "../src/odata/metadata.js";

describe("classify — класс объекта по префиксу EntitySet", () => {
  it("Catalog_ → catalog + shortName", () => {
    expect(classify("Catalog_Контрагенты")).toEqual({ class: "catalog", shortName: "Контрагенты" });
  });
  it("Document_ → document", () => {
    expect(classify("Document_РеализацияТоваровУслуг").class).toBe("document");
  });
  it("AccountingRegister_ → accountingRegister", () => {
    expect(classify("AccountingRegister_Хозрасчетный").class).toBe("accountingRegister");
  });
  it("неизвестный префикс → other (имя как есть)", () => {
    expect(classify("Foo_Bar")).toEqual({ class: "other", shortName: "Foo_Bar" });
  });
});
