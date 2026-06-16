import { describe, it, expect } from "vitest";
import { NotPublishedError, publicationHelp } from "../src/odata/publication.js";

describe("вежливый фоллбэк публикации", () => {
  it("publicationHelp называет объект и путь в «Состав OData»", () => {
    const txt = publicationHelp([{ label: "Справочник «Контрагенты»", candidates: ["Catalog_Контрагенты"] }]);
    expect(txt).toContain("Контрагенты");
    expect(txt).toContain("Состав");
  });
  it("NotPublishedError несёт список недостающих объектов", () => {
    const e = new NotPublishedError([{ label: "X", candidates: ["Y"] }]);
    expect(e.missing).toHaveLength(1);
    expect(e.name).toBe("NotPublishedError");
  });
});
