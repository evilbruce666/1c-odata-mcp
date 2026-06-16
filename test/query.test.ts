import { describe, it, expect } from "vitest";
import { odataString, odataGuid, contains, and, or, buildQuery } from "../src/odata/query.js";

describe("odataString", () => {
  it("оборачивает в одинарные кавычки", () => {
    expect(odataString("Рога и Копыта")).toBe("'Рога и Копыта'");
  });
  it("удваивает одинарную кавычку (защита от инъекции)", () => {
    expect(odataString("O'Brien")).toBe("'O''Brien'");
  });
});

describe("odataGuid", () => {
  it("формирует guid-литерал", () => {
    expect(odataGuid("13fb122a-a706-11e5-ba67-3085a93ddca2")).toBe(
      "guid'13fb122a-a706-11e5-ba67-3085a93ddca2'",
    );
  });
  it("убирает фигурные скобки", () => {
    expect(odataGuid("{13fb122a-a706-11e5-ba67-3085a93ddca2}")).toContain("guid'13fb122a");
  });
  it("бросает на некорректном GUID", () => {
    expect(() => odataGuid("не-гуид")).toThrow();
  });
});

describe("buildQuery — кодирование (регресс 1С)", () => {
  it("пробел кодируется как %20, НЕ как + (иначе 1С отвечает 400 на $filter)", () => {
    const q = buildQuery({ filter: "Account_Key eq guid'x'" });
    expect(q).toContain("%20");
    expect(q).not.toContain("+");
  });
  it("$format=json присутствует всегда", () => {
    expect(buildQuery({})).toContain("$format=json");
  });
  it("прокидывает select/top/skip", () => {
    const q = buildQuery({ select: ["a", "b"], top: 5, skip: 10 });
    expect(q).toContain("$top=5");
    expect(q).toContain("$skip=10");
    expect(q).toContain("$select=");
  });
});

describe("операторы фильтра", () => {
  it("contains → substringof", () => {
    expect(contains("Description", "Ро")).toBe("substringof('Ро', Description)");
  });
  it("and склеивает и пропускает undefined", () => {
    expect(and("a eq 1", undefined, "b eq 2")).toBe("a eq 1 and b eq 2");
  });
  it("or оборачивает в скобки при нескольких условиях", () => {
    expect(or("a", "b")).toBe("(a or b)");
  });
});
