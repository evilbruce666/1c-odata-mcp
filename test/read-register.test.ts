import { afterEach, describe, expect, it, vi } from "vitest";
import type { Connection } from "../src/context.js";
import { readRegisterRows } from "../src/odata/register-rows.js";
import { callTool } from "./helpers/call-tool.js";

const TOOL = "read.schema.read_register";
const SETTINGS = "InformationRegister_НастройкиСистемыНалогообложения";
const ACCRUALS = "InformationRegister_РасчетНачисленияНалоговНаЕНС";
const FLAT = `${ACCRUALS}_RecordType`;
const REF = "11111111-1111-4111-8111-111111111111";
const row = (i: number, extra: Record<string, unknown> = {}) => ({
  СрокУплаты: `2030-0${1 + (i % 9)}-28T00:00:00`,
  Начислено: i,
  Секретный: "не должно попасть",
  ...extra,
});

interface Page {
  register: string;
  matched: number;
  returned: number;
  truncated: boolean;
  nextOffset?: number;
  snapshot: string;
  dataChanged?: boolean;
  note: string;
  rows: Array<Record<string, unknown>>;
}

async function read(
  args: Record<string, unknown>,
  rows: Array<Record<string, unknown>> = [row(1)],
  schema?: "ut11",
) {
  const got = await callTool(
    TOOL,
    { register: ACCRUALS, select: ["СрокУплаты", "Начислено"], ...args },
    (q) => (q.base === "Catalog_Организации" ? [{ Ref_Key: REF, Description: "Организация" }] : rows),
    schema,
  );
  const query = got.queries.find((q) => /Register_/.test(q.base));
  return {
    ...got,
    page: got.result.structuredContent as unknown as Page,
    error: got.result.isError ? JSON.stringify(got.result.content) : undefined,
    params: query?.params,
    base: query?.base,
  };
}

afterEach(() => vi.unstubAllGlobals());

describe("что читается", () => {
  it("у регистра с регистратором — таблица записей, и в ответе только выбранные поля", async () => {
    const { page, base, params, violations } = await read({});
    expect(violations).toEqual([]);
    expect(base).toBe(FLAT);
    expect(page.register).toBe(FLAT);
    expect(params!.get("$select")).toBe("СрокУплаты,Начислено");
    expect(page.rows).toEqual([{ СрокУплаты: "2030-02-28T00:00:00", Начислено: 1 }]);
    expect(page.note).toContain("не складывайте");
  });

  it("имя таблицы записей принимается так же, а независимый регистр читается сам", async () => {
    await expect(read({ register: FLAT })).resolves.toMatchObject({ base: FLAT });
    await expect(read({ register: SETTINGS, select: ["Period"] })).resolves.toMatchObject({ base: SETTINGS });
  });

  it("не регистр — отказ с профильным инструментом, без запроса", async () => {
    for (const [register, hint] of [
      ["Catalog_Контрагенты", "find_counterparty"],
      ["Document_РеализацияТоваровУслуг", "search_documents"],
      [`${SETTINGS}/SliceLast()`, "list_entities"],
    ]) {
      const got = await read({ register });
      expect(got.error, register).toContain(hint);
      expect(got.queries).toEqual([]);
    }
  });

  it("неопубликованный регистр — подсказка о «Составе»", async () => {
    const got = await read({ register: "InformationRegister_НетТакого" });
    expect(got.error).toContain("Состав");
  });

  it("поле, которого нет, и поле-коллекция — отказ до запроса со ссылкой на describe_entity", async () => {
    const unknown = await read({ select: ["НетПоля"] });
    expect(unknown.error).toContain("describe_entity");
    expect(unknown.queries).toEqual([]);
    const nested = await read({ register: "Document_X", select: ["a"] });
    expect(nested.error).toContain("не регистр");
  });
});

describe("отборы: неприменимый — отказ, а не пропуск", () => {
  it("organization и from/to превращаются в условия запроса", async () => {
    const { params } = await read({
      register: SETTINGS,
      select: ["Period"],
      organization: "Организация",
      from: "2030-01-01",
      to: "2030-03-31",
    });
    const filter = params!.get("$filter")!;
    expect(filter).toContain(`Организация_Key eq guid'${REF}'`);
    expect(filter).toContain("Period ge datetime'2030-01-01T00:00:00'");
    expect(filter).toContain("Period le datetime'2030-03-31T23:59:59'");
  });

  it("нет поля Period — отказ, строки не читаются", async () => {
    const got = await read({
      register: "InformationRegister_РасчетНалогаУСН",
      select: ["ПериодРасчета"],
      from: "2030-01-01",
    });
    expect(got.error).toContain("нет поля «Period»");
    expect(got.params).toBeUndefined();
  });

  it("литералы по типу поля: дата целым днём, число, ссылка", async () => {
    const { params } = await read({
      filters: [
        { field: "СрокУплаты", op: "le", value: "2030-03-31" },
        { field: "СрокУплаты", op: "ge", value: "2030-01-01" },
        { field: "Начислено", op: "gt", value: 0 },
        { field: "СчетУчета_Key", op: "eq", value: REF },
      ],
    });
    expect(params!.get("$filter")).toBe(
      `СрокУплаты le datetime'2030-03-31T23:59:59' and СрокУплаты ge datetime'2030-01-01T00:00:00' and ` +
        `Начислено gt 0 and СчетУчета_Key eq guid'${REF}'`,
    );
  });

  it("дата без времени в «равно» и «не равно» — весь день, а не полночь", async () => {
    const filter = async (op: string, value = "2030-03-31") =>
      (
        await read({
          filters: [
            { field: "СрокУплаты", op, value },
            { field: "Начислено", op: "gt", value: 0 },
          ],
        })
      ).params!.get("$filter");
    await expect(filter("eq")).resolves.toBe(
      "(СрокУплаты ge datetime'2030-03-31T00:00:00' and СрокУплаты le datetime'2030-03-31T23:59:59') and Начислено gt 0",
    );
    await expect(filter("ne")).resolves.toBe(
      "(СрокУплаты lt datetime'2030-03-31T00:00:00' or СрокУплаты gt datetime'2030-03-31T23:59:59') and Начислено gt 0",
    );
    // С временем — ровно этот момент.
    await expect(filter("eq", "2030-03-31T12:00:00")).resolves.toBe(
      "СрокУплаты eq datetime'2030-03-31T12:00:00' and Начислено gt 0",
    );
  });

  it("составная ссылка — только с типом, через cast", async () => {
    const filters = [{ field: "Recorder", op: "eq", value: REF }];
    const bare = await read({ filters });
    expect(bare.error).toContain("type");
    expect(bare.params).toBeUndefined();
    const typed = await read({ filters: [{ ...filters[0], type: "Document_РегламентнаяОперация" }] });
    expect(typed.violations).toEqual([]);
    expect(typed.params!.get("$filter")).toContain(`cast(guid'${REF}', 'Document_РегламентнаяОперация')`);
    const bad = await read({ filters: [{ ...filters[0], type: "x') or true or ('" }] });
    expect(bad.error).toContain("type");
  });

  it("неподходящее значение или сравнение — отказ с причиной", async () => {
    for (const filter of [
      { field: "СрокУплаты", op: "eq", value: "31.03.2030" },
      { field: "СчетУчета_Key", op: "gt", value: REF },
      { field: "СчетУчета_Key", op: "eq", value: "' or true" },
      { field: "Начислено", op: "eq", value: "5" },
      { field: "Начислено", op: "contains", value: "5" },
    ]) {
      const got = await read({ filters: [filter] });
      expect(got.error, JSON.stringify(filter)).toContain("filters");
      expect(got.params).toBeUndefined();
    }
  });

  it("кавычка в строке не выходит из литерала", async () => {
    const { params } = await read({
      register: SETTINGS,
      select: ["Period"],
      filters: [{ field: "СистемаНалогообложения", op: "eq", value: "a' or Period ne null or 'b" }],
    });
    expect(params!.get("$filter")).toBe("СистемаНалогообложения eq 'a'' or Period ne null or ''b'");
  });
});

describe("страницы и предел", () => {
  const many = (n: number) => Array.from({ length: n }, (_, i) => row(i));

  it("один запрос с пределом без $skip; страницы режутся в коннекторе и не пересекаются", async () => {
    const first = await read({ limit: 10 }, many(25));
    expect(first.queries.filter((q) => q.base === FLAT)).toHaveLength(1);
    expect(first.params!.get("$top")).toBe("1000");
    expect(first.params!.get("$skip") ?? "0").toBe("0");
    expect(first.page).toMatchObject({ matched: 25, returned: 10, truncated: true, nextOffset: 10 });
    // 1С отдала те же строки в другом порядке — страницы от этого не меняются.
    const second = await read({ limit: 10, offset: 10, snapshot: first.page.snapshot }, many(25).reverse());
    const seen = [...first.page.rows, ...second.page.rows].map((r) => r["Начислено"]);
    expect(new Set(seen).size).toBe(20);
    expect(second.page.dataChanged).toBeUndefined();
  });

  it("изменилась любая строка — dataChanged, даже если их столько же", async () => {
    const first = await read({ limit: 2 }, many(5));
    const changed = many(5).map((r, i) => (i === 4 ? { ...r, Начислено: 100 } : r));
    const second = await read({ limit: 2, offset: 2, snapshot: first.page.snapshot }, changed);
    expect(second.page.dataChanged).toBe(true);
  });

  it("больше предела — без продолжения и с объяснением; offset — отказ", async () => {
    const { page } = await read({ limit: 10 }, many(1500));
    expect(page).toMatchObject({ matched: 1000, returned: 10, truncated: true });
    expect(page.nextOffset).toBeUndefined();
    expect(page.note).toContain("произвольная часть");
    const next = await read({ limit: 10, offset: 10 }, many(1500));
    expect(next.error).toContain("продолжения нет");
  });

  it("orderBy — сортирует 1С, поле добавляется в выборку, descending меняет порядок", async () => {
    const { page, params } = await read(
      { select: ["СрокУплаты"], orderBy: "Начислено", descending: true },
      many(3),
    );
    expect(params!.get("$orderby")).toBe("Начислено desc");
    expect(page.rows.map((r) => r["Начислено"])).toEqual([2, 1, 0]);
  });
});

describe("в Управлении торговлей регистры тоже читаются", () => {
  it("регистр склада", async () => {
    const got = await read(
      { register: "AccumulationRegister_ТоварыНаСкладах", select: ["Period", "ВНаличии"] },
      [{ Period: "2030-01-01T00:00:00", ВНаличии: 3 }],
      "ut11",
    );
    expect(got.error).toBeUndefined();
    expect(got.violations).toEqual([]);
    expect(got.base).toBe("AccumulationRegister_ТоварыНаСкладах_RecordType");
  });
});

describe("поля с паролями и токенами", () => {
  const register = "InformationRegister_Проверка";
  const read = (name: string) =>
    readRegisterRows(
      {
        behavior: { maxRows: 10 },
        available: async () => new Set([register]),
        getMetadata: async () => ({
          entities: new Map([
            [register, { entitySet: register, properties: [{ name, type: "Edm.String", nullable: true }] }],
          ]),
        }),
        client: { getCollection: async () => ({ value: [] }) },
      } as unknown as Connection,
      { register, select: [name] },
    );

  it("не читаются по имени", async () => {
    for (const name of [
      "ПарольДоступа",
      "Password",
      "СекретныйКлюч",
      "КлиентскийСекрет",
      "ClientSecret",
      "ТокенДоступа",
    ])
      await expect(read(name), name).rejects.toThrow("пароль или токен");
  });

  it("секретарь — должность, а не секрет", async () => {
    for (const name of ["Секретарь", "СекретарьКомиссии", "ФИОСекретаря", "Secretary"])
      await expect(read(name), name).resolves.toMatchObject({ rows: [] });
  });
});
