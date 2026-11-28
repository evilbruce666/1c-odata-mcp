import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import { REGISTERS, TAX_PATENT_DOCUMENTS } from "../src/config/mapping.js";
import type { Connection } from "../src/context.js";
import { ODataError } from "../src/odata/errors.js";
import { TAX_REGIME_KEYS, readTaxRegime } from "../src/odata/tax-regime.js";
import { registerOrganizationTools } from "../src/tools/organization.js";
import { bp30Metadata, ut11Metadata } from "./helpers/metadata-snapshot.js";

const ORG = "11111111-1111-4111-8111-111111111111";
const SETTINGS = REGISTERS.taxSettings[0];
const CALC = REGISTERS.usnCalculation[0];
const PATENT = TAX_PATENT_DOCUMENTS[0];
const AS_OF = "2030-05-10";

type Rows = Record<string, Array<Record<string, unknown>> | Error>;

/** Подключение с настоящей схемой БП 3.0 (без hidden-объектов) и своими строками по объектам. */
function connection(
  rows: Rows,
  options: { hidden?: string[]; metadata?: ReturnType<typeof bp30Metadata> } = {},
) {
  const metadata = options.metadata ?? bp30Metadata();
  const queries: Array<{ set: string; params: URLSearchParams }> = [];
  const getCollection = vi.fn(async (path: string) => {
    const [set = "", query = ""] = path.split("?");
    const params = new URLSearchParams(query);
    queries.push({ set: decodeURIComponent(set), params });
    const own = rows[decodeURIComponent(set)] ?? [];
    if (own instanceof Error) throw own;
    const skip = Number(params.get("$skip") ?? 0);
    return { value: own.slice(skip, skip + Number(params.get("$top") ?? own.length)) };
  });
  const conn = {
    cfg: { name: "test" },
    behavior: { pageSize: 100, maxRows: 1000 },
    getMetadata: async () => metadata,
    available: async () => new Set([...metadata.entities.keys()].filter((k) => !options.hidden?.includes(k))),
    client: { getCollection },
  } as unknown as Connection;
  return { conn, queries };
}

const setting = (extra: Record<string, unknown> = {}) => ({
  Period: "2020-03-01T00:00:00",
  СистемаНалогообложения: "Упрощенная",
  ОбъектНалогообложенияУСН: "Доходы",
  ПрименяетсяУСНПатент: false,
  ПрименяетсяАУСН: false,
  ПлательщикНДСПрименяющийУСН: false,
  ...extra,
});

describe("режим по настройкам 1С", () => {
  it("УСН «доходы» и патент — два ключа, дата настроек и значение 1С", async () => {
    const { conn, queries } = connection({ [SETTINGS]: [setting({ ПрименяетсяУСНПатент: true })] });
    await expect(readTaxRegime(conn, ORG, AS_OF)).resolves.toEqual({
      regime: {
        regimes: ["usn_income", "psn"],
        source: "settings",
        asOf: AS_OF,
        system: "Упрощенная",
        settingsFrom: "2020-03-01",
        vatPayerOnUsn: false,
      },
      notes: [],
    });
    expect(queries).toHaveLength(1);
  });

  it("объект налогообложения и патент меняют ответ", async () => {
    const regimes = async (extra: Record<string, unknown>) =>
      (await readTaxRegime(connection({ [SETTINGS]: [setting(extra)] }).conn, ORG, AS_OF)).regime?.regimes;
    await expect(regimes({})).resolves.toEqual(["usn_income"]);
    await expect(regimes({ ОбъектНалогообложенияУСН: "ДоходыМинусРасходы" })).resolves.toEqual([
      "usn_income_expense",
    ]);
  });

  it("берёт последнюю запись не позже даты: будущая смена режима ещё не действует", async () => {
    const { conn, queries } = connection({ [SETTINGS]: [setting()] });
    await readTaxRegime(conn, ORG, AS_OF);
    const q = queries[0]!.params;
    expect(q.get("$filter")).toContain(`Period le datetime'${AS_OF}T23:59:59'`);
    expect(q.get("$filter")).toContain(ORG);
    expect(q.get("$orderby")).toBe("Period desc");
  });

  it("не УСН — значение 1С без придуманных ключей", async () => {
    for (const system of ["Общая", "НалогНаПрофессиональныйДоход", "АУСН"]) {
      const { conn } = connection({ [SETTINGS]: [setting({ СистемаНалогообложения: system })] });
      const found = await readTaxRegime(conn, ORG, AS_OF);
      expect(found.regime, system).toMatchObject({ regimes: [], system });
      expect(found.regime).not.toHaveProperty("vatPayerOnUsn");
      expect(found.notes.join(" ")).toContain("только для УСН");
    }
  });

  it("регистр опубликован, записей нет — режим не угадывается по косвенным признакам", async () => {
    const { conn, queries } = connection({ [CALC]: [{ ПериодРасчета: "2030-03-31T00:00:00" }] });
    const found = await readTaxRegime(conn, ORG, AS_OF);
    expect(found.regime).toBeUndefined();
    expect(found.notes.join(" ")).toContain("нет записей");
    expect(queries.map((q) => q.set)).toEqual([SETTINGS]);
  });

  it("запрашивает только поля, которые есть в схеме этого релиза", async () => {
    const full = bp30Metadata();
    const entity = full.entities.get(SETTINGS)!;
    const entities = new Map(full.entities);
    entities.set(SETTINGS, {
      ...entity,
      properties: entity.properties.filter((p) => p.name !== "ПрименяетсяАУСН"),
    });
    const { conn, queries } = connection({ [SETTINGS]: [setting()] }, { metadata: { ...full, entities } });
    await readTaxRegime(conn, ORG, AS_OF);
    expect(queries[0]!.params.get("$select")).not.toContain("ПрименяетсяАУСН,");
    expect(queries[0]!.params.get("$select")).toContain("СистемаНалогообложения");
  });
});

describe("режим косвенно, когда регистр настроек не опубликован", () => {
  const hidden = [SETTINGS];
  const calc = (period: string, object = "Доходы") => ({
    ПериодРасчета: `${period}T00:00:00`,
    ОбъектНалогообложения: object,
  });

  it("объект из расчёта УСН, патент из действующей операции, период и совет опубликовать", async () => {
    const { conn, queries } = connection(
      { [CALC]: [calc("2030-03-31")], [PATENT]: [{ Ref_Key: ORG }] },
      { hidden },
    );
    const found = await readTaxRegime(conn, ORG, AS_OF);
    expect(found.regime).toEqual({
      regimes: ["usn_income", "psn"],
      source: "inferred",
      asOf: AS_OF,
      calculatedFor: "2030-03-31",
    });
    expect(found.notes.join(" ")).toContain("«Система налогообложения»");
    const patent = queries.find((q) => q.set === PATENT)!.params.get("$filter")!;
    expect(patent).toContain("Posted eq true");
    expect(patent).toContain(`ДатаНачала le datetime'${AS_OF}T23:59:59'`);
    expect(patent).toContain(`ДатаОкончания ge datetime'${AS_OF}T00:00:00'`);
  });

  it("следов нет — режим не назван, ОСНО не предполагается", async () => {
    const found = await readTaxRegime(connection({}, { hidden }).conn, ORG, AS_OF);
    expect(found.regime).toBeUndefined();
    expect(found.notes.join(" ")).toContain("не определён");
  });

  it("противоречие объектов за один период — ключ УСН не выводится", async () => {
    const rows = [calc("2030-03-31"), calc("2030-03-31", "ДоходыМинусРасходы")];
    const found = await readTaxRegime(connection({ [CALC]: rows }, { hidden }).conn, ORG, AS_OF);
    expect(found.regime).toBeUndefined();
    expect(found.notes.join(" ")).toContain("не определяется однозначно");
  });

  it("прежний объект в предыдущем периоде противоречием не считается", async () => {
    const rows = [calc("2030-03-31", "ДоходыМинусРасходы"), calc("2029-12-31")];
    const found = await readTaxRegime(connection({ [CALC]: rows }, { hidden }).conn, ORG, AS_OF);
    expect(found.regime?.regimes).toEqual(["usn_income_expense"]);
  });

  it("давний расчёт не выдаётся за нынешний режим", async () => {
    const found = await readTaxRegime(
      connection({ [CALC]: [calc("2027-12-31")] }, { hidden }).conn,
      ORG,
      AS_OF,
    );
    expect(found.regime).toBeUndefined();
    expect(found.notes.join(" ")).toContain("2027-12-31");
  });

  it("в Управлении торговлей без налогового учёта — ни режима, ни совета публиковать", async () => {
    const { conn, queries } = connection({}, { metadata: ut11Metadata() });
    await expect(readTaxRegime(conn, ORG, AS_OF)).resolves.toEqual({ notes: [] });
    expect(queries).toEqual([]);
  });
});

describe("словарь режимов и карточка", () => {
  it("каждый режим справочника сроков есть в словаре функции", () => {
    const rules = (
      JSON.parse(readFileSync("tax-calendar.json", "utf8")) as { rules: Array<{ regimes: string[] }> }
    ).rules;
    const unknown = rules
      .flatMap((r) => r.regimes)
      .filter((r) => !(TAX_REGIME_KEYS as readonly string[]).includes(r));
    expect(unknown).toEqual([]);
  });

  it("ошибка чтения режима не отменяет карточку: реквизиты есть, причина — в заметках", async () => {
    const denied = new ODataError({ kind: "auth", message: "Недостаточно прав" });
    const { conn } = connection({
      Catalog_Организации: [{ Ref_Key: ORG, Description: "Организация" }],
      [SETTINGS]: denied,
    });
    (conn.client as unknown as { getEntity: () => Promise<unknown> }).getEntity = async () => ({
      Description: "Организация",
      ИНН: "7700000000",
    });
    let handler!: (args: Record<string, unknown>) => Promise<{
      isError?: boolean;
      structuredContent?: { inn?: string; taxRegime?: unknown; notes?: string[] };
    }>;
    registerOrganizationTools(
      { registerTool: (_n: string, _c: unknown, h: typeof handler) => (handler = h) } as never,
      { db: () => conn } as never,
    );
    const result = await handler({});
    expect(result.isError).toBeFalsy();
    expect(result.structuredContent?.inn).toBe("7700000000");
    expect(result.structuredContent?.taxRegime).toBeUndefined();
    expect(result.structuredContent?.notes?.join(" ")).toContain("«Система налогообложения»");
    expect(result.structuredContent?.notes?.join(" ")).toContain("Недостаточно прав");
  });
});
