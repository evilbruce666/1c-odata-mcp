import { describe, expect, it, vi } from "vitest";
import { createServer } from "../src/mcp/server.js";
import {
  contractSettlementKind,
  pickSettlementAccounts,
  settlementSideOf,
} from "../src/odata/settlement-accounts.js";
import { emptySettlementWarning, withSettlementAccounts } from "../src/tools/write.js";

// Заглушки схемы в этом файле без регистра бухгалтерии, а запись в такую базу запрещена. Здесь проверяется сама
// запись в Бухгалтерию; запрет — в ut-compat.test.ts и fake-1c.test.ts на настоящих схемах.
vi.mock("../src/odata/configuration.js", async (original) => ({
  ...(await original<typeof import("../src/odata/configuration.js")>()),
  withoutAccounting: async () => false,
}));

// Форма записей и значения — как в регистре «Счета учёта расчётов с контрагентами» живой базы БП 3.0.
const E = "00000000-0000-0000-0000-000000000000";
const ORG = "5bc0b4f6-7054-11e6-a2f5-000c29b3b732";
const CP = "3228a150-cbc7-11f1-9afa-07e45db8596f";
const CTR = "32d61f60-cbc7-11f1-97e1-e1f082ae62fe";
const A = {
  "62.01": "a6201000-0000-4000-8000-000000000000",
  "62.02": "a6202000-0000-4000-8000-000000000000",
  "62.21": "a6221000-0000-4000-8000-000000000000",
  "62.22": "a6222000-0000-4000-8000-000000000000",
  "60.01": "a6001000-0000-4000-8000-000000000000",
  "60.02": "a6002000-0000-4000-8000-000000000000",
  "76.09": "a7609000-0000-4000-8000-000000000000",
  "76.02": "a7602000-0000-4000-8000-000000000000",
} as const;

const rec = (
  key: { org?: string; cp?: string; ctr?: string; kind?: string },
  acc: { buyer?: string; buyerAdv?: string; supplier?: string; supplierAdv?: string },
) => ({
  Организация_Key: key.org ?? E,
  Контрагент_Key: key.cp ?? E,
  Договор_Key: key.ctr ?? E,
  ВидРасчетовПоДоговору: key.kind ?? "",
  СчетУчетаРасчетовСПокупателем_Key: acc.buyer ?? E,
  СчетУчетаАвансовПолученных_Key: acc.buyerAdv ?? E,
  СчетУчетаРасчетовСПоставщиком_Key: acc.supplier ?? E,
  СчетУчетаАвансовВыданных_Key: acc.supplierAdv ?? E,
});
const RUB = "РасчетыВВалютеРегламентированногоУчета";
const GENERAL = [
  rec({}, { buyer: A["62.01"], buyerAdv: A["62.02"], supplier: A["60.01"], supplierAdv: A["60.02"] }),
  rec(
    { kind: RUB },
    { buyer: A["62.01"], buyerAdv: A["62.02"], supplier: A["60.01"], supplierAdv: A["60.02"] },
  ),
  rec({ kind: "РасчетыВИностраннойВалюте" }, { buyer: A["62.21"], buyerAdv: A["62.22"] }),
];
const key = { orgKey: ORG, counterpartyRef: CP, contractRef: CTR, kind: RUB } as const;

describe("подбор записи регистра счетов расчётов", () => {
  it("особая настройка договора важнее общей (настройка из формы договора — с видом «в валюте регл. учёта»)", () => {
    const rows = [
      ...GENERAL,
      rec({ org: ORG, cp: CP, ctr: CTR, kind: RUB }, { buyer: A["76.09"], buyerAdv: A["76.02"] }),
    ];
    expect(pickSettlementAccounts(rows, key, "customer")).toMatchObject({
      settle: A["76.09"],
      advance: A["76.02"],
      level: "договор",
    });
  });

  it("без особой настройки — общая запись своего вида расчётов", () => {
    expect(pickSettlementAccounts(GENERAL, key, "customer")).toMatchObject({
      settle: A["62.01"],
      level: "общая",
    });
    expect(
      pickSettlementAccounts(GENERAL, { ...key, kind: "РасчетыВИностраннойВалюте" }, "customer"),
    ).toMatchObject({
      settle: A["62.21"],
      advance: A["62.22"],
    });
  });

  it("запись чужого договора или другого вида расчётов не подходит", () => {
    const other = "99999999-9999-4999-8999-999999999999";
    const rows = [
      ...GENERAL,
      rec({ ctr: other }, { buyer: A["76.09"] }),
      rec({ ctr: CTR, kind: "РасчетыВИностраннойВалюте" }, { buyer: A["76.09"] }),
    ];
    expect(pickSettlementAccounts(rows, key, "customer")?.settle).toBe(A["62.01"]);
  });

  it("каждый счёт — из самой частной записи, где он заполнен", () => {
    const rows = [...GENERAL, rec({ ctr: CTR }, { buyer: A["76.09"] })];
    expect(pickSettlementAccounts(rows, key, "customer")).toMatchObject({
      settle: A["76.09"],
      advance: A["62.02"],
      level: "договор",
    });
  });

  it("поставщик берёт ресурсы поставщика", () => {
    expect(pickSettlementAccounts(GENERAL, key, "supplier")).toMatchObject({
      settle: A["60.01"],
      advance: A["60.02"],
    });
  });

  it("подходящих записей нет — undefined (останутся счета по умолчанию)", () => {
    expect(pickSettlementAccounts([], key, "customer")).toBeUndefined();
  });

  it("вид расчётов и сторона по договору", () => {
    expect(contractSettlementKind({ РасчетыВУсловныхЕдиницах: true }, "840", "643")).toBe(
      "РасчетыВУсловныхЕдиницах",
    );
    expect(contractSettlementKind({}, "643", "643")).toBe(RUB);
    expect(contractSettlementKind({}, "840", "643")).toBe("РасчетыВИностраннойВалюте");
    expect(contractSettlementKind({}, "933", "933")).toBe(RUB);
    expect(settlementSideOf("СПокупателем")).toBe("customer");
    expect(settlementSideOf("СПоставщиком")).toBe("supplier");
    expect(settlementSideOf("Прочее")).toBeUndefined();
  });
});

const REG = "InformationRegister_СчетаУчетаРасчетовСКонтрагентами";
const ACT = "Document_РеализацияТоваровУслуг";
const PAY = "Document_ПоступлениеНаРасчетныйСчет";

const connWith = (registerRows: unknown[] | undefined, currency = "643") => {
  const getCollection = vi.fn(async (path: string) => {
    if (path.startsWith(REG)) return { value: registerRows ?? [] };
    if (path.startsWith("ChartOfAccounts_Хозрасчетный")) {
      return { value: Object.entries(A).map(([Code, Ref_Key]) => ({ Ref_Key, Code, Description: Code })) };
    }
    return { value: [] };
  });
  const getEntity = vi.fn(async (path: string) =>
    path.startsWith("Catalog_Контрагенты")
      ? { Ref_Key: CP, Parent_Key: E, IsFolder: false, DeletionMark: false }
      : path.startsWith("Catalog_Валюты")
        ? { Code: currency }
        : {
            Ref_Key: CTR,
            Owner_Key: CP,
            Организация_Key: ORG,
            ВидДоговора: "СПокупателем",
            ВалютаВзаиморасчетов_Key: "c0000000-0000-4000-8000-000000000000",
            РасчетыВУсловныхЕдиницах: false,
          },
  );
  const available = new Set([
    "Catalog_ДоговорыКонтрагентов",
    "Catalog_Контрагенты",
    "Catalog_Валюты",
    "ChartOfAccounts_Хозрасчетный",
    ...(registerRows ? [REG] : []),
  ]);
  const props = (names: string[]) => ({ properties: names.map((name) => ({ name })) });
  return {
    getCollection,
    conn: {
      cfg: { name: "default" },
      behavior: { pageSize: 100, maxRows: 1000, analyticsMaxRows: 200_000 },
      available: async () => available,
      getMetadata: async () => ({
        entities: new Map([
          [
            "Catalog_Контрагенты",
            {
              properties: ["Ref_Key", "Parent_Key", "IsFolder", "DeletionMark"].map((name) => ({
                name,
                type: name.endsWith("_Key") ? "Edm.Guid" : "Edm.Boolean",
              })),
            },
          ],
          [ACT, props(["СчетУчетаРасчетовСКонтрагентом_Key", "СчетУчетаРасчетовПоАвансам_Key"])],
          [PAY, props(["СчетУчетаРасчетовСКонтрагентом_Key"])],
          [
            `${PAY}_РасшифровкаПлатежа`,
            props(["СуммаПлатежа", "СуммаВзаиморасчетов", "КурсВзаиморасчетов", "КратностьВзаиморасчетов"]),
          ],
        ]),
      }),
      client: { getCollection, getEntity },
    } as never,
  };
};

describe("счета расчётов в документе — из настройки, а не зашитые", () => {
  const withSpecial = [
    ...GENERAL,
    rec({ org: ORG, cp: CP, ctr: CTR, kind: RUB }, { buyer: A["76.09"], buyerAdv: A["76.02"] }),
  ];
  const act = {
    Организация_Key: ORG,
    Контрагент_Key: CP,
    ДоговорКонтрагента_Key: CTR,
    СчетУчетаРасчетовСКонтрагентом_Key: A["62.01"],
  };

  it("реализация: зашитый 62.01 заменяется настройкой договора, авансы — тоже; предпросмотр говорит откуда", async () => {
    const { conn, getCollection } = connWith(withSpecial);
    const r = await withSettlementAccounts(conn, ACT, act);
    expect(r.payload["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["76.09"]);
    expect(r.payload["СчетУчетаРасчетовПоАвансам_Key"]).toBe(A["76.02"]);
    expect(r.notes[0]).toContain("76.09 / 76.02 — из настройки договора");
    const url = decodeURIComponent(getCollection.mock.calls.find((c) => c[0].startsWith(REG))![0]);
    expect(url).toContain(`Договор_Key eq guid'${CTR}'`);
    expect(url).toContain(`Договор_Key eq guid'${E}'`);
  });

  it("оплата: счета ставятся только в строки расшифровки по этому договору", async () => {
    const { conn } = connWith(withSpecial);
    const r = await withSettlementAccounts(conn, PAY, {
      ВидОперации: "ОплатаПокупателя",
      Организация_Key: ORG,
      ДоговорКонтрагента_Key: CTR,
      РасшифровкаПлатежа: [
        {
          ДоговорКонтрагента_Key: CTR,
          СчетУчетаРасчетовСКонтрагентом_Key: A["62.01"],
          СчетУчетаРасчетовПоАвансам_Key: A["62.02"],
        },
      ],
    });
    const row = (r.payload["РасшифровкаПлатежа"] as Record<string, unknown>[])[0]!;
    expect(row["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["76.09"]);
    expect(row["СчетУчетаРасчетовПоАвансам_Key"]).toBe(A["76.02"]);
    expect(r.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
  });

  it.each([
    [PAY, "ОплатаПокупателя", "customer"],
    ["Document_СписаниеСРасчетногоСчета", "ОплатаПоставщику", "supplier"],
    ["Document_ПриходныйКассовыйОрдер", "ОплатаПокупателя", "customer"],
    ["Document_РасходныйКассовыйОрдер", "ОплатаПоставщику", "supplier"],
  ] as const)("%s: настройка заполняет строки, не создавая шапку", async (set, op, side) => {
    const { conn } = connWith(GENERAL);
    // В метаданных каждой денежной сущности шапка существует: это не основание её заполнять.
    const c = conn as unknown as {
      getMetadata: () => Promise<{ entities: Map<string, { properties: { name: string }[] }> }>;
    };
    const meta = await c.getMetadata();
    meta.entities.set(set, meta.entities.get(PAY)!);
    c.getMetadata = async () => meta;
    const row = { ДоговорКонтрагента_Key: CTR, СуммаПлатежа: 1500 };
    const r = await withSettlementAccounts(conn, set, {
      Организация_Key: ORG,
      ДоговорКонтрагента_Key: CTR,
      ВидОперации: op,
      РасшифровкаПлатежа: [row],
    });
    expect(r.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
    expect(r.payload["РасшифровкаПлатежа"]).toMatchObject([
      {
        СчетУчетаРасчетовСКонтрагентом_Key: A[side === "customer" ? "62.01" : "60.01"],
        СчетУчетаРасчетовПоАвансам_Key: A[side === "customer" ? "62.02" : "60.02"],
      },
    ]);
    expect(row).toEqual({ ДоговорКонтрагента_Key: CTR, СуммаПлатежа: 1500 });
  });

  it("строгие значения без договора заполняют строку, не добавляя шапку", async () => {
    const { conn } = connWith(undefined);
    const c = conn as unknown as { cfg: object };
    c.cfg = {
      name: "default",
      accounting: {
        strictDefaults: true,
        defaults: { receivable: ["62.01"], receivableAdvance: ["62.02"] },
      },
    };
    const r = await withSettlementAccounts(conn, PAY, {
      ВидОперации: "ОплатаПокупателя",
      РасшифровкаПлатежа: [{ СуммаПлатежа: 10 }],
    });
    expect(r.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
    expect(r.payload["РасшифровкаПлатежа"]).toMatchObject([
      {
        СчетУчетаРасчетовСКонтрагентом_Key: A["62.01"],
        СчетУчетаРасчетовПоАвансам_Key: A["62.02"],
      },
    ]);
  });

  it("явные суммы и чужая строка сохраняются, явная шапка не переписывается", async () => {
    const { conn } = connWith(withSpecial);
    const other = { ДоговорКонтрагента_Key: E, СчетУчетаРасчетовСКонтрагентом_Key: A["60.01"] };
    const r = await withSettlementAccounts(conn, PAY, {
      ВидОперации: "ОплатаПокупателя",
      ДоговорКонтрагента_Key: CTR,
      СчетУчетаРасчетовСКонтрагентом_Key: A["60.01"],
      РасшифровкаПлатежа: [
        {
          ДоговорКонтрагента_Key: CTR,
          СуммаПлатежа: 1500,
          СуммаВзаиморасчетов: 123,
          КурсВзаиморасчетов: 2,
          КратностьВзаиморасчетов: 10,
        },
        other,
      ],
    });
    expect(r.payload["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["60.01"]);
    expect(r.payload["РасшифровкаПлатежа"]).toMatchObject([
      {
        СчетУчетаРасчетовСКонтрагентом_Key: A["76.09"],
        СуммаВзаиморасчетов: 123,
        КурсВзаиморасчетов: 2,
        КратностьВзаиморасчетов: 10,
      },
      other,
    ]);
  });

  it.each(["ПеречислениеНалога", "ПеречислениеЗП", "ПереводНаДругойСчет", "ПрочееСписание"])(
    "%s: шапка без платёжной расшифровки остаётся доступной",
    async (op) => {
      const { conn } = connWith(withSpecial);
      const r = await withSettlementAccounts(conn, PAY, {
        ВидОперации: op,
        ДоговорКонтрагента_Key: CTR,
      });
      expect(r.payload["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["76.09"]);
    },
  );

  it("не денежный документ с расшифровкой сохраняет заполнение шапки", async () => {
    const { conn } = connWith(withSpecial);
    const r = await withSettlementAccounts(conn, ACT, {
      ДоговорКонтрагента_Key: CTR,
      РасшифровкаПлатежа: [{ ДоговорКонтрагента_Key: CTR }],
    });
    expect(r.payload["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["76.09"]);
  });

  const payment = {
    ВидОперации: "ОплатаПокупателя",
    Организация_Key: ORG,
    ДоговорКонтрагента_Key: CTR,
    РасшифровкаПлатежа: [{ ДоговорКонтрагента_Key: CTR, СуммаПлатежа: 1500 }],
  };

  it("оплата по рублёвому договору: сумма взаиморасчётов = сумме платежа, курс 1 (иначе 1С проводит без проводок)", async () => {
    const { conn } = connWith(undefined);
    const r = await withSettlementAccounts(conn, PAY, payment);
    expect((r.payload["РасшифровкаПлатежа"] as Record<string, unknown>[])[0]).toMatchObject({
      СуммаВзаиморасчетов: 1500,
      КурсВзаиморасчетов: 1,
      КратностьВзаиморасчетов: 1,
    });
  });

  it("валютный договор: курс не угадываем — сумму взаиморасчётов не заполняем и говорим об этом", async () => {
    const { conn } = connWith(GENERAL, "840");
    const r = await withSettlementAccounts(conn, PAY, payment);
    const row = (r.payload["РасшифровкаПлатежа"] as Record<string, unknown>[])[0]!;
    expect(r.payload).not.toHaveProperty("СчетУчетаРасчетовСКонтрагентом_Key");
    expect(row["СуммаВзаиморасчетов"]).toBeUndefined();
    expect(row["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["62.21"]);
    expect(r.notes.join(" ")).toContain("курс MCP не заполняет");
  });

  it("регистр не опубликован — счета по умолчанию остаются, предпросмотр это говорит", async () => {
    const { conn } = connWith(undefined);
    const r = await withSettlementAccounts(conn, ACT, act);
    expect(r.payload).toEqual(act);
    expect(r.notes[0]).toContain("по умолчанию (62.01)");
    expect(r.notes[0]).toContain("не опубликован");
  });

  it("без договора ничего не подбираем и регистр не читаем", async () => {
    const { conn, getCollection } = connWith(withSpecial);
    const { ДоговорКонтрагента_Key: _drop, ...noContract } = act;
    expect(await withSettlementAccounts(conn, ACT, noContract)).toEqual({ payload: noContract, notes: [] });
    expect(getCollection).not.toHaveBeenCalled();
  });

  it("после записи: пустой счёт расчётов при договоре — предупреждение; без договора — нет", () => {
    const created = {
      СчетУчетаРасчетовСКонтрагентом_Key: E,
      РасшифровкаПлатежа: [{ СчетУчетаРасчетовСКонтрагентом_Key: E }],
    };
    expect(emptySettlementWarning(created, { ДоговорКонтрагента_Key: CTR })).toContain("пуст счёт расчётов");
    expect(emptySettlementWarning(created, {})).toBeUndefined();
    const zeroSum = {
      СчетУчетаРасчетовСКонтрагентом_Key: A["62.01"],
      РасшифровкаПлатежа: [
        {
          ДоговорКонтрагента_Key: CTR,
          СчетУчетаРасчетовСКонтрагентом_Key: A["62.01"],
          СуммаВзаиморасчетов: 0,
        },
      ],
    };
    expect(emptySettlementWarning(zeroSum, { ДоговорКонтрагента_Key: CTR })).toContain(
      "сумма взаиморасчётов",
    );
    expect(
      emptySettlementWarning(
        { СчетУчетаРасчетовСКонтрагентом_Key: A["62.01"] },
        { ДоговорКонтрагента_Key: CTR },
      ),
    ).toBeUndefined();
  });
});

describe("copy_document: счета расчётов — как в исходном документе", () => {
  type Handler = (
    args: Record<string, unknown>,
    extra: Record<string, unknown>,
  ) => Promise<{
    structuredContent?: Record<string, unknown>;
  }>;
  const source = {
    Ref_Key: "50000000-0000-4000-8000-000000000000",
    Number: "0000-000010",
    Организация_Key: ORG,
    Контрагент_Key: CP,
    ДоговорКонтрагента_Key: CTR,
    ВидОперации: "Услуги",
    // Бухгалтер поставил в исходном документе особый счёт вручную.
    СчетУчетаРасчетовСКонтрагентом_Key: A["76.09"],
    СчетУчетаРасчетовПоАвансам_Key: A["76.02"],
    Услуги: [{ LineNumber: 1, Сумма: 1, Количество: 1, Цена: 1 }],
  };
  const copyTool = (registerRows: unknown[]) => {
    const { conn } = connWith(registerRows);
    const c = conn as unknown as {
      cfg: Record<string, unknown>;
      behavior: Record<string, unknown>;
      available: () => Promise<Set<string>>;
      client: Record<string, unknown>;
    };
    c.cfg = { name: "default", writable: true };
    c.behavior = { pageSize: 100, maxRows: 1000, analyticsMaxRows: 200_000, writeOperationMarker: false };
    const avail = c.available;
    c.available = async () => new Set([...(await avail()), ACT]);
    const getEntity = c.client["getEntity"] as (p: string) => Promise<unknown>;
    c.client["getEntity"] = async (p: string) => (p.startsWith(ACT) ? { ...source } : getEntity(p));
    c.client["prepareCreate"] = vi.fn(async () => undefined);
    const server = createServer({
      db: () => conn,
      databases: () => [],
      behavior: {},
    } as never) as unknown as {
      _registeredTools: Record<string, { handler: Handler }>;
    };
    return server._registeredTools["write.document.copy_document"]!.handler;
  };
  // Общая настройка в регистре — 62.01/62.02: если бы копия её применила, счёт 76.09 потерялся бы.
  const generalOnly = GENERAL;

  it("копия сохраняет счета исходного документа и говорит об этом", async () => {
    const r = await copyTool(generalOnly)({ database: "default", entitySet: ACT, ref: source.Ref_Key }, {});
    const sc = r.structuredContent!;
    const payload = sc["payload"] as Record<string, unknown>;
    expect(payload["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["76.09"]);
    expect(payload["СчетУчетаРасчетовПоАвансам_Key"]).toBe(A["76.02"]);
    expect((sc["notes"] as string[]).join(" ")).toContain("как в исходном документе");
  });

  it("в копии сменили договор — счета берутся из настройки нового договора", async () => {
    const other = "77777777-7777-4777-8777-777777777777";
    const r = await copyTool(generalOnly)(
      { database: "default", entitySet: ACT, ref: source.Ref_Key, fields: { ДоговорКонтрагента_Key: other } },
      {},
    );
    const payload = r.structuredContent!["payload"] as Record<string, unknown>;
    expect(payload["СчетУчетаРасчетовСКонтрагентом_Key"]).toBe(A["62.01"]);
  });
});

describe("MCP ancestor settlement policy", () => {
  const near = "77777777-7777-4777-8777-777777777777";
  const far = "88888888-8888-4888-8888-888888888888";
  it("chooses individual and contract overrides, then nearest current group separately per account", () => {
    const rows = [
      ...GENERAL,
      rec({ cp: far, org: ORG }, { buyer: A["60.01"], buyerAdv: A["60.02"] }),
      rec({ cp: near }, { buyer: A["76.09"], buyerAdv: A["76.02"] }),
      rec({ cp: CP }, { buyer: A["62.21"] }),
      rec({ cp: CP, ctr: CTR }, { buyerAdv: A["62.22"] }),
    ];
    const k = { ...key, groupRefs: [near, far] };
    expect(pickSettlementAccounts(rows, k, "customer")).toMatchObject({
      settle: A["62.21"],
      advance: A["62.22"],
    });
    const choice = pickSettlementAccounts(rows.slice(0, -2), k, "customer");
    expect(choice).toMatchObject({
      settle: A["76.09"],
      advance: A["76.02"],
      level: "группа",
      provenance: {
        settle: { level: "группа", counterpartyRef: near, groupDepth: 0 },
        advance: { level: "группа", counterpartyRef: near, groupDepth: 0 },
      },
    });
  });
});
