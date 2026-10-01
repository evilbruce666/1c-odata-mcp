import { describe, it, expect } from "vitest";
import type { Connection } from "../src/context.js";
import {
  ACCOUNT_FILTER_BATCH,
  BALANCE_AND_TURNOVERS,
  accountsOrgFilter,
  balanceAndTurnoversPath,
  nextDay,
  turnoversByAccounts,
} from "../src/odata/accounting.js";
import { AggregateOverflowError } from "../src/odata/aggregate.js";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import type { ServerContext } from "../src/context.js";
import { aggregateAccountTurnover, registerRegisterTools } from "../src/tools/registers.js";

const REG = "AccountingRegister_Хозрасчетный";
const guid = (n: number): string => `00000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const ORG = "11111111-2222-3333-4444-555555555555";

/**
 * Фейковое подключение: отдаёт строки виртуальной таблицы постранично ($top/$skip),
 * уважает фильтр по Account_Key из $filter и записывает все запрошенные пути.
 */
function fakeConn(rows: Array<Record<string, unknown>>, cap = 200_000) {
  const paths: string[] = [];
  const client = {
    async getCollection(path: string): Promise<{ value: unknown[] }> {
      paths.push(path);
      const dec = decodeURIComponent(path);
      const keys = [...dec.matchAll(/Account_Key eq guid'([^']+)'/g)].map((m) => m[1]);
      const top = Number(/\$top=(\d+)/.exec(path)?.[1] ?? rows.length);
      const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
      const pool = rows.filter((r) => keys.includes(String(r["Account_Key"])));
      return { value: pool.slice(skip, skip + top) };
    },
  };
  const conn = {
    client,
    behavior: { analyticsMaxRows: cap, pageSize: 100, maxRows: 1000 },
    available: async () => new Set([REG]),
  } as unknown as Connection;
  return { conn, paths };
}

const f = BALANCE_AND_TURNOVERS.fields;
/** Строка как её отдаёт 1С: все шесть ресурсов присутствуют (по умолчанию 0). */
const row = (acc: string, v: Partial<Record<keyof typeof f, number | string>>): Record<string, unknown> => ({
  Account_Key: acc,
  ...Object.fromEntries(Object.values(f).map((name) => [name, 0])),
  ...Object.fromEntries(Object.entries(v).map(([k, val]) => [f[k as keyof typeof f], val])),
});

describe("balanceAndTurnoversPath", () => {
  it("period — path-параметрами; EndPeriod = начало дня, следующего за to", () => {
    expect(balanceAndTurnoversPath(REG, "2026-08-01", "2026-08-31")).toBe(
      `${REG}/BalanceAndTurnovers(StartPeriod=datetime'2026-08-01T00:00:00',EndPeriod=datetime'2026-09-01T00:00:00')`,
    );
  });

  it("конец года переходит в 1 января следующего", () => {
    expect(balanceAndTurnoversPath(REG, "2025-01-01", "2025-12-31")).toContain(
      "EndPeriod=datetime'2026-01-01T00:00:00'",
    );
  });
});

describe("nextDay", () => {
  it.each([
    ["2026-08-15", "2026-08-16"],
    ["2026-08-31", "2026-09-01"],
    ["2026-04-30", "2026-05-01"],
    ["2025-12-31", "2026-01-01"],
    ["2025-02-28", "2025-03-01"], // не високосный
    ["2024-02-28", "2024-02-29"], // високосный
    ["2024-02-29", "2024-03-01"],
    ["2000-02-28", "2000-02-29"], // кратен 400 — високосный
    ["1900-02-28", "1900-03-01"], // кратен 100, не 400 — не високосный
    ["9999-01-31", "9999-02-01"],
  ])("%s → %s", (from, to) => {
    expect(nextDay(from)).toBe(to);
  });

  it("несуществующая дата — ошибка, а не «перекат» в следующий месяц", () => {
    expect(() => nextDay("2025-02-29")).toThrow(/Несуществующая дата/);
    expect(() => nextDay("2026-04-31")).toThrow(/Несуществующая дата/);
    expect(() => nextDay("2026-13-01")).toThrow(/Несуществующая дата/);
    expect(() => nextDay("2026-01-00")).toThrow(/Несуществующая дата/);
  });

  it("не зависит от таймзоны процесса", () => {
    const tz = process.env.TZ;
    try {
      for (const zone of ["UTC", "Pacific/Kiritimati", "Pacific/Pago_Pago", "Europe/Moscow"]) {
        process.env.TZ = zone;
        expect(nextDay("2026-08-31")).toBe("2026-09-01");
      }
    } finally {
      if (tz === undefined) delete process.env.TZ;
      else process.env.TZ = tz;
    }
  });
});

describe("accountsOrgFilter", () => {
  it("счета через OR, организация — на стороне 1С", () => {
    expect(accountsOrgFilter([guid(1), guid(2)], ORG)).toBe(
      `(Account_Key eq guid'${guid(1)}' or Account_Key eq guid'${guid(2)}') and Организация_Key eq guid'${ORG}'`,
    );
  });

  it("без организации — только счета", () => {
    expect(accountsOrgFilter([guid(1)])).toBe(`Account_Key eq guid'${guid(1)}'`);
  });

  it("кривой GUID не попадает в запрос", () => {
    expect(() => accountsOrgFilter(["x' or 1 eq 1"])).toThrow(/GUID/);
  });
});

describe("turnoversByAccounts", () => {
  it("строит запрос к BalanceAndTurnovers с фильтром счёта и организации, без $select", async () => {
    const { conn, paths } = fakeConn([row(guid(1), { turnoverDr: 10 })]);
    const { rows, meta } = await turnoversByAccounts(conn, [guid(1)], "2025-01-01", "2025-01-31", ORG);
    expect(rows).toHaveLength(1);
    expect(meta.rowsScanned).toBe(1);
    expect(paths).toHaveLength(1);
    const p = paths[0]!;
    expect(p.startsWith(balanceAndTurnoversPath(REG, "2025-01-01", "2025-01-31") + "?")).toBe(true);
    const dec = decodeURIComponent(p);
    expect(dec).toContain(`Организация_Key eq guid'${ORG}'`);
    expect(dec).toContain(`Account_Key eq guid'${guid(1)}'`);
    expect(p).not.toContain("$select");
  });

  it("листает все страницы (1500 строк > размера страницы), ничего не теряя", async () => {
    const rows = Array.from({ length: 1500 }, () => row(guid(1), { turnoverDr: 1 }));
    const { conn } = fakeConn(rows);
    const res = await turnoversByAccounts(conn, [guid(1)], "2025-01-01", "2025-12-31");
    expect(res.rows).toHaveLength(1500);
  });

  it("много субсчетов — режет $filter на пачки и собирает всё", async () => {
    const n = ACCOUNT_FILTER_BATCH + 5;
    const keys = Array.from({ length: n }, (_, i) => guid(i + 1));
    const { conn, paths } = fakeConn(keys.map((k) => row(k, { closingDr: 1 })));
    const res = await turnoversByAccounts(conn, keys, "2025-01-01", "2025-01-31");
    expect(paths).toHaveLength(2);
    expect(res.rows).toHaveLength(n);
  });

  it("переполнение потолка — явная ошибка, а не частичная ОСВ", async () => {
    const rows = Array.from({ length: 50 }, () => row(guid(1), { turnoverDr: 1 }));
    const { conn } = fakeConn(rows, 10);
    await expect(turnoversByAccounts(conn, [guid(1)], "2025-01-01", "2025-01-31")).rejects.toBeInstanceOf(
      AggregateOverflowError,
    );
  });

  it("переполнение суммарно по пачкам счетов тоже ловится", async () => {
    // каждая пачка ≤ потолка, но вместе — больше
    const keys = Array.from({ length: ACCOUNT_FILTER_BATCH + 1 }, (_, i) => guid(i + 1));
    const rows = keys.map((k) => row(k, { turnoverDr: 1 }));
    const { conn } = fakeConn(rows, ACCOUNT_FILTER_BATCH);
    await expect(turnoversByAccounts(conn, keys, "2025-01-01", "2025-01-31")).rejects.toBeInstanceOf(
      AggregateOverflowError,
    );
  });

  it("разбивка по периодам от 1С — ошибка (иначе остатки задвоятся)", async () => {
    const rows = [
      { ...row(guid(1), { openingDr: 1 }), Period: "2025-01-01T00:00:00" },
      { ...row(guid(1), { openingDr: 1 }), Period: "2025-02-01T00:00:00" },
    ];
    const { conn } = fakeConn(rows);
    await expect(turnoversByAccounts(conn, [guid(1)], "2025-01-01", "2025-02-28")).rejects.toThrow(
      /разбивку по периодам/,
    );
  });

  it("строка с неизвестными именами полей — ошибка, а не ОСВ из нулей", async () => {
    const alien = { Account_Key: guid(1), AmountOpeningBalanceDr: 123.45, Foo: "секрет" };
    const { conn } = fakeConn([alien]);
    const err = await turnoversByAccounts(conn, [guid(1)], "2025-01-01", "2025-01-31").catch((e: Error) => e);
    expect(err).toBeInstanceOf(Error);
    const msg = (err as Error).message;
    expect(msg).toContain("BalanceAndTurnovers");
    expect(msg).toContain("Account_Key, AmountOpeningBalanceDr, Foo"); // фактические ключи
    expect(msg).toContain("СуммаOpeningBalanceDr"); // чего не хватает
    expect(msg).not.toContain("123.45"); // значения не выводим
    expect(msg).not.toContain("секрет");
  });

  it("частично совпавшая структура (не все шесть ресурсов) — тоже ошибка", async () => {
    const partial = row(guid(1), { turnoverDr: 1 });
    delete partial["СуммаClosingBalanceCr"];
    const { conn } = fakeConn([partial]);
    await expect(turnoversByAccounts(conn, [guid(1)], "2025-01-01", "2025-01-31")).rejects.toThrow(
      /нет полей СуммаClosingBalanceCr/,
    );
  });

  it("пустой набор счетов — без запроса к 1С", async () => {
    const { conn, paths } = fakeConn([]);
    const res = await turnoversByAccounts(conn, [], "2025-01-01", "2025-01-31");
    expect(res.rows).toEqual([]);
    expect(paths).toHaveLength(0);
  });
});

describe("aggregateAccountTurnover", () => {
  const accounts = [
    { key: guid(1), code: "60.01", description: "Расчеты с поставщиками" },
    { key: guid(2), code: "60.02", description: "Авансы выданные" },
    { key: guid(3), code: "60.03", description: "Без движений" },
  ];

  it("Dr/Cr копятся раздельно (развёрнуто) и по субсчетам, в порядке плана счетов", () => {
    const rows = [
      row(guid(2), { openingDr: 50, turnoverCr: 50 }),
      row(guid(1), { openingCr: 100, turnoverDr: 70, turnoverCr: 20, closingCr: 50 }),
      row(guid(1), { openingDr: 30, turnoverDr: 5, closingDr: 35 }),
    ];
    const r = aggregateAccountTurnover(rows, accounts);
    expect(r.total).toEqual({
      openingDr: 8000,
      openingCr: 10000,
      turnoverDr: 7500,
      turnoverCr: 7000,
      closingDr: 3500,
      closingCr: 5000,
    });
    expect(r.byAccount.map((a) => a.account.code)).toEqual(["60.01", "60.02"]);
    expect(r.byAccount[0]!.sums.openingDr).toBe(3000);
    expect(r.consistent).toBe(true);
  });

  it("суммирует в копейках без float-дрейфа; строки-числа OData тоже", () => {
    const rows = Array.from({ length: 1000 }, () => row(guid(1), { turnoverDr: 0.1, turnoverCr: "0.2" }));
    const r = aggregateAccountTurnover(rows, accounts);
    expect(r.total.turnoverDr).toBe(10000); // 100.00 ₽ ровно
    expect(r.total.turnoverCr).toBe(20000);
  });

  it("нарушение Сн + ОбДт − ОбКт = Ск видно во флаге consistent", () => {
    const r = aggregateAccountTurnover([row(guid(1), { openingDr: 10, closingDr: 5 })], accounts);
    expect(r.consistent).toBe(false);
  });

  it("нет строк — нули", () => {
    const r = aggregateAccountTurnover([], accounts);
    expect(r.total.closingDr).toBe(0);
    expect(r.byAccount).toEqual([]);
    expect(r.consistent).toBe(true);
  });
});

// ─── Инструмент целиком: через MCP-клиент в памяти (валидирует и outputSchema) ───

const CHART = "ChartOfAccounts_Хозрасчетный";
const ORGS = "Catalog_Организации";

function fakeCtx(regRows: Array<Record<string, unknown>>) {
  const paths: string[] = [];
  const chart = [
    { Ref_Key: guid(1), Code: "60.01", Description: "Расчеты с поставщиками" },
    { Ref_Key: guid(2), Code: "60.02", Description: "Авансы выданные" },
  ];
  const client = {
    async getCollection(path: string): Promise<{ value: unknown[] }> {
      paths.push(path);
      const dec = decodeURIComponent(path);
      const top = Number(/\$top=(\d+)/.exec(path)?.[1] ?? 1000);
      const skip = Number(/\$skip=(\d+)/.exec(path)?.[1] ?? 0);
      let pool: unknown[] = [];
      if (path.startsWith(CHART)) {
        const pfx = /startswith\(Code, '([^']*)'\)/.exec(dec)?.[1] ?? "";
        pool = chart.filter((a) => a.Code.startsWith(pfx));
      } else if (path.startsWith(ORGS)) {
        pool = dec.includes("substringof('Ромашка', Description)")
          ? [{ Ref_Key: ORG, Description: "ООО Ромашка" }]
          : [];
      } else if (path.startsWith(REG)) {
        const keys = [...dec.matchAll(/Account_Key eq guid'([^']+)'/g)].map((m) => m[1]);
        pool = regRows.filter((r) => keys.includes(String(r["Account_Key"])));
      }
      return { value: pool.slice(skip, skip + top) };
    },
  };
  const conn = {
    cfg: { name: "main" },
    client,
    behavior: { analyticsMaxRows: 200_000, pageSize: 100, maxRows: 1000 },
    available: async () => new Set([REG, CHART, ORGS]),
  };
  const ctx = { db: () => conn } as unknown as ServerContext;
  return { ctx, paths };
}

async function callTool(ctx: ServerContext, args: Record<string, unknown>) {
  const server = new McpServer({ name: "t", version: "0" });
  registerRegisterTools(server, ctx);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: "c", version: "0" });
  await client.connect(b);
  try {
    return await client.callTool({ name: "read.accounting.get_account_turnover", arguments: args });
  } finally {
    await client.close();
  }
}

const text = (r: Awaited<ReturnType<typeof callTool>>): string =>
  (r.content as Array<{ text: string }>)[0]!.text;

describe("read.accounting.get_account_turnover", () => {
  it("возвращает ОСВ: итог, субсчета, период, scan; организация — фильтром в 1С", async () => {
    const { ctx, paths } = fakeCtx([
      row(guid(1), { openingCr: 100.1, turnoverDr: 70.2, turnoverCr: 20.3, closingCr: 50.2 }),
      row(guid(2), { openingDr: 10, turnoverCr: 10 }),
    ]);
    const r = await callTool(ctx, {
      organization: "Ромашка",
      account: "60",
      from: "2025-01-01",
      to: "2025-03-31",
    });
    expect(r.isError).toBeFalsy();
    const s = r.structuredContent as Record<string, unknown>;
    expect(s).toMatchObject({
      database: "main",
      organization: "ООО Ромашка",
      account: "60",
      period: { from: "2025-01-01", to: "2025-03-31" },
      openingDebit: 10,
      openingCredit: 100.1,
      debitTurnover: 70.2,
      creditTurnover: 30.3,
      closingDebit: 0,
      closingCredit: 50.2,
      consistent: true,
    });
    expect((s["accounts"] as Array<{ code: string }>).map((a) => a.code)).toEqual(["60.01", "60.02"]);
    expect((s["scan"] as { rowsScanned: number }).rowsScanned).toBe(2);
    const regCall = decodeURIComponent(paths.find((p) => p.startsWith(REG))!);
    expect(regCall).toContain(
      "/BalanceAndTurnovers(StartPeriod=datetime'2025-01-01T00:00:00',EndPeriod=datetime'2025-04-01T00:00:00')",
    );
    expect(regCall).toContain(`Организация_Key eq guid'${ORG}'`);
  });

  it("неизвестная структура ответа — isError, а не {consistent:true, все суммы 0}", async () => {
    const { ctx } = fakeCtx([{ Account_Key: guid(1), Something: 1 }]);
    const r = await callTool(ctx, { account: "60", from: "2025-01-01", to: "2025-01-31" });
    expect(r.isError).toBe(true);
    expect(r.structuredContent).toBeUndefined();
    expect(text(r)).toMatch(/BalanceAndTurnovers.*несовместимой структуры/);
    expect(text(r)).toContain("Account_Key, Something");
  });

  it("from > to — понятная ошибка без запросов к 1С", async () => {
    const { ctx, paths } = fakeCtx([]);
    const r = await callTool(ctx, { account: "51", from: "2025-02-01", to: "2025-01-01" });
    expect(r.isError).toBe(true);
    expect(text(r)).toMatch(/Период задан наоборот/);
    expect(paths).toHaveLength(0);
  });

  it("несуществующий счёт — понятная ошибка", async () => {
    const { ctx } = fakeCtx([]);
    const r = await callTool(ctx, { account: "99", from: "2025-01-01", to: "2025-01-31" });
    expect(r.isError).toBe(true);
    expect(text(r)).toMatch(/Счёт "99" не найден/);
  });

  it("мусор в коде счёта отсекается валидацией входа", async () => {
    const { ctx, paths } = fakeCtx([]);
    const r = await callTool(ctx, { account: "60') or (1 eq 1", from: "2025-01-01", to: "2025-01-31" });
    expect(r.isError).toBe(true);
    expect(paths).toHaveLength(0);
  });

  it("нет движений и остатков — нули и пояснение", async () => {
    const { ctx } = fakeCtx([]);
    const r = await callTool(ctx, { account: "60.01", from: "2025-01-01", to: "2025-01-31" });
    const s = r.structuredContent as Record<string, unknown>;
    expect(s["closingDebit"]).toBe(0);
    expect(s["accounts"]).toEqual([]);
    expect(s["note"]).toBeTypeOf("string");
  });
});

describe("redBalanceNote — пояснение «красного» сальдо", () => {
  it("отрицательная сумма в Дт/Кт → пояснение с чистым сальдо; без отрицательных — пусто", async () => {
    const { redBalanceNote } = await import("../src/tools/registers.js");
    const t = {
      openingDr: 0,
      openingCr: 0,
      turnoverDr: 0,
      turnoverCr: 0,
      closingDr: -36_007_292,
      closingCr: 9_720_289,
    };
    expect(redBalanceNote(t).note).toContain("-457275.81");
    expect(redBalanceNote({ ...t, closingDr: 100, closingCr: 0 })).toEqual({});
  });
});
