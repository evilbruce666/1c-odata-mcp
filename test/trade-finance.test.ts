import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { ageBucket } from "../src/odata/partner-settlements.js";
import { localToday } from "../src/odata/settlements-aging.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/**
 * Деньги и расчёты с партнёрами в Управлении торговлей на синтетике по снимку схемы УТ 11.5. Метрика: итоги совпадают с
 * ручным расчётом до копейки; отбор по организации и партнёру исключает чужие строки.
 */

afterEach(() => {
  vi.unstubAllGlobals();
});

const uuid = (kind: string, i: number) => `${kind}0000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const ORG = (i: number) => uuid("e", i);
const PARTNER = (i: number) => uuid("f", i);
const KEY = (i: number) => uuid("a", i);

const BANK = [
  { БанковскийСчет_Key: uuid("b", 1), Организация_Key: ORG(1), СуммаBalance: 1000, СуммаРеглBalance: 1000 },
  // Валютный счёт: 100 в валюте, 9 000 в рублях.
  { БанковскийСчет_Key: uuid("b", 2), Организация_Key: ORG(2), СуммаBalance: 100, СуммаРеглBalance: 9000 },
];
const CASH = [
  { Касса_Key: uuid("c", 1), Организация_Key: ORG(1), СуммаBalance: 500.5, СуммаРеглBalance: 500.5 },
];
const POS = [{ КассаККМ_Key: uuid("d", 1), Организация_Key: ORG(1), СуммаBalance: 70, СуммаРеглBalance: 70 }];
const TRANSIT = [
  {
    ВидПереводаДенежныхСредств: "ИнкассацияВБанк",
    Организация_Key: ORG(1),
    СуммаBalance: 30,
    СуммаРеглBalance: 30,
  },
];

const CLIENT_ROWS = [
  // Партнёр 1, организация 1: долг 100 просрочен 40 дней (на 2026-11-10), 50 — в срок; предоплата 20.
  { k: 1, due: "2026-10-01T00:00:00", debt: 100, pre: 0 },
  { k: 1, due: "2026-12-01T00:00:00", debt: 50, pre: 0 },
  { k: 1, due: "2026-09-01T00:00:00", debt: 0, pre: 20 },
  // Партнёр 2, организация 2: долг 300 просрочен больше 180 дней.
  { k: 2, due: "2026-01-01T00:00:00", debt: 300, pre: 0 },
];
const KEYS = new Map([
  [KEY(1), { Партнер_Key: PARTNER(1), Организация_Key: ORG(1) }],
  [KEY(2), { Партнер_Key: PARTNER(2), Организация_Key: ORG(2) }],
]);

const guidsIn = (text: string) =>
  [...text.matchAll(/guid'([0-9a-f-]{36})'/gi)].map((m) => m[1]!.toLowerCase());
const byOrg = <T extends { Организация_Key: string }>(rows: T[], q: FakeQuery) => {
  const orgs = guidsIn(q.args["Condition"] ?? "");
  return orgs.length ? rows.filter((r) => orgs.includes(r.Организация_Key)) : rows;
};

function data(q: FakeQuery) {
  const filter = q.params.get("$filter") ?? "";
  if (q.fn === "Balance") {
    if (q.base === "AccumulationRegister_ДенежныеСредстваБезналичные") return byOrg(BANK, q);
    if (q.base === "AccumulationRegister_ДенежныеСредстваНаличные") return byOrg(CASH, q);
    if (q.base === "AccumulationRegister_ДенежныеСредстваВКассахККМ") return byOrg(POS, q);
    if (q.base === "AccumulationRegister_ДенежныеСредстваВПути") return byOrg(TRANSIT, q);
    if (q.base === "AccumulationRegister_РасчетыСКлиентамиПоСрокам")
      return CLIENT_ROWS.map((r) => ({
        АналитикаУчетаПоПартнерам_Key: KEY(r.k),
        АналитикаУчетаПоПартнерам: KEYS.get(KEY(r.k)),
        ДатаПлановогоПогашения: r.due,
        ДолгРеглBalance: r.debt,
        ПредоплатаРеглBalance: r.pre,
      }));
    if (q.base === "AccumulationRegister_РасчетыСПоставщикамиПоСрокам") return [];
  }
  if (q.base === "Catalog_Организации") {
    return [
      { Ref_Key: ORG(1), Description: "Организация 1", ИНН: "7700000011" },
      { Ref_Key: ORG(2), Description: "Организация 2", ИНН: "7700000022" },
    ].filter((o) => !guidsIn(filter).length || guidsIn(filter).includes(o.Ref_Key));
  }
  if (q.base === "Catalog_Контрагенты") {
    // Контрагент 9 принадлежит партнёру 1.
    return [{ Ref_Key: uuid("c", 9), Партнер_Key: PARTNER(1) }].filter((c) =>
      guidsIn(filter).includes(c.Ref_Key),
    );
  }
  if (q.base === "Catalog_Партнеры") {
    return [
      { Ref_Key: PARTNER(1), Description: "Партнёр 1" },
      { Ref_Key: PARTNER(2), Description: "Партнёр 2" },
    ].filter((p) => guidsIn(filter).includes(p.Ref_Key));
  }
  return undefined;
}

async function setup() {
  const fake = fake1C({ READ_ONLY: "true" }, { schema: "ut11", data });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (name: string, args: object = {}) => {
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake };
}

describe("get_cash_balance в УТ — регистры денежных средств", () => {
  it("итоги и строки по счетам, кассам и кассам ККМ; валютный счёт — в рублях и в валюте", async () => {
    const { call, fake } = await setup();
    const r = await call("read.analytics.get_cash_balance");
    expect(r.isError, r.text).toBeFalsy();
    const s = r.structuredContent;
    // 1000 + 9000 + 500,50 + 70 = 10 570,50; в пути 30.
    expect([s["available"], s["inTransit"], s["special"], s["source"]]).toEqual([
      10570.5,
      30,
      0,
      "registers",
    ]);
    const balances = s["balances"] as { kind: string; amount: number; currencyAmount?: number }[];
    expect(balances.find((x) => x.amount === 9000)).toMatchObject({ kind: "bank", currencyAmount: 100 });
    expect(balances.filter((x) => x.currencyAmount !== undefined)).toHaveLength(1);
    expect(fake.violations).toEqual([]);
  });

  it("организация — отбор в 1С; дата — Period следующего дня", async () => {
    const { call, fake } = await setup();
    const s = (await call("read.analytics.get_cash_balance", { organization: ORG(2) })).structuredContent;
    expect(s["available"]).toBe(9000);
    expect(s["inTransit"]).toBe(0);
    await call("read.analytics.get_cash_balance", { asOf: "2026-10-31" });
    expect(fake.violations).toEqual([]);
  });

  it("деньги в пути в разных валютах одного вида не складываются; вид перевода — не Ref_Key (ревью)", async () => {
    TRANSIT.push({
      ВидПереводаДенежныхСредств: "ИнкассацияВБанк",
      Организация_Key: ORG(1),
      СуммаBalance: 100,
      СуммаРеглBalance: 9000,
      Валюта_Key: uuid("9", 1),
    } as never);
    try {
      const { call } = await setup();
      const s = (await call("read.analytics.get_cash_balance")).structuredContent;
      const transit = (
        s["balances"] as { kind: string; amount: number; currencyAmount?: number; holderRef?: string }[]
      ).filter((b) => b.kind === "transit");
      expect(transit.map((t) => [t.amount, t.currencyAmount])).toEqual([
        [9000, 100],
        [30, undefined],
      ]);
      expect(transit.every((t) => t.holderRef === undefined)).toBe(true);
    } finally {
      TRANSIT.pop();
    }
  });

  it("отрицательные деньги в пути — примечание о незакрытых переводах", async () => {
    TRANSIT.push({
      ВидПереводаДенежныхСредств: "ПеремещениеВДругуюКассу",
      Организация_Key: ORG(1),
      СуммаBalance: -500,
      СуммаРеглBalance: -500,
    });
    try {
      const { call } = await setup();
      const s = (await call("read.analytics.get_cash_balance")).structuredContent;
      expect(s["inTransit"]).toBe(-470);
      expect(s["note"]).toContain("отправка в регистре не отражена");
    } finally {
      TRANSIT.pop();
    }
  });
});

describe("get_partner_settlements — расчёты по срокам", () => {
  it("долг, предоплата, просрочка и сроки на дату; net по партнёрам", async () => {
    const { call, fake } = await setup();
    const r = await call("read.analytics.get_partner_settlements", { side: "clients", asOf: "2026-11-10" });
    expect(r.isError, r.text).toBeFalsy();
    const s = r.structuredContent;
    // Долг 450, предоплата 20, net 430; просрочено 100 (40 дней) + 300 (больше 180), в срок 50.
    expect([s["debt"], s["prepayment"], s["net"], s["overdue"]]).toEqual([450, 20, 430, 400]);
    expect(s["aging"]).toEqual({
      notDue: 50,
      days1to30: 0,
      days31to90: 100,
      days91to180: 0,
      over180: 300,
      noDueDate: 0,
    });
    expect(s["partners"]).toEqual([
      {
        partner: "Партнёр 2",
        partnerRef: PARTNER(2),
        debt: 300,
        prepayment: 0,
        net: 300,
        overdue: 300,
        maxOverdueDays: 313,
      },
      {
        partner: "Партнёр 1",
        partnerRef: PARTNER(1),
        debt: 150,
        prepayment: 20,
        net: 130,
        overdue: 100,
        maxOverdueDays: 40,
      },
    ]);
    expect(s["note"]).toContain("незачтённые авансы");
    expect(fake.violations).toEqual([]);
  });

  it("отбор по организации и партнёру исключает чужие строки; overdueOnly", async () => {
    const { call } = await setup();
    const org = (
      await call("read.analytics.get_partner_settlements", {
        side: "clients",
        organization: ORG(1),
        asOf: "2026-11-10",
      })
    ).structuredContent;
    expect([org["debt"], org["prepayment"]]).toEqual([150, 20]);
    const partner = (
      await call("read.analytics.get_partner_settlements", {
        side: "clients",
        partner: `{${PARTNER(2)}}`,
        asOf: "2026-11-10",
      })
    ).structuredContent;
    expect(partner["debt"]).toBe(300);
    const overdue = (
      await call("read.analytics.get_partner_settlements", {
        side: "clients",
        overdueOnly: true,
        asOf: "2026-11-01",
      })
    ).structuredContent;
    // На 01.11 у партнёра 1 просрочено 100 (31 день), у партнёра 2 — 300.
    expect((overdue["partners"] as unknown[]).length).toBe(2);
    const early = (
      await call("read.analytics.get_partner_settlements", {
        side: "clients",
        overdueOnly: true,
        asOf: "2025-12-31",
      })
    ).structuredContent;
    expect(early["partners"]).toEqual([]);
  });

  it("в Бухгалтерии — понятный отказ с именем регистра", async () => {
    const fake = fake1C({ READ_ONLY: "true" });
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({
      name: "read.analytics.get_partner_settlements",
      arguments: { side: "suppliers" },
    })) as {
      isError?: boolean;
      content: { text: string }[];
    };
    expect(r.isError).toBe(true);
    expect(r.content[0]!.text).toContain("Расчёты с поставщиками по срокам");
  });

  it("Ref_Key контрагента — расчёты его партнёра с примечанием; неизвестный Ref_Key — отказ (ревью: тихий ноль)", async () => {
    const { call } = await setup();
    const s = (
      await call("read.analytics.get_partner_settlements", {
        side: "clients",
        partner: uuid("c", 9),
        asOf: "2026-11-10",
      })
    ).structuredContent;
    expect([s["debt"], s["prepayment"]]).toEqual([150, 20]);
    expect(s["note"]).toContain("Ref_Key контрагента");
    const unknown = await call("read.analytics.get_partner_settlements", {
      side: "clients",
      partner: uuid("c", 77),
    });
    expect(unknown.isError).toBe(true);
    expect(unknown.text).toContain("не найден");
  });

  it("отрицательная строка долга входит в сроки и просрочку: просрочка не больше долга (ревью)", async () => {
    CLIENT_ROWS.push({ k: 2, due: "2026-01-01T00:00:00", debt: -30, pre: 0 });
    try {
      const { call } = await setup();
      const s = (
        await call("read.analytics.get_partner_settlements", {
          side: "clients",
          partner: PARTNER(2),
          asOf: "2026-11-10",
        })
      ).structuredContent;
      expect([s["debt"], s["overdue"]]).toEqual([270, 270]);
      const aging = s["aging"] as Record<string, number>;
      expect(Object.values(aging).reduce((t, x) => t + x, 0)).toBe(270);
    } finally {
      CLIENT_ROWS.pop();
    }
  });

  it("без asOf — местная дата, а не UTC (ревью: на востоке утром «сегодня» было вчера)", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    // 23:30 по UTC — в часовом поясе восточнее UTC уже следующий день.
    vi.setSystemTime(new Date("2026-11-12T23:30:00Z"));
    try {
      const { call } = await setup();
      const s = (await call("read.analytics.get_partner_settlements", { side: "clients" })).structuredContent;
      expect(s["on"]).toBe(localToday());
    } finally {
      vi.useRealTimers();
    }
  });

  it("корзины сроков: граница — день в день в срок", () => {
    expect(ageBucket("2026-11-10", "2026-11-10")).toBe("notDue");
    expect(ageBucket("2026-11-09", "2026-11-10")).toBe("days1to30");
    expect(ageBucket("2026-10-11", "2026-11-10")).toBe("days1to30");
    expect(ageBucket("2026-10-10", "2026-11-10")).toBe("days31to90");
    expect(ageBucket(undefined, "2026-11-10")).toBe("noDueDate");
  });
});
