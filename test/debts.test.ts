import { describe, expect, it, vi, afterEach } from "vitest";
import {
  addDays,
  ageCounterparty,
  contractTermDays,
  nonCalendarTermKind,
  daysBetween,
  settlementRows,
  type DocInfo,
} from "../src/odata/settlements-aging.js";
import { cashKind, contractInfo, invoiceStatus } from "../src/tools/debts.js";
import { fake1C } from "./helpers/fake-1c.js";

const CP = "11111111-1111-1111-1111-111111111111";
const CONTRACT = "22222222-2222-2222-2222-222222222222";
const SALE = "33333333-3333-3333-3333-333333333333";
const PAY = "44444444-4444-4444-4444-444444444444";

const balanceRow = (dr: number, cr: number, doc?: { ref: string; type: string }) => ({
  ExtDimension1: CP,
  ExtDimension1_Type: "StandardODATA.Catalog_Контрагенты",
  ExtDimension2: CONTRACT,
  ExtDimension2_Type: "StandardODATA.Catalog_ДоговорыКонтрагентов",
  ExtDimension3: doc?.ref ?? "",
  ExtDimension3_Type: doc ? `StandardODATA.${doc.type}` : "StandardODATA.Undefined",
  СуммаBalanceDr: dr,
  СуммаBalanceCr: cr,
});

describe("settlementRows", () => {
  it("субконто по типу: контрагент, договор, документ расчётов; знак — долг в нашу сторону", () => {
    const rows = settlementRows(
      [balanceRow(1000.1, 0, { ref: SALE, type: "Document_РеализацияТоваровУслуг" })],
      "receivable",
    );
    expect(rows).toEqual([
      {
        counterparty: CP,
        contract: CONTRACT,
        contractSet: "Catalog_ДоговорыКонтрагентов",
        document: SALE,
        documentSet: "Document_РеализацияТоваровУслуг",
        cents: 100010,
      },
    ]);
  });

  it("кредиторка: кредитовое сальдо 60 — наш долг (> 0)", () => {
    const [r] = settlementRows([balanceRow(0, 500)], "payable");
    expect(r!.cents).toBe(50000);
    expect(r!.document).toBeUndefined();
  });

  it("строка без контрагента пропускается", () => {
    expect(settlementRows([{ СуммаBalanceDr: 10 }], "receivable")).toEqual([]);
  });
});

describe("contractTermDays", () => {
  it("срок только при флаге «Установлен срок оплаты»", () => {
    expect(contractTermDays({ УстановленСрокОплаты: true, СрокОплаты: 30 })).toBe(30);
    expect(contractTermDays({ УстановленСрокОплаты: true, СрокОплаты: 0 })).toBe(0);
    expect(contractTermDays({ УстановленСрокОплаты: false, СрокОплаты: 30 })).toBeUndefined();
    expect(contractTermDays(undefined)).toBeUndefined();
  });
});

describe("даты", () => {
  it("дни между датами и сдвиг через границу месяца/года", () => {
    expect(daysBetween("2026-01-31", "2026-03-01")).toBe(29);
    expect(addDays("2025-12-15", 30)).toBe("2026-01-14");
  });
});

describe("ageCounterparty", () => {
  const docs = new Map<string, DocInfo>([
    [SALE, { date: "2026-08-01", number: "0000-000090" }],
    ["sale-2", { date: "2026-10-10" }],
  ]);
  const row = (cents: number, document?: string) => ({
    counterparty: CP,
    contract: CONTRACT,
    cents,
    ...(document ? { document, documentSet: "Document_РеализацияТоваровУслуг" } : {}),
  });

  it("срок из договора: просрочка по дням, в срок, аванс отдельно", () => {
    const a = ageCounterparty(
      [row(100000, SALE), row(50000, "sale-2"), row(-30000, PAY)],
      docs,
      new Map([[CONTRACT, 30]]),
      "2026-10-21",
    );
    // 01.08 + 30 = 31.08 → на 21.10 просрочено 51 день.
    expect(a.overdueCents).toBe(100000);
    expect(a.buckets["31-60"]).toBe(100000);
    expect(a.maxOverdueDays).toBe(51);
    expect(a.notDueCents).toBe(50000);
    expect(a.advanceCents).toBe(30000);
    expect(a.netCents).toBe(120000);
    expect(a.documents[0]).toMatchObject({
      status: "overdue",
      dueDate: "2026-08-31",
      overdueDays: 51,
      termSource: "contract",
      number: "0000-000090",
    });
    expect(a.ageBuckets).toEqual({ "0-30": 50000, "31-60": 0, "61-90": 100000, "90+": 0 });
  });

  it("срок не задан — no_term с возрастом; defaultTermDays включает просрочку", () => {
    const noTerm = ageCounterparty([row(100000, SALE)], docs, new Map(), "2026-10-21");
    expect(noTerm.noTermCents).toBe(100000);
    expect(noTerm.overdueCents).toBe(0);
    expect(noTerm.documents[0]).toMatchObject({ status: "no_term", ageDays: 81 });

    const withDefault = ageCounterparty([row(100000, SALE)], docs, new Map(), "2026-10-21", 10);
    expect(withDefault.documents[0]).toMatchObject({
      status: "overdue",
      termSource: "default",
      overdueDays: 71,
    });
    expect(withDefault.buckets["61-90"]).toBe(100000);
  });

  it("без документа расчётов или его даты возраст не угадывается", () => {
    const a = ageCounterparty(
      [row(100000), row(5000, "unknown")],
      docs,
      new Map([[CONTRACT, 30]]),
      "2026-10-21",
    );
    expect(a.noDocumentCents).toBe(105000);
    expect(a.overdueCents).toBe(0);
    expect(a.documents.every((d) => d.status === "no_document")).toBe(true);
  });
});

describe("invoiceStatus", () => {
  it("по сумме оплат", () => {
    expect(invoiceStatus(1000, 0)).toBe("unpaid");
    expect(invoiceStatus(1000, 400)).toBe("partial");
    expect(invoiceStatus(1000, 1000)).toBe("paid");
    expect(invoiceStatus(1000, 1200)).toBe("overpaid");
  });
});

describe("cashKind", () => {
  it("вид денег по коду счёта", () => {
    expect(cashKind("50.01")).toBe("cash");
    expect(cashKind("51")).toBe("bank");
    expect(cashKind("52")).toBe("currency");
    expect(cashKind("55.03")).toBe("special");
    expect(cashKind("57.03")).toBe("transit");
    expect(cashKind("58")).toBe("other");
  });
});

describe("nonCalendarTermKind: срок не в календарных днях", () => {
  const term = { УстановленСрокОплаты: true, СрокОплаты: 10 };
  it.each([
    ["Рабочие", "Рабочие"],
    ["Банковские", "Банковские"],
    ["Календарные", undefined],
    ["календарных", undefined],
    ["", undefined],
    ["  ", undefined],
  ])("%j → %j", (kind, expected) => {
    expect(nonCalendarTermKind({ ...term, ВидДнейСрокаОплаты: kind })).toBe(expected);
  });
  it("без флага срока и без договора — не считается", () => {
    expect(
      nonCalendarTermKind({ УстановленСрокОплаты: false, ВидДнейСрокаОплаты: "Рабочие" }),
    ).toBeUndefined();
    expect(nonCalendarTermKind({ ...term })).toBeUndefined();
    expect(nonCalendarTermKind(undefined)).toBeUndefined();
  });
});

describe("contractInfo: договоры со сроком не в календарных днях", () => {
  const CONTRACT = "11111111-1111-4111-8111-111111111111";
  const row = { contract: CONTRACT, contractSet: "Catalog_ДоговорыКонтрагентов", cents: 100 } as never;
  const run = async (contract: Record<string, unknown>) => {
    const { conn } = fake1C();
    const spy = vi.spyOn(conn.client, "getCollection").mockResolvedValue({
      value: [{ Ref_Key: CONTRACT, Description: "Договор №1", ...contract }] as never[],
    });
    return { info: await contractInfo(conn, [row]), url: decodeURIComponent(String(spy.mock.calls[0]![0])) };
  };

  it("рабочие дни: договор назван, срок остаётся как есть", async () => {
    const { info, url } = await run({
      УстановленСрокОплаты: true,
      СрокОплаты: 10,
      ВидДнейСрокаОплаты: "Рабочие",
    });
    expect(url).toContain("ВидДнейСрокаОплаты");
    expect(info.nonCalendar).toEqual([{ contract: "Договор №1", dayKind: "Рабочие" }]);
    expect(info.terms.get(CONTRACT)).toBe(10);
  });

  it("пустое поле, как в живых базах: без предупреждения", async () => {
    const { info } = await run({ УстановленСрокОплаты: true, СрокОплаты: 10, ВидДнейСрокаОплаты: "" });
    expect(info.nonCalendar).toEqual([]);
  });
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});
