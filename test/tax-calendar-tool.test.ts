import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { netAccruals } from "../src/odata/tax-calendar.js";
import { callTool } from "./helpers/call-tool.js";

const TOOL = "read.accounting.get_tax_calendar";
const ORG = "11111111-1111-4111-8111-111111111111";
const ref = (i: number) => `00000000-0000-4000-8000-${String(i).padStart(12, "0")}`;
const USN = ref(1);
const PSN = ref(2);
const IPFIX = ref(3);
const PATENT = ref(9);
const EMPTY = "00000000-0000-0000-0000-000000000000";
const SET = {
  org: "Catalog_Организации",
  chart: "ChartOfAccounts_Хозрасчетный",
  accruals: "InformationRegister_РасчетНачисленияНалоговНаЕНС_RecordType",
  settings: "InformationRegister_НастройкиСистемыНалогообложения",
  journal: "Document_РегламентированныйОтчет",
  notices: "Document_УведомлениеОбИсчисленныхСуммахНалогов",
  calc: "InformationRegister_РасчетНалогаУСН_RecordType",
} as const;
const ACCOUNTS = [
  { Ref_Key: USN, Code: "68.12", Description: "Налог при УСН" },
  { Ref_Key: PSN, Code: "68.14", Description: "Налог при ПСН" },
  { Ref_Key: IPFIX, Code: "69.06.6", Description: "Взносы ИП" },
];
const KBK = "18210501011011000110";

type Rows = Record<string, Array<Record<string, unknown>> | Error>;
const accrual = (account: string, periodEnd: string, due: string, amount: number, patent = EMPTY) => ({
  СчетУчета_Key: account,
  ПериодРасчета: `${periodEnd}T00:00:00`,
  СрокУплаты: `${due}T00:00:00`,
  ДокументПатент_Key: patent,
  Начислено: amount,
  Active: true,
});
const setting = (extra: Record<string, unknown> = {}) => ({
  Period: "2020-01-01T00:00:00",
  СистемаНалогообложения: "Упрощенная",
  ОбъектНалогообложенияУСН: "Доходы",
  ПрименяетсяУСНПатент: false,
  ...extra,
});
const notice = (day: string, periodEnd: string, amount: number) => ({
  Ref_Key: ref(Number(day.replaceAll("-", "")) + amount),
  Date: `${day}T00:00:00`,
  Налоги: [{ КодБК: KBK, КонецПериода: `${periodEnd}T00:00:00`, Сумма: amount }],
});
const report = (year: number, extra: Record<string, unknown> = {}) => ({
  Ref_Key: ref(year),
  ИсточникОтчета: "РегламентированныйОтчетУСН",
  ДатаНачала: `${year}-01-01T00:00:00`,
  ДатаОкончания: `${year}-12-31T00:00:00`,
  Вид: 0,
  ДатаПодписи: `${year + 1}-04-01T00:00:00`,
  ДатаОтправки: "0001-01-01T00:00:00",
  ...extra,
});

interface Item {
  kind: string;
  title: string;
  due: string;
  status: string;
  dueSource: string;
  amount?: number;
  lawDue?: string;
  notices?: unknown[];
  report?: Record<string, unknown>;
}
interface Calendar {
  taxpayer: string;
  regime: { regimes: string[]; source: string };
  items: Item[];
  warnings: string[];
  notes: string[];
  missing: string[];
  rulesCoverage: { years: number[]; missing: number[] };
  notCovered: string[];
}

async function calendar(
  rows: Rows = {},
  args: Record<string, unknown> = {},
  options: { type?: string; registered?: string } = {},
) {
  const base: Rows = {
    [SET.org]: [
      {
        Ref_Key: ORG,
        Description: "Организация",
        ЮридическоеФизическоеЛицо: options.type ?? "ФизическоеЛицо",
        ДатаРегистрации: `${options.registered ?? "2015-01-01"}T00:00:00`,
      },
    ],
    [SET.chart]: ACCOUNTS,
    [SET.accruals]: [],
    [SET.settings]: [setting()],
    [SET.journal]: [],
    [SET.notices]: [],
    [SET.calc]: [],
    ...rows,
  };
  const { result, violations } = await callTool(
    TOOL,
    { from: "2027-01-01", to: "2027-04-30", ...args },
    (q) => {
      const own = base[q.base];
      if (own instanceof Error) throw own;
      return own;
    },
  );
  expect(violations).toEqual([]);
  expect(result.isError, JSON.stringify(result.content)).toBeFalsy();
  return result.structuredContent as unknown as Calendar;
}
const find = (c: Calendar, title: RegExp) => c.items.filter((i) => title.test(i.title));

beforeEach(() => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date("2027-01-15T12:00:00"));
});
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("сведение записей регистра начислений", () => {
  it("сторно на соседней дате гасит начисление того же патента: платить нечего", () => {
    const net = netAccruals([
      accrual(PSN, "2023-01-31", "2024-04-01", 100, PATENT),
      accrual(PSN, "2023-03-31", "2024-04-01", -100, PATENT),
      accrual(PSN, "2023-01-31", "2025-01-09", 200, PATENT),
      accrual(PSN, "2023-03-31", "2024-12-31", -200, PATENT),
    ]);
    expect(net).toEqual({ due: [], excess: [] });
  });

  it("чужой патент и далёкая дата сторно не гасят: остаток показан отдельно", () => {
    const net = netAccruals([
      accrual(PSN, "2026-01-31", "2026-04-01", 100, PATENT),
      accrual(PSN, "2026-03-31", "2026-04-01", -40, ref(10)),
      accrual(USN, "2026-03-31", "2026-04-28", 500),
      accrual(USN, "2026-06-30", "2026-07-28", -70),
    ]);
    expect(net.due.map((a) => a.amount).sort()).toEqual([100, 500]);
    expect(net.excess.map((a) => a.amount).sort()).toEqual([-40, -70].sort());
  });

  it("строки одного налога и срока складываются, неактивные не считаются", () => {
    const net = netAccruals([
      accrual(USN, "2026-03-31", "2026-04-28", 10.1),
      accrual(USN, "2026-03-31", "2026-04-28", 20.2),
      { ...accrual(USN, "2026-03-31", "2026-04-28", 999), Active: false },
    ]);
    expect(net.due).toMatchObject([{ amount: 30.3, periodEnd: "2026-03-31" }]);
  });
});

describe("платежи", () => {
  it("начисление 1С заменяет строку справочника: одна обязанность — одна строка, сумма и срок из 1С", async () => {
    const c = await calendar({ [SET.accruals]: [accrual(USN, "2026-12-31", "2027-04-28", 1234)] });
    expect(find(c, /^Налог по УСН за 2026/)).toMatchObject([
      { status: "accrued", dueSource: "1c", amount: 1234, due: "2027-04-28" },
    ]);
    expect(c.warnings).toEqual([]);
  });

  it("срок в 1С не как в справочнике — обе даты и предупреждение", async () => {
    const c = await calendar({ [SET.accruals]: [accrual(USN, "2026-12-31", "2027-04-27", 5)] });
    expect(find(c, /^Налог по УСН за 2026/)).toMatchObject([{ due: "2027-04-27", lawDue: "2027-04-28" }]);
    expect(c.warnings.join(" ")).toContain("2027-04-28");
  });

  it("начисление без правила (патент) — своя строка с названием счёта", async () => {
    const c = await calendar({ [SET.accruals]: [accrual(PSN, "2027-01-31", "2027-04-01", 77, PATENT)] });
    expect(find(c, /Налог при ПСН/)).toMatchObject([{ status: "accrued", amount: 77 }]);
  });

  it("без начисления — срок из статьи и без суммы", async () => {
    const c = await calendar();
    for (const item of c.items) expect(item).not.toHaveProperty("amount");
    expect(find(c, /^Налог по УСН за 2026/)).toMatchObject([{ status: "not_accrued", dueSource: "law" }]);
  });

  it("отрицательный остаток — предупреждение, а не платёж", async () => {
    const c = await calendar({ [SET.accruals]: [accrual(USN, "2027-03-31", "2027-04-28", -50)] });
    expect(c.items.filter((i) => i.amount !== undefined)).toEqual([]);
    expect(c.warnings.join(" ")).toContain("отрицательный остаток");
  });
});

describe("режим решает, какие сроки из справочника показать", () => {
  const titles = async (args: Record<string, unknown>, rows: Rows = {}) =>
    (await calendar(rows, args)).items.map((i) => i.title);

  it("параметр regime меняет набор и помечается как слова пользователя", async () => {
    const usn = await titles({ regime: ["usn_income"] });
    const psn = await titles({ regime: ["psn"] });
    expect(usn.some((t) => t.startsWith("Декларация по УСН"))).toBe(true);
    expect(psn.some((t) => t.includes("УСН"))).toBe(false);
    expect((await calendar({}, { regime: ["psn"] })).regime).toMatchObject({ source: "user" });
  });

  it("режим берётся за период обязанности: после ухода с УСН декларация за прошлый год остаётся", async () => {
    const rows = {
      [SET.settings]: [
        setting({ Period: "2027-01-01T00:00:00", СистемаНалогообложения: "Общая" }),
        setting(),
      ],
    };
    const got = await titles({}, rows);
    expect(got).toContain("Декларация по УСН за 2026 год");
    expect(got.some((t) => /2027/.test(t) && t.startsWith("Аванс по УСН"))).toBe(false);
  });

  it("взносы ИП — при любом режиме, даже неизвестном; на НПД и АУСН их нет", async () => {
    const contributions = async (rows: Rows) =>
      (await calendar(rows, { from: "2027-06-01", to: "2027-07-31" })).items.filter((i) => /1%/.test(i.title))
        .length;
    await expect(
      contributions({ [SET.settings]: [setting({ СистемаНалогообложения: "Общая" })] }),
    ).resolves.toBe(1);
    await expect(contributions({ [SET.settings]: [] })).resolves.toBe(1);
    for (const extra of [
      { СистемаНалогообложения: "НалогНаПрофессиональныйДоход" },
      { ПрименяетсяАУСН: true },
    ])
      await expect(contributions({ [SET.settings]: [setting(extra)] })).resolves.toBe(0);
  });

  it("режим не определён — об этом сказано и предложен параметр", async () => {
    const c = await calendar({ [SET.settings]: [] });
    expect(c.regime).toMatchObject({ source: "unknown", regimes: [] });
    expect(c.notes.join(" ")).toContain("параметром regime");
  });

  it("у организации свои сроки и нет взносов ИП", async () => {
    const c = await calendar({}, {}, { type: "ЮридическоеЛицо" });
    expect(c.taxpayer).toBe("organization");
    expect(find(c, /^Декларация/)).toMatchObject([{ due: "2027-03-25" }]);
    expect(c.items.filter((i) => /взнос/i.test(i.title))).toEqual([]);
  });

  it("обязанности периодов до регистрации не показываются", async () => {
    const c = await calendar({}, {}, { registered: "2027-01-10" });
    expect(find(c, /за 2026 год/)).toEqual([]);
  });
});

describe("отчёты", () => {
  const status = async (journal: Rows[string]) =>
    find(await calendar({ [SET.journal]: journal }), /^Декларация/)[0];

  it("есть документ — prepared, с датой отправки — sent, нет — not_found", async () => {
    await expect(status([report(2026)])).resolves.toMatchObject({
      status: "prepared",
      report: { signed: "2027-04-01" },
    });
    await expect(status([report(2026, { ДатаОтправки: "2027-04-02T00:00:00" })])).resolves.toMatchObject({
      status: "sent",
      report: { sent: "2027-04-02" },
    });
    await expect(status([report(2025)])).resolves.toMatchObject({ status: "not_found" });
  });

  it("журнал не прочитался — unknown и объяснение, остальной календарь на месте", async () => {
    const c = await calendar({ [SET.journal]: new Error("нет прав") });
    expect(find(c, /^Декларация/)).toMatchObject([{ status: "unknown" }]);
    expect(c.missing.join(" ")).toContain("Статус отчётов неизвестен");
    expect(find(c, /^Налог по УСН/)).toHaveLength(1);
  });
});

describe("уведомления", () => {
  const Q1 = { from: "2027-04-01", to: "2027-04-30" };
  const q1 = async (rows: Rows) => {
    const c = await calendar(rows, Q1);
    return { item: c.items.find((i) => i.kind === "notice")!, warnings: c.warnings };
  };
  const accrued = [accrual(USN, "2027-03-31", "2027-04-28", 300)];
  const calc = (toPay: number) => [{ ПериодРасчета: "2027-03-31T00:00:00", НалогКУплате: toPay }];

  it("квартал не рассчитан — unknown, а не «не найдено»", async () => {
    await expect(q1({})).resolves.toMatchObject({ item: { status: "unknown" } });
  });

  it("рассчитан, платить нечего — not_required; есть что платить — not_found", async () => {
    await expect(q1({ [SET.calc]: calc(0) })).resolves.toMatchObject({ item: { status: "not_required" } });
    await expect(q1({ [SET.calc]: calc(300) })).resolves.toMatchObject({ item: { status: "not_found" } });
    await expect(q1({ [SET.accruals]: accrued })).resolves.toMatchObject({ item: { status: "not_found" } });
  });

  it("расходится с начислением — обе суммы в предупреждении, уведомления не складываются молча", async () => {
    const got = await q1({
      [SET.accruals]: accrued,
      [SET.notices]: [notice("2027-04-20", "2027-03-31", 100), notice("2027-04-22", "2027-03-31", 50)],
    });
    expect(got.item).toMatchObject({ status: "prepared" });
    expect(got.item.notices).toHaveLength(2);
    expect(got.warnings.join(" ")).toMatch(/300 ₽.*100 ₽.*50 ₽/);
  });

  it("совпадает одно из уведомлений или их сумма — предупреждения нет", async () => {
    for (const amounts of [[300], [100, 300], [100, 200]]) {
      const notices = amounts.map((a, i) => notice(`2027-04-2${i}`, "2027-03-31", a));
      await expect(q1({ [SET.accruals]: accrued, [SET.notices]: notices })).resolves.toMatchObject({
        warnings: [],
      });
    }
  });
});

describe("границы", () => {
  it("год вне справочника — сказано об этом, и ни одной чужой даты", async () => {
    const c = await calendar({}, { from: "2028-01-01", to: "2028-04-30" });
    expect(c.rulesCoverage).toMatchObject({ years: [], missing: [2028] });
    expect(c.items).toEqual([]);
    expect(c.warnings.join(" ")).toContain("2028");
    expect(c.notCovered.length).toBeGreaterThan(0);
  });

  it("неверное окно — отказ", async () => {
    for (const args of [
      { from: "2027-05-01", to: "2027-04-01" },
      { from: "2025-01-01", to: "2027-12-31" },
    ]) {
      const { result } = await callTool(TOOL, args);
      expect(result.isError).toBe(true);
    }
    await expect(callTool(TOOL, { from: "2027-02-30" })).rejects.toThrow();
  });

  it("в Управлении торговлей — отказ без совета публиковать регистры бухучёта", async () => {
    const { result } = await callTool(TOOL, {}, undefined, "ut11");
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.content)).toContain("налоговому учёту Бухгалтерии");
  });
});
