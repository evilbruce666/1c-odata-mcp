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
  usnBook: "AccumulationRegister_КнигаУчетаДоходовИРасходов",
  patentBook: "AccumulationRegister_КнигаУчетаДоходовПатент",
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
  reports?: Array<Record<string, unknown>>;
}
interface Calendar {
  taxpayer: string;
  regime: { regimes: string[]; source: string };
  items: Item[];
  warnings: string[];
  notes: string[];
  missing: string[];
  vatThreshold?: Array<{
    year: number;
    through: string;
    limit: number;
    basis: string;
    incomeUsn: number;
    incomePatent?: number;
    income: number;
    exceeded?: boolean;
  }>;
  rulesCoverage: { years: number[]; missing: number[] };
  notCovered: string[];
}

/** Сколько запросов к 1С сделал последний вызов, и обороты книг доходов, которые он спрашивал. */
let requests = 0;
let turnovers: Array<{ base: string; fn?: string; args: Record<string, string> }> = [];
/** Строки книги дохода по году начала оборотов. */
type Books = Partial<Record<string, Record<number, number>>>;

async function calendar(
  rows: Rows = {},
  args: Record<string, unknown> = {},
  options: { type?: string; registered?: string; books?: Books } = {},
) {
  turnovers = [];
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
    [SET.usnBook]: [],
    [SET.patentBook]: [],
    ...rows,
  };
  const { result, violations, fake } = await callTool(
    TOOL,
    { from: "2027-01-01", to: "2027-04-30", ...args },
    (q) => {
      const own = base[q.base];
      if (own instanceof Error) throw own;
      if (q.base === SET.usnBook || q.base === SET.patentBook) {
        turnovers.push({ base: q.base, ...(q.fn ? { fn: q.fn } : {}), args: q.args });
        const year = Number(/\d{4}/.exec(q.args["StartPeriod"] ?? "")?.[0]);
        const income = options.books?.[q.base]?.[year];
        if (income === undefined) return own;
        // Графа 4 книги УСН — все поступления: в доход она не идёт.
        return q.base === SET.usnBook
          ? [{ Графа4Turnover: income * 3, Графа5Turnover: income }]
          : [{ Графа4Turnover: income }];
      }
      return own;
    },
  );
  requests = fake.requests.length;
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

  it("годовой налог и аванс за 1 квартал на одном счёте и сроке — две обязанности", () => {
    const net = netAccruals([
      accrual(USN, "2026-12-31", "2027-04-28", 900),
      accrual(USN, "2027-03-31", "2027-04-28", 100),
    ]);
    expect(net.due.map((a) => [a.periodEnd, a.amount]).sort()).toEqual([
      ["2026-12-31", 900],
      ["2027-03-31", 100],
    ]);
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

  it("ИП: налог за год и аванс за 1 квартал 28 апреля — две строки, уведомление сверяется со своей суммой", async () => {
    const c = await calendar({
      [SET.accruals]: [
        accrual(USN, "2026-12-31", "2027-04-28", 900),
        accrual(USN, "2027-03-31", "2027-04-28", 100),
      ],
      [SET.notices]: [notice("2027-04-20", "2027-03-31", 100)],
    });
    expect(find(c, /^Налог по УСН за 2026/)).toMatchObject([{ status: "accrued", amount: 900 }]);
    expect(find(c, /^Аванс по УСН за 1 квартал 2027/)).toMatchObject([{ status: "accrued", amount: 100 }]);
    expect(c.warnings).toEqual([]);
  });

  it("без плана счетов начисление не принимает название уведомления или отчёта", async () => {
    const c = await calendar({
      [SET.chart]: [],
      [SET.accruals]: [accrual(USN, "2027-03-31", "2027-04-28", 100)],
    });
    expect(c.items.filter((i) => i.amount !== undefined)).toMatchObject([
      { kind: "payment", title: "Начислено в 1С: налог без названия счёта" },
    ]);
    expect(c.missing.join(" ")).toContain("Названия налогов");
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

  it("ОСНО известна из настроек: не «не определён», а правил нет", async () => {
    const c = await calendar(
      { [SET.settings]: [setting({ СистемаНалогообложения: "Общая" })] },
      {},
      { type: "ЮридическоеЛицо" },
    );
    expect(c.regime).toMatchObject({ source: "settings", system: "Общая", regimes: [] });
    expect(c.items).toEqual([]);
    expect(c.notes.join(" ")).toContain("«Общая»");
    expect(c.notes.join(" ")).not.toContain("не определён");
    expect(c.notes.join(" ")).not.toContain("taxRegime");
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
      reports: [{ signed: "2027-04-01", corrective: false }],
    });
    await expect(status([report(2026, { ДатаОтправки: "2027-04-02T00:00:00" })])).resolves.toMatchObject({
      status: "sent",
      reports: [{ sent: "2027-04-02" }],
    });
    await expect(status([report(2025)])).resolves.toMatchObject({ status: "not_found" });
  });

  it("первичный и корректирующий — оба, первичный первым", async () => {
    await expect(
      status([report(2026, { Вид: 1, ДатаПодписи: "2027-09-01T00:00:00" }), report(2026)]),
    ).resolves.toMatchObject({
      reports: [
        { signed: "2027-04-01", corrective: false },
        { signed: "2027-09-01", corrective: true },
      ],
    });
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

describe("запросы к 1С", () => {
  // Основной тест норм проходит короткую ветку: у синтетических настроек нет УСН. Здесь — все источники.
  it("полная ветка ИП — не больше тринадцати (из них четыре — книги доходов за два года), режим одним запросом на все периоды", async () => {
    await calendar({ [SET.accruals]: [accrual(USN, "2026-12-31", "2027-04-28", 1)] }, { from: "2026-01-01" });
    expect(requests).toBeLessThanOrEqual(13);
  });

  it("без регистра настроек — расчёт и патенты читаются по разу, а не на каждую дату", async () => {
    const { fake } = await callTool(TOOL, { from: "2026-01-01", to: "2027-04-30" }, (q) =>
      q.base === SET.settings ? [] : undefined,
    );
    expect(fake.requests.filter((r) => r.resource === SET.settings)).toHaveLength(1);
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

  it("в Управлении торговлей с несколькими организациями отказ сразу, без просьбы выбрать", async () => {
    const orgs = [1, 2].map((i) => ({ Ref_Key: ref(i), Description: `Организация ${i}` }));
    const { result, queries } = await callTool(
      TOOL,
      {},
      (q) => (q.base === SET.org ? orgs : undefined),
      "ut11",
    );
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.content)).toContain("налоговому учёту Бухгалтерии");
    expect(JSON.stringify(result.content)).not.toContain("несколько организаций");
    expect(queries).toEqual([]);
  });

  it("в Управлении торговлей — отказ без совета публиковать регистры бухучёта", async () => {
    const { result } = await callTool(TOOL, {}, undefined, "ut11");
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.content)).toContain("налоговому учёту Бухгалтерии");
  });
});

describe("порог НДС на УСН", () => {
  const notPayer = setting({ ПлательщикНДСПрименяющийУСН: false });
  const vatWarnings = (c: Calendar) => c.warnings.filter((w) => w.startsWith("Порог НДС"));

  it("ИП: доход УСН и патента складывается, превышение за прошлый год — предупреждение", async () => {
    const c = await calendar(
      { [SET.settings]: [setting({ ПлательщикНДСПрименяющийУСН: false, ПрименяетсяУСНПатент: true })] },
      {},
      { books: { [SET.usnBook]: { 2026: 15_000_000, 2027: 100 }, [SET.patentBook]: { 2026: 6_000_000 } } },
    );
    expect(c.vatThreshold).toMatchObject([
      {
        year: 2026,
        through: "2026-12-31",
        incomeUsn: 15_000_000,
        incomePatent: 6_000_000,
        income: 21_000_000,
        limit: 20_000_000,
        exceeded: true,
        basis: "абз. 5 п. 1 ст. 145 НК РФ",
      },
      { year: 2027, through: "2027-01-15", income: 100, exceeded: false, basis: "абз. 3 п. 5 ст. 145 НК РФ" },
    ]);
    expect(vatWarnings(c)).toHaveLength(1);
    expect(vatWarnings(c)[0]).toContain("освобождения от НДС в 2027 году нет");
    expect(vatWarnings(c)[0]).toContain("не плательщик НДС — проверьте");
  });

  it("обороты — по своей организации, с начала года по сегодня, через функцию Turnovers", async () => {
    await calendar({ [SET.settings]: [notPayer] }, {}, { type: "ЮридическоеЛицо" });
    expect(turnovers).toMatchObject([
      { base: SET.usnBook, fn: "Turnovers" },
      { base: SET.usnBook, fn: "Turnovers" },
    ]);
    expect(turnovers[1]!.args).toMatchObject({
      StartPeriod: "datetime'2027-01-01T00:00:00'",
      EndPeriod: "datetime'2027-01-15T23:59:59'",
      Condition: `Организация_Key eq guid'${ORG}'`,
    });
  });

  it("превышение в этом году называет начало обязанности и не определяет месяц", async () => {
    const c = await calendar(
      { [SET.settings]: [notPayer] },
      {},
      { type: "ЮридическоеЛицо", books: { [SET.usnBook]: { 2027: 20_000_000.01 } } },
    );
    expect(vatWarnings(c)).toHaveLength(1);
    expect(vatWarnings(c)[0]).toContain("с 1-го числа месяца после месяца превышения");
  });

  it("ровно порог — не превышение; приближение предупреждением не бывает", async () => {
    const c = await calendar(
      { [SET.settings]: [notPayer] },
      {},
      { type: "ЮридическоеЛицо", books: { [SET.usnBook]: { 2026: 20_000_000, 2027: 19_999_999 } } },
    );
    expect(c.vatThreshold!.map((f) => f.exceeded)).toEqual([false, false]);
    expect(vatWarnings(c)).toEqual([]);
  });

  it("книга патента не читается — «не превышен» не говорится, причина названа один раз", async () => {
    const c = await calendar(
      { [SET.settings]: [notPayer], [SET.patentBook]: new Error("нет прав") },
      {},
      { books: { [SET.usnBook]: { 2026: 21_000_000, 2027: 5 } } },
    );
    expect(c.vatThreshold!.map((f) => f.exceeded)).toEqual([true, undefined]);
    expect(c.vatThreshold![1]).not.toHaveProperty("incomePatent");
    expect(c.missing.filter((m) => m.startsWith("Доход по патенту"))).toHaveLength(1);
  });

  it("в книгах нет записей — «не проверен», а не ноль", async () => {
    const c = await calendar({ [SET.settings]: [notPayer] });
    expect(c.vatThreshold).toBeUndefined();
    expect(c.notes.filter((n) => /Порог НДС за 202[67] год не проверен/.test(n))).toHaveLength(2);
  });

  it("год до регистрации не проверяется", async () => {
    const c = await calendar(
      { [SET.settings]: [notPayer] },
      {},
      { registered: "2027-01-10", books: { [SET.usnBook]: { 2027: 1 } } },
    );
    expect(c.vatThreshold!.map((f) => f.year)).toEqual([2027]);
    expect(c.notes.join(" ")).not.toContain("2026 год не проверен");
  });

  it("режим назван параметром — «плательщик НДС» всё равно берётся из настроек 1С", async () => {
    const named = { regime: ["usn_income"] };
    const books = { [SET.usnBook]: { 2026: 21_000_000 } };
    const payer = await calendar(
      { [SET.settings]: [setting({ ПлательщикНДСПрименяющийУСН: true })] },
      named,
      { books },
    );
    expect(payer.vatThreshold).toBeUndefined();
    expect(vatWarnings(payer)).toEqual([]);
    expect(turnovers).toEqual([]);
    const other = await calendar({ [SET.settings]: [notPayer] }, named, { books });
    expect(vatWarnings(other)[0]).toContain("не плательщик НДС — проверьте");
  });

  it("настройки при названном режиме не читаются — порог проверяется, причина названа", async () => {
    const c = await calendar(
      { [SET.settings]: new Error("нет прав") },
      { regime: ["usn_income"] },
      { books: { [SET.usnBook]: { 2026: 1 } } },
    );
    expect(c.vatThreshold).toHaveLength(1);
    expect(c.missing.join(" ")).toContain("плательщик ли НДС");
  });

  it("плательщику НДС и не на УСН порог не нужен: книги не читаются", async () => {
    for (const extra of [{ ПлательщикНДСПрименяющийУСН: true }, { СистемаНалогообложения: "Общая" }]) {
      const c = await calendar({ [SET.settings]: [setting(extra)] });
      expect(c.vatThreshold).toBeUndefined();
      expect(turnovers).toEqual([]);
    }
  });
});
