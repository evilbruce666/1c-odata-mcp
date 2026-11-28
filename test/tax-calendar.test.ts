import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * tax-calendar.json — сроки, которых в 1С может ещё не быть. Справочник устаревает молча, а ошибка в дате стоит
 * пеней, поэтому тест пересчитывает перенос каждой даты независимо от записанной, следит за полнотой набора по
 * годам и в декабре краснеет, если нет правил на следующий год.
 */
type Rule = {
  id: string;
  obligation: string;
  kind: string;
  title: string;
  taxpayers: string[];
  regimes: string[];
  period: { from: string; to: string };
  nominal: string;
  due: string;
  basis: string;
  condition?: string;
  match?: Record<string, unknown>;
};
const raw = readFileSync("tax-calendar.json", "utf8");
const data = JSON.parse(raw) as {
  dueYears: number[];
  notCovered: string[];
  workingWeekends: string[];
  nonWorkingWeekdays: Record<string, string[]>;
  sources: Record<string, string>;
  rules: Rule[];
};

const day = (iso: string) => new Date(`${iso}T00:00:00Z`);
const iso = (d: Date) => d.toISOString().slice(0, 10);
const isWeekend = (d: string) => day(d).getUTCDay() === 0 || day(d).getUTCDay() === 6;
function isWorking(d: string): boolean {
  if (data.workingWeekends.includes(d)) return true;
  return !isWeekend(d) && !data.nonWorkingWeekdays[d.slice(0, 4)]!.includes(d);
}
/** П. 7 ст. 6.1 НК РФ: срок, выпавший на нерабочий день, переносится на ближайший рабочий. */
function shifted(nominal: string): string {
  const d = day(nominal);
  while (!isWorking(iso(d))) d.setUTCDate(d.getUTCDate() + 1);
  return iso(d);
}
/** С декабря следующий год обязан быть в справочнике: окно календаря уже заходит в него. */
function requiredYears(today: Date): number[] {
  const y = today.getUTCFullYear();
  return today.getUTCMonth() === 11 ? [y, y + 1] : [y];
}

describe("tax-calendar.json", () => {
  it("в справочнике есть текущий год, а с декабря — и следующий", () => {
    const missing = requiredYears(new Date()).filter((y) => !data.dueYears.includes(y));
    expect(missing, "добавьте в tax-calendar.json правила на эти годы по НК РФ").toEqual([]);
  });

  it("пороги НДС есть на каждый год сроков и на год перед ним", () => {
    const { inYear, previousYear } = (
      data as unknown as {
        vatThreshold: Record<"inYear" | "previousYear", { limits: Record<string, number> }>;
      }
    ).vatThreshold;
    for (const year of requiredYears(new Date())) {
      expect(inYear.limits[String(year)], `в течение ${year}`).toBeGreaterThan(0);
      expect(previousYear.limits[String(year - 1)], `доход за ${year - 1}`).toBeGreaterThan(0);
    }
  });

  it("декабрь требует следующий год", () => {
    expect(requiredYears(day("2027-11-30"))).toEqual([2027]);
    expect(requiredYears(day("2027-12-01"))).toEqual([2027, 2028]);
  });

  it("производственный календарь: 247 рабочих дней в 2026 и 2027 годах", () => {
    for (const year of ["2026", "2027"]) {
      let working = 0;
      for (const d = day(`${year}-01-01`); iso(d).startsWith(year); d.setUTCDate(d.getUTCDate() + 1))
        if (isWorking(iso(d))) working++;
      expect(working, year).toBe(247);
      expect(data.nonWorkingWeekdays[year]!.filter(isWeekend)).toEqual([]);
    }
  });

  it("каждый срок равен дате по статье с переносом с нерабочего дня", () => {
    const wrong = data.rules
      .filter(
        (r) => data.nonWorkingWeekdays[r.nominal.slice(0, 4)] === undefined || shifted(r.nominal) !== r.due,
      )
      .map((r) => `${r.id}: ${r.nominal} → ${r.due}`);
    expect(wrong).toEqual([]);
  });

  it("контрольные даты: переносы, которые легко пропустить", () => {
    const due = Object.fromEntries(data.rules.map((r) => [r.id, r.due]));
    // 28.03.2026 — суббота: годовой налог УСН организации. Совпадает со сроком, который считает 1С.
    expect(due["usn_tax-org-2025-12-31"]).toBe("2026-03-30");
    // 25.04.2026 — суббота: декларация ИП.
    expect(due["usn_declaration-ip-2025-12-31"]).toBe("2026-04-27");
    // В 2027 году на воскресенье выпадают ровно три даты по статьям.
    const sundays = data.rules.filter(
      (r) => r.nominal.startsWith("2027") && day(r.nominal).getUTCDay() === 0,
    );
    expect([...new Set(sundays.map((r) => r.nominal))].sort()).toEqual([
      "2027-03-28",
      "2027-04-25",
      "2027-07-25",
    ]);
  });

  it("у каждого года полный набор обязанностей организации и ИП", () => {
    const common = { usn_declaration: 1, usn_tax: 1, usn_advance_notice: 3, usn_advance: 3 };
    const expected = {
      organization: common,
      ip: { ...common, ip_fixed_contributions: 1, ip_extra_contribution: 1 },
    };
    for (const year of data.dueYears) {
      for (const [taxpayer, want] of Object.entries(expected)) {
        const got: Record<string, number> = {};
        for (const r of data.rules)
          if (r.due.startsWith(String(year)) && r.taxpayers.includes(taxpayer))
            got[r.obligation] = (got[r.obligation] ?? 0) + 1;
        expect(got, `${year} ${taxpayer}`).toEqual(want);
      }
    }
  });

  it("записи не повторяются и ссылаются на статью с известным источником", () => {
    expect(new Set(data.rules.map((r) => r.id)).size).toBe(data.rules.length);
    const keys = data.rules.map((r) => [r.obligation, r.period.to, ...r.taxpayers].join(" "));
    expect(new Set(keys).size).toBe(keys.length);
    for (const r of data.rules) {
      expect(["report", "notice", "payment"], r.id).toContain(r.kind);
      expect(data.dueYears.map(String), r.id).toContain(r.due.slice(0, 4));
      expect(r.period.from <= r.period.to, r.id).toBe(true);
      // Фиксированные взносы платятся до конца своего года (28 декабря), остальное — после окончания периода.
      if (r.obligation === "ip_fixed_contributions")
        expect(r.nominal, r.id).toBe(`${r.period.to.slice(0, 4)}-12-28`);
      else expect(r.period.to < r.nominal, r.id).toBe(true);
      const articles = r.basis.match(/ст\. \d+(\.\d+)?/g) ?? [];
      expect(articles.length, r.id).toBeGreaterThan(0);
      // ст. 430 задаёт сам взнос 1%, срок — в ст. 432.
      expect(
        articles.some((a) => data.sources[a]),
        r.id,
      ).toBe(true);
    }
  });

  it("уведомления: только при сумме к уплате и только по авансам УСН", () => {
    const notices = data.rules.filter((r) => r.kind === "notice");
    expect(notices.filter((r) => r.obligation !== "usn_advance_notice")).toEqual([]);
    for (const r of notices) expect(r.condition, r.id).toContain("больше нуля");
  });

  it("сумм налогов в правилах нет: суммы считает 1С", () => {
    const allowed = new Set(
      "id obligation kind title taxpayers regimes period nominal due basis condition match".split(" "),
    );
    for (const r of data.rules)
      expect(
        Object.keys(r).filter((k) => !allowed.has(k)),
        r.id,
      ).toEqual([]);
    // Фиксированный взнос и предел 1% меняются каждый год; единственное допустимое число — порог 300 000 ₽.
    const rubles = raw.match(/\d[\d  ]{2,}\s?(₽|руб)/g) ?? [];
    expect(rubles.filter((s) => !s.replace(/\D/g, "").startsWith("300000"))).toEqual([]);
  });

  it("сроки патента не записаны датами: они зависят от самого патента", () => {
    expect(data.rules.filter((r) => /патент/i.test(r.title))).toEqual([]);
    expect(data.notCovered.some((s) => s.startsWith("Сроки патента"))).toBe(true);
  });

  it("справочник входит в npm-пакет", () => {
    expect((JSON.parse(readFileSync("package.json", "utf8")) as { files: string[] }).files).toContain(
      "tax-calendar.json",
    );
  });
});
