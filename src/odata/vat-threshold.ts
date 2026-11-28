import type { TaxRules } from "./tax-calendar-rules.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Порог НДС на УСН (ст. 145 НК РФ): доход по книгам доходов 1С рядом с порогом из справочника.
 *
 * Доход — обороты, которые считает 1С: книга доходов и расходов УСН (графа 5 — доходы, учитываемые в базе; на живой
 * базе она совпала с «Доходы» расчёта УСН) и, у ИП, книга доходов по патенту — для порога доходы складываются.
 * Расчёт УСН источником не взят: его строки появляются только при закрытии квартала. «Приближается к порогу» не
 * бывает: это суждение, а не факт.
 */

export interface VatThresholdFact {
  year: number;
  /** По какую дату взяты обороты. */
  through: string;
  limit: number;
  /** Что значит превышение и по какой норме. */
  appliesTo: string;
  basis: string;
  incomeUsn: number;
  incomePatent?: number;
  /** incomeUsn плюс incomePatent. */
  income: number;
  /** Нет — не известно: доход ниже порога, но учтён не весь. */
  exceeded?: boolean;
}

export type IncomeBook = "usn" | "patent";

export interface VatThresholdInput {
  taxpayer: "ip" | "organization";
  today: string;
  registered?: string | undefined;
  /** «Плательщик НДС» из настроек 1С; undefined — настройки этого не говорят. */
  vatPayer?: boolean | undefined;
  thresholds: TaxRules["vatThreshold"];
  /** Обороты книги за год по дату; undefined — книгу прочитать нельзя (причину пишет вызывающий). */
  turnover: (book: IncomeBook, from: string, to: string) => Promise<ODataEntity[] | undefined>;
}

export interface VatThresholdResult {
  facts: VatThresholdFact[];
  warnings: string[];
  notes: string[];
}

const money = (n: number) => Math.round(n * 100) / 100;
const rub = (n: number) =>
  money(n)
    .toLocaleString("ru-RU")
    .replace(/\u00a0/g, " ");
const total = (rows: readonly ODataEntity[], field: string) =>
  money(rows.reduce((sum, r) => sum + (Number(r[field]) || 0), 0));

export async function readVatThreshold(input: VatThresholdInput): Promise<VatThresholdResult> {
  const { taxpayer, today, thresholds } = input;
  const result: VatThresholdResult = { facts: [], warnings: [], notes: [] };
  const current = Number(today.slice(0, 4));
  for (const year of [current - 1, current]) {
    const ended = year < current;
    const through = ended ? `${year}-12-31` : today;
    if (input.registered && input.registered > through) continue;
    const rule = ended ? thresholds.previousYear : thresholds.inYear;
    const limit = rule.limits[String(year)];
    if (limit === undefined) {
      result.notes.push(`Порог НДС за ${year} год не проверен: в справочнике его нет. Обновите коннектор.`);
      continue;
    }
    const usn = await input.turnover("usn", `${year}-01-01`, through);
    if (!usn) return result;
    const patent = taxpayer === "ip" ? await input.turnover("patent", `${year}-01-01`, through) : [];
    // 1С не отдаёт строку, где все запрошенные ресурсы нулевые: нет строк — в книгах за год пусто.
    if (!usn.length && !patent?.length) {
      result.notes.push(
        `Порог НДС за ${year} год не проверен: ` +
          (patent
            ? "в книгах доходов 1С за этот год нет записей."
            : "в книге доходов УСН записей за этот год нет, а книга патента не прочитана."),
      );
      continue;
    }
    const incomeUsn = total(usn, "Графа5Turnover");
    const incomePatent = taxpayer === "ip" && patent ? total(patent, "Графа4Turnover") : undefined;
    const income = money(incomeUsn + (incomePatent ?? 0));
    const exceeded = income > limit;
    const appliesTo = ended
      ? `освобождение от НДС в ${year + 1} году по доходу за ${year}`
      : `превышение в течение ${year} года`;
    result.facts.push({
      year,
      through,
      limit,
      appliesTo,
      basis: rule.basis,
      incomeUsn,
      ...(incomePatent !== undefined ? { incomePatent } : {}),
      income,
      // Без патентной книги «не превышен» было бы утверждением по неполному доходу.
      ...(exceeded || patent ? { exceeded } : {}),
    });
    if (!exceeded) continue;
    const consequence = ended
      ? `по этому основанию освобождения от НДС в ${year + 1} году нет`
      : "обязанности плательщика НДС начинаются с 1-го числа месяца после месяца превышения (какой это " +
        "месяц, календарь не определяет)";
    result.warnings.push(
      `Порог НДС на УСН: доход за ${year} год по книгам 1С на ${through} — ${rub(income)} ₽, больше ` +
        `${rub(limit)} ₽ — ${consequence} (${rule.basis}).` +
        (input.vatPayer === false ? " В настройках 1С организация не плательщик НДС — проверьте." : "") +
        (year === 2025 && taxpayer === "ip"
          ? " За 2025 год у части ИП из этого дохода исключаются проценты по вкладам — см. п. 1 ст. 145."
          : ""),
    );
  }
  if (result.facts.length)
    result.notes.push(
      "Порог НДС: доход — обороты книг доходов 1С на дату (записи появляются при проведении документов и " +
        "закрытии месяца). Курсовые разницы и субсидии, которые статья не учитывает, не вычтены; доходы на АУСН " +
        "не прибавлены.",
    );
  return result;
}
