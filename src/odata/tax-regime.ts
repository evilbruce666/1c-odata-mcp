import type { Connection } from "../context.js";
import { REGISTERS, TAX_PATENT_DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { tradeWithoutAccounting } from "./configuration.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid } from "./query.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Режим налогообложения организации — одна функция на всех потребителей (карточка организации, налоговый
 * календарь).
 *
 * Основной источник — периодический регистр «Система налогообложения». Если он не опубликован, режим выводится
 * косвенно: объект УСН — из расчёта налога, патент — из операции с патентом, действующей на дату. Косвенный вывод
 * всегда помечен, а отсутствие следов УСН не означает ОСНО.
 */

/** Ключи режимов — те же, что в regimes у правил tax-calendar.json. */
export const TAX_REGIME_KEYS = ["usn_income", "usn_income_expense", "psn", "osno"] as const;
export type TaxRegimeKey = (typeof TAX_REGIME_KEYS)[number];

export interface TaxRegime {
  /** Ключи выведены только для УСН и патента; другие системы называет system. */
  regimes: TaxRegimeKey[];
  source: "settings" | "inferred";
  asOf: string;
  /** Значение 1С как есть: «Упрощенная», «Общая», «НалогНаПрофессиональныйДоход»… */
  system?: string;
  /** С какой даты действует запись настроек (source=settings). */
  settingsFrom?: string;
  /** На какой расчётный период УСН опирается косвенный вывод (source=inferred). */
  calculatedFor?: string;
  vatPayerOnUsn?: boolean;
}

export interface TaxRegimeResult {
  regime?: TaxRegime;
  /** Что неясно и что сделать, чтобы стало ясно. */
  notes: string[];
}

const SETTINGS_LABEL =
  "регистр сведений «Система налогообложения» (InformationRegister_НастройкиСистемыНалогообложения)";
const PUBLISH_HINT =
  `Опубликуйте ${SETTINGS_LABEL} в «Составе OData», затем вызовите read.system.health_check (он перечитывает ` +
  "схему) и повторите запрос.";

/**
 * Старше этого расчёт УСН не используется: режим с тех пор мог смениться, а календарь по старому режиму
 * подсказал бы чужие сроки. Годовой расчёт делается раз в квартал, 15 месяцев — с запасом на позднее закрытие.
 */
const CALCULATION_MAX_AGE_MONTHS = 15;

const USN_OBJECTS: Record<string, TaxRegimeKey> = {
  Доходы: "usn_income",
  ДоходыМинусРасходы: "usn_income_expense",
};

const text = (v: unknown): string => (typeof v === "string" ? v : "");
const dateOnly = (v: unknown): string | undefined => {
  const s = text(v);
  return !s || s.startsWith("0001-01-01") ? undefined : s.slice(0, 10);
};
const endOfDay = (ymd: string) => `datetime'${ymd}T23:59:59'`;
const startOfDay = (ymd: string) => `datetime'${ymd}T00:00:00'`;

function monthsBefore(ymd: string, months: number): string {
  const d = new Date(`${ymd}T00:00:00Z`);
  d.setUTCMonth(d.getUTCMonth() - months);
  return d.toISOString().slice(0, 10);
}

/** Из перечисленных полей — те, что есть в $metadata этой базы: состав регистра зависит от релиза. */
async function presentFields(conn: Connection, set: string, wanted: readonly string[]): Promise<string[]> {
  const names = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);
  return wanted.filter((f) => names.has(f));
}

async function fromSettings(
  conn: Connection,
  set: string,
  orgRef: string,
  asOf: string,
): Promise<TaxRegimeResult> {
  const select = await presentFields(conn, set, [
    "Period",
    "СистемаНалогообложения",
    "ПрименяетсяУСНПатент",
    "ПрименяетсяАУСН",
    "ОбъектНалогообложенияУСН",
    "ПлательщикНДСПрименяющийУСН",
  ]);
  // Записи с будущей датой — запланированная смена режима, текущим он ещё не стал.
  const { rows } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(cmp("Организация_Key", "eq", odataGuid(orgRef)), cmp("Period", "le", endOfDay(asOf))),
      select,
      orderby: "Period desc",
    },
    1,
    1,
  );
  const row: ODataEntity | undefined = rows[0];
  if (!row)
    return {
      notes: [`В регистре «Система налогообложения» нет записей этой организации на ${asOf}.`],
    };

  const system = text(row["СистемаНалогообложения"]);
  const regimes: TaxRegimeKey[] = [];
  const notes: string[] = [];
  const usn = system === "Упрощенная" && row["ПрименяетсяАУСН"] !== true;
  if (usn) {
    const key = USN_OBJECTS[text(row["ОбъектНалогообложенияУСН"])];
    if (key) regimes.push(key);
    else notes.push("Объект налогообложения УСН в настройках не заполнен.");
    if (row["ПрименяетсяУСНПатент"] === true) regimes.push("psn");
  } else {
    notes.push("Ключи режима выводятся только для УСН; значение настроек 1С — в taxRegime.system.");
  }
  return {
    regime: {
      regimes,
      source: "settings",
      asOf,
      ...(system ? { system } : {}),
      ...(dateOnly(row["Period"]) ? { settingsFrom: dateOnly(row["Period"])! } : {}),
      ...(usn && typeof row["ПлательщикНДСПрименяющийУСН"] === "boolean"
        ? { vatPayerOnUsn: row["ПлательщикНДСПрименяющийУСН"] }
        : {}),
    },
    notes,
  };
}

async function inferred(conn: Connection, orgRef: string, asOf: string): Promise<TaxRegimeResult> {
  const available = await conn.available();
  const calcSet = resolveEntity(REGISTERS.usnCalculation, available);
  const patentSet = resolveEntity(TAX_PATENT_DOCUMENTS, available);
  const org = cmp("Организация_Key", "eq", odataGuid(orgRef));
  const regimes: TaxRegimeKey[] = [];
  const notes: string[] = [];
  let calculatedFor: string | undefined;

  if (calcSet) {
    // Две строки одного периода возможны (деятельность с торговым сбором и без) — их объекты должны совпадать.
    const { rows } = await fetchAll(
      conn.client,
      calcSet,
      {
        filter: and(org, cmp("ПериодРасчета", "le", endOfDay(asOf))),
        select: ["ПериодРасчета", "ОбъектНалогообложения"],
        orderby: "ПериодРасчета desc",
      },
      2,
      2,
    );
    const latest = dateOnly(rows[0]?.["ПериодРасчета"]);
    if (latest && latest < monthsBefore(asOf, CALCULATION_MAX_AGE_MONTHS)) {
      notes.push(
        `Последний расчёт УСН — за период до ${latest}: он слишком старый, чтобы судить о режиме сейчас.`,
      );
    } else if (latest) {
      const objects = new Set(
        rows
          .filter((r) => dateOnly(r["ПериодРасчета"]) === latest)
          .map((r) => text(r["ОбъектНалогообложения"])),
      );
      const key = objects.size === 1 ? USN_OBJECTS[[...objects][0]!] : undefined;
      if (key) {
        regimes.push(key);
        calculatedFor = latest;
      } else
        notes.push(`В расчёте УСН за период до ${latest} объект налогообложения не определяется однозначно.`);
    }
  }

  if (patentSet) {
    const { rows } = await fetchAll(
      conn.client,
      patentSet,
      {
        filter: and(
          org,
          "Posted eq true",
          cmp("ДатаНачала", "le", endOfDay(asOf)),
          cmp("ДатаОкончания", "ge", startOfDay(asOf)),
        ),
        select: ["Ref_Key"],
      },
      1,
      1,
    );
    if (rows.length) regimes.push("psn");
  }

  if (!regimes.length) {
    notes.push(`Режим налогообложения не определён. ${PUBLISH_HINT}`);
    return { notes };
  }
  notes.push(
    "Режим определён косвенно, по расчёту УСН и действующим патентам; другие системы так не видны. " +
      PUBLISH_HINT,
  );
  return {
    regime: { regimes, source: "inferred", asOf, ...(calculatedFor ? { calculatedFor } : {}) },
    notes,
  };
}

/**
 * Режим организации на дату. Пустой результат без заметок — вопрос к этой базе не относится (в Управлении
 * торговлей налогового учёта нет).
 */
export async function readTaxRegime(
  conn: Connection,
  orgRef: string,
  asOf: string,
): Promise<TaxRegimeResult> {
  const settings = resolveEntity(REGISTERS.taxSettings, await conn.available());
  if (settings) return fromSettings(conn, settings, orgRef, asOf);
  if (await tradeWithoutAccounting(conn)) return { notes: [] };
  return inferred(conn, orgRef, asOf);
}
