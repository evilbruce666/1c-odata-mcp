import type { Connection } from "../context.js";
import { REGISTERS, TAX_PATENT_DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { tradeWithoutAccounting } from "./configuration.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid, type QueryOptions } from "./query.js";
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
  /** В настройках «Упрощенная» с признаком «Применяется АУСН»: это автоматизированная УСН, а не обычная. */
  ausn?: true;
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
  /** Записи настроек есть, но все позже даты: с какой даты они начинаются. */
  settingsStart?: string;
  /** АУСН или налог на профессиональный доход: взносы за себя ИП не платит. */
  noContributions?: boolean;
}

const SETTINGS_LABEL =
  "регистр сведений «Система налогообложения» (InformationRegister_НастройкиСистемыНалогообложения)";
const PUBLISH_HINT =
  `Опубликуйте ${SETTINGS_LABEL} в «Составе OData», затем вызовите read.system.health_check (он перечитывает ` +
  "схему) и повторите запрос.";

/**
 * Старше этого расчёт УСН не используется: режим с тех пор мог смениться. 18 месяцев — потому что у ИП год закрывают
 * иногда только к декларации (28 апреля), и расчёту за позапрошлый конец года тогда около 16 месяцев: при меньшем
 * пороге календарь терял бы сроки УСН как раз перед ними (замер дат создания расчётов на живых базах, 11.2026).
 * От недавней смены режима порог не защищает — только регистр настроек; поэтому расчёт старше года помечается.
 */
const CALCULATION_MAX_AGE_MONTHS = 18;
const CALCULATION_CHECK_AGE_MONTHS = 12;

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

/** Столько смен режима за историю организации с большим запасом. */
const SETTINGS_ROWS = 100;

/** Все записи настроек организации не позже даты, новые первыми. */
async function settingsHistory(
  conn: Connection,
  set: string,
  orgRef: string,
  until: string,
): Promise<ODataEntity[]> {
  const select = await presentFields(conn, set, [
    "Period",
    "СистемаНалогообложения",
    "ПрименяетсяУСНПатент",
    "ПрименяетсяАУСН",
    "ПрименяетсяНалогНаПрофессиональныйДоход",
    "ОбъектНалогообложенияУСН",
    "ПлательщикНДСПрименяющийУСН",
  ]);
  // Записи с будущей датой — запланированная смена режима, текущим он ещё не стал.
  const { rows } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(cmp("Организация_Key", "eq", odataGuid(orgRef)), cmp("Period", "le", endOfDay(until))),
      select,
      orderby: "Period desc",
    },
    SETTINGS_ROWS,
    SETTINGS_ROWS,
  );
  return rows;
}

/** rows — история настроек, новые первыми. */
function fromSettings(rows: readonly ODataEntity[], asOf: string): TaxRegimeResult {
  const row = rows.find((r) => text(r["Period"]).slice(0, 10) <= asOf);
  if (!row)
    return {
      notes: [`В регистре «Система налогообложения» нет записей этой организации на ${asOf}.`],
      ...(dateOnly(rows.at(-1)?.["Period"]) ? { settingsStart: dateOnly(rows.at(-1)!["Period"])! } : {}),
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
    notes.push("Ключи режима выводятся только для УСН; значение настроек 1С — в поле system.");
  }
  return {
    regime: {
      regimes,
      source: "settings",
      asOf,
      ...(system ? { system } : {}),
      ...(row["ПрименяетсяАУСН"] === true ? { ausn: true as const } : {}),
      ...(dateOnly(row["Period"]) ? { settingsFrom: dateOnly(row["Period"])! } : {}),
      ...(usn && typeof row["ПлательщикНДСПрименяющийУСН"] === "boolean"
        ? { vatPayerOnUsn: row["ПлательщикНДСПрименяющийУСН"] }
        : {}),
    },
    notes,
    ...(row["ПрименяетсяАУСН"] === true ||
    row["ПрименяетсяНалогНаПрофессиональныйДоход"] === true ||
    system === "АУСН" ||
    system === "НалогНаПрофессиональныйДоход"
      ? { noContributions: true }
      : {}),
  };
}

/** Кварталов расчёта и патентов за историю организации с запасом. */
const HISTORY_ROWS = 200;

interface Traces {
  /** Строки расчёта УСН, новые первыми. */
  calculations: ODataEntity[];
  patents: ODataEntity[];
}

/** Следы режима не позже until — по одному запросу на расчёт и на патенты, сколько бы дат ни спрашивали. */
async function traces(conn: Connection, orgRef: string, until: string): Promise<Traces> {
  const available = await conn.available();
  const org = cmp("Организация_Key", "eq", odataGuid(orgRef));
  const read = async (set: string | undefined, options: QueryOptions): Promise<ODataEntity[]> =>
    set ? (await fetchAll(conn.client, set, options, HISTORY_ROWS, HISTORY_ROWS)).rows : [];
  const calculations = await read(resolveEntity(REGISTERS.usnCalculation, available), {
    filter: and(org, cmp("ПериодРасчета", "le", endOfDay(until))),
    select: ["ПериодРасчета", "ОбъектНалогообложения"],
    orderby: "ПериодРасчета desc",
  });
  const patents = await read(resolveEntity(TAX_PATENT_DOCUMENTS, available), {
    filter: and(org, "Posted eq true", cmp("ДатаНачала", "le", endOfDay(until))),
    select: ["ДатаНачала", "ДатаОкончания"],
    orderby: "ДатаНачала desc",
  });
  return { calculations, patents };
}

function inferred({ calculations, patents }: Traces, asOf: string): TaxRegimeResult {
  const regimes: TaxRegimeKey[] = [];
  const notes: string[] = [];
  let calculatedFor: string | undefined;

  const known = calculations.filter((r) => (dateOnly(r["ПериодРасчета"]) ?? "9999") <= asOf);
  const latest = dateOnly(known[0]?.["ПериодРасчета"]);
  if (latest && latest < monthsBefore(asOf, CALCULATION_MAX_AGE_MONTHS)) {
    notes.push(
      `Последний расчёт УСН — за период до ${latest}: он слишком старый, чтобы судить о режиме сейчас.`,
    );
  } else if (latest) {
    // Две строки одного периода возможны (деятельность с торговым сбором и без) — их объекты должны совпадать.
    const objects = new Set(
      known
        .filter((r) => dateOnly(r["ПериодРасчета"]) === latest)
        .map((r) => text(r["ОбъектНалогообложения"])),
    );
    const key = objects.size === 1 ? USN_OBJECTS[[...objects][0]!] : undefined;
    if (key) {
      regimes.push(key);
      calculatedFor = latest;
      if (latest < monthsBefore(asOf, CALCULATION_CHECK_AGE_MONTHS))
        notes.push(
          `Расчёт УСН, по которому определён режим, — за период до ${latest}, ему больше года: проверьте режим.`,
        );
    } else
      notes.push(`В расчёте УСН за период до ${latest} объект налогообложения не определяется однозначно.`);
  }

  const active = (p: ODataEntity) =>
    (dateOnly(p["ДатаНачала"]) ?? "9999") <= asOf && (dateOnly(p["ДатаОкончания"]) ?? "") >= asOf;
  if (patents.some(active)) regimes.push("psn");

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
 * Читатель режима организации на разные даты не позже until: история настроек (а без неё — расчёт УСН и патенты)
 * читается один раз. Пустой результат без заметок — вопрос к этой базе не относится (в Управлении торговлей
 * налогового учёта нет).
 */
export async function taxRegimeReader(
  conn: Connection,
  orgRef: string,
  until: string,
): Promise<(asOf: string) => Promise<TaxRegimeResult>> {
  const settings = await publishedSettings(conn);
  if (settings) {
    const rows = await settingsHistory(conn, settings, orgRef, until);
    return (asOf) => Promise.resolve(fromSettings(rows, asOf));
  }
  if (await tradeWithoutAccounting(conn)) return () => Promise.resolve({ notes: [] });
  const found = await traces(conn, orgRef, until);
  return (asOf) => Promise.resolve(inferred(found, asOf));
}

const publishedSettings = async (conn: Connection) =>
  resolveEntity(REGISTERS.taxSettings, await conn.available());

/**
 * «Плательщик НДС» по настройкам 1С — для случая, когда режим назвал пользователь: это отдельный факт 1С, и названный
 * режим его не отменяет. undefined — настройки не опубликованы или этого не говорят.
 */
export async function vatPayerFromSettings(
  conn: Connection,
  orgRef: string,
  asOf: string,
): Promise<boolean | undefined> {
  const settings = await publishedSettings(conn);
  if (!settings) return undefined;
  return fromSettings(await settingsHistory(conn, settings, orgRef, asOf), asOf).regime?.vatPayerOnUsn;
}

/** Режим организации на дату. */
export async function readTaxRegime(
  conn: Connection,
  orgRef: string,
  asOf: string,
): Promise<TaxRegimeResult> {
  return (await taxRegimeReader(conn, orgRef, asOf))(asOf);
}
