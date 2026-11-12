import {
  accountingChart,
  accountingRegister,
  defaultAccountCodes,
  type AccountDefault,
} from "../config/accounting.js";
import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid, odataString, or } from "./query.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { ensurePublished, requireEntity } from "./publication.js";
import { lookupWithRefresh } from "./cache.js";
import { buildQuery } from "./query.js";
import {
  AggregateOverflowError,
  addMeta,
  emptyMeta,
  fetchAllForAggregation,
  type ScanMeta,
} from "./aggregate.js";
import type { EntityMeta, ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";
import { ODataError } from "./errors.js";

/**
 * Аналитика для 1С:Бухгалтерия 3.0 строится на регистре бухгалтерии
 * «Хозрасчетный» и его виртуальной таблице Balance. И дебиторка (сч. 62),
 * и остатки товаров (сч. 41) — это сальдо по соответствующему счёту,
 * с контрагентом/номенклатурой в ExtDimension1.
 */

export interface Account {
  key: string;
  code: string;
  description: string;
}

/**
 * План счетов целиком — из кэша подключения (cache.ts): он меняется редко, а нужен почти каждому
 * инструменту. Несколько сотен строк, одним запросом с громким переполнением вместо усечения.
 */
async function loadChart(conn: Connection): Promise<Account[]> {
  const chart = await requireEntity(conn, accountingChart(conn.cfg), "План счетов «Хозрасчётный»");
  const { rows } = await fetchAllForAggregation(
    conn,
    chart,
    { select: ["Ref_Key", "Code", "Description"], orderby: "Code" },
    "план счетов",
  );
  return rows.map((r) => ({
    key: String(r["Ref_Key"] ?? ""),
    code: String(r["Code"] ?? ""),
    description: String(r["Description"] ?? ""),
  }));
}

const chartKey = (conn: Connection): string => `chart:${accountingChart(conn.cfg).join("|")}`;

function fromChart<R>(
  conn: Connection,
  pick: (chart: Account[]) => R,
  missing: (r: R) => boolean,
): Promise<R> {
  return lookupWithRefresh(conn, chartKey(conn), () => loadChart(conn), pick, missing);
}

/** Возвращает счета, код которых начинается с одного из префиксов (напр. "62", "41"). */
export async function resolveAccounts(conn: Connection, prefixes: readonly string[]): Promise<Account[]> {
  return fromChart(
    conn,
    (chart) => chart.filter((a) => prefixes.some((p) => a.code.startsWith(p))),
    (r) => r.length === 0,
  );
}

/**
 * Возвращает карту «код счёта → Ref_Key» для заданных точных кодов
 * (напр. ["41.01","60.01","90.01.1"]). Нужно для заполнения счетов учёта
 * в документах — через OData автозаполнение 1С не срабатывает.
 */
export async function accountsByCode(
  conn: Connection,
  codes: readonly string[],
): Promise<Map<string, string>> {
  const wanted = new Set(codes);
  if (wanted.size === 0) return new Map();
  return fromChart(
    conn,
    (chart) => new Map(chart.filter((a) => wanted.has(a.code)).map((a) => [a.code, a.key])),
    // Коды — кандидаты по порядку предпочтения, часть может отсутствовать законно: перечитываем,
    // только если не нашлось ни одного.
    (r) => r.size === 0,
  );
}

/** Счета по умолчанию только для нужных групп; общая точка для будущего поиска настройки расчётов. */
export async function resolveDefaultAccounts(
  conn: Connection,
  groups: readonly AccountDefault[],
): Promise<Map<AccountDefault, string | undefined>> {
  const candidates = [...new Set(groups)].map(
    (group) => [group, defaultAccountCodes(conn.cfg, group)] as const,
  );
  const refs = await accountsByCode(conn, [...new Set(candidates.flatMap(([, codes]) => [...codes]))]);
  const result = new Map<AccountDefault, string | undefined>();
  for (const [group, codes] of candidates) {
    const ref = codes.map((code) => refs.get(code)).find((value) => value !== undefined);
    if (!ref && conn.cfg?.accounting?.strictDefaults)
      throw new InputError(
        `В плане счетов нет счёта из accounting.defaults.${group} — проверьте коды счетов этой базы.`,
      );
    result.set(group, ref);
  }
  return result;
}

/** Счета учёта номенклатуры (Ref_Key каждого; undefined если не задан). */
export interface NomAccounts {
  goods?: string; // СчетУчета (41.xx)
  incomingVat?: string; // НДС по приобретённым ценностям (19.xx)
  outgoingVat?: string; // НДС по реализации (90.03)
  income?: string; // доходы от реализации (90.01.x)
  expense?: string; // расходы от реализации / себестоимость (90.02.x)
}

const NOM_ACCOUNTS_REG = ["InformationRegister_СчетаУчетаНоменклатуры"] as const;
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";
const norm = (g: unknown): string | undefined => {
  const s = typeof g === "string" ? g : "";
  return s && s !== EMPTY_GUID ? s : undefined;
};

/**
 * Берёт счета учёта для номенклатуры из регистра «Счета учёта номенклатуры»,
 * выбирая самую специфичную подходящую запись (пустые измерения — «джокер»).
 * Возвращает undefined, если регистр не опубликован (вызывающий откатится на коды).
 */
export async function nomenclatureAccounts(
  conn: Connection,
  orgKey: string,
  nomRef: string,
): Promise<NomAccounts | undefined> {
  const available = await conn.available();
  const reg = resolveEntity(NOM_ACCOUNTS_REG, available);
  if (!reg) return undefined;

  // Вид номенклатуры нужен для матчинга измерения ВидНоменклатуры в регистре.
  let vidRef: string | undefined;
  const nomSet = resolveEntity(CATALOGS.nomenclature, available);
  if (nomSet) {
    try {
      const item = await conn.client.getEntity(
        `${nomSet}(guid'${nomRef.replace(/[{}']/g, "")}')${buildQuery({ select: ["ВидНоменклатуры_Key"] })}`,
      );
      vidRef = norm(item["ВидНоменклатуры_Key"]);
    } catch {
      // нет такого поля/объекта — матчим без вида
    }
  }

  const { rows } = await fetchAll(
    conn.client,
    reg,
    {
      select: [
        "Организация_Key",
        "Номенклатура_Key",
        "ВидНоменклатуры_Key",
        "СчетУчета_Key",
        "СчетУчетаНДСПоПриобретеннымЦенностям_Key",
        "СчетУчетаНДСПоРеализации_Key",
        "СчетДоходовОтРеализации_Key",
        "СчетРасходовОтРеализации_Key",
      ],
    },
    conn.behavior.pageSize,
    500,
  );
  let best: ODataEntity | undefined;
  let bestScore = -1;
  for (const r of rows) {
    const nm = norm(r["Номенклатура_Key"]);
    const og = norm(r["Организация_Key"]);
    const vd = norm(r["ВидНоменклатуры_Key"]);
    if (nm && nm !== nomRef) continue;
    if (og && og !== orgKey) continue;
    if (vd && vd !== vidRef) continue; // запись для другого вида номенклатуры
    const score = (nm ? 4 : 0) + (vd ? 2 : 0) + (og ? 1 : 0);
    if (score > bestScore) {
      bestScore = score;
      best = r;
    }
  }
  if (!best) return undefined;
  return {
    goods: norm(best["СчетУчета_Key"]),
    incomingVat: norm(best["СчетУчетаНДСПоПриобретеннымЦенностям_Key"]),
    outgoingVat: norm(best["СчетУчетаНДСПоРеализации_Key"]),
    income: norm(best["СчетДоходовОтРеализации_Key"]),
    expense: norm(best["СчетРасходовОтРеализации_Key"]),
  };
}

/**
 * Тянет строки сальдо регистра Хозрасчетный, отфильтрованные по набору счетов
 * и (необязательно) по организации. accountKeys = null — сальдо по всем счетам одним
 * снимком (для проверок, которым нужно много групп счетов сразу).
 */
export async function balanceByAccounts(
  conn: Connection,
  accountKeys: string[] | null,
  orgKey?: string,
  asOf?: string,
): Promise<ODataEntity[]> {
  if (accountKeys?.length === 0) return [];
  const reg = await requireEntity(conn, accountingRegister(conn.cfg), "Регистр бухгалтерии «Хозрасчётный»");
  const filter =
    and(
      accountKeys ? or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))) : undefined,
      orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
    ) || undefined;
  // Параметр Period — НЕ через $filter, а path-параметром у виртуальной таблицы:
  // .../AccountingRegister_Хозрасчетный/Balance(Period=datetime'<следующий день>T00:00:00').
  // Сальдо берётся на НАЧАЛО момента Period, движения в сам момент не входят (проверено на базе:
  // остаток на D 00:00:00 ещё без документа, датированного D 00:00:00). При 23:59:59 терялись бы
  // движения последней секунды дня, а закрытие месяца 1С проводит ровно в 23:59:59.
  // Без параметра виртуальная таблица возвращает текущее сальдо.
  const balancePath = asOf ? `${reg}/Balance(Period=datetime'${nextDay(asOf)}T00:00:00')` : `${reg}/Balance`;
  // Через безопасную выборку: сальдо берём ПОЛНОСТЬЮ (иначе дебиторка/остатки
  // занижаются), с громким переполнением вместо тихой обрезки.
  const read = (f: string | undefined) =>
    fetchAllForAggregation(conn, balancePath, { filter: f }, `сальдо на ${asOf ?? "сейчас"}`);
  if (!accountKeys) return (await read(filter)).rows;
  const { rows } = await bySplittingAccounts(accountKeys, (keys) =>
    read(
      and(
        or(...keys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))),
        orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
      ) || undefined,
    ),
  );
  return rows;
}

/**
 * Выборка по счетам: не успела за таймаут — читаем половинами и складываем. Счета не пересекаются,
 * поэтому строки не задваиваются. Один счёт не успел — настоящая ошибка таймаута.
 */
export async function bySplittingAccounts(
  keys: readonly string[],
  read: (keys: readonly string[]) => Promise<{ rows: ODataEntity[]; meta: ScanMeta }>,
): Promise<{ rows: ODataEntity[]; meta: ScanMeta }> {
  try {
    return await read(keys);
  } catch (e) {
    if (!(e instanceof ODataError) || e.kind !== "timeout" || keys.length <= 1) throw e;
    const middle = Math.ceil(keys.length / 2);
    const a = await bySplittingAccounts(keys.slice(0, middle), read);
    const b = await bySplittingAccounts(keys.slice(middle), read);
    return { rows: [...a.rows, ...b.rows], meta: addMeta(a.meta, b.meta) };
  }
}

/**
 * Виртуальная таблица «Остатки и обороты» регистра бухгалтерии через OData.
 *
 * ВСЁ, что зависит от точного синтаксиса 1С, собрано здесь, в одном месте, чтобы
 * после живого теста править одну константу, а не инструмент:
 *  - имя виртуальной таблицы (по документации платформы — BalanceAndTurnovers;
 *    если конкретная база ответит 404/400 — поменять name, напр. на BalanceAndTurnover);
 *  - имена path-параметров периода (StartPeriod/EndPeriod — как Period у Balance);
 *  - имена ресурсных полей (ресурс «Сумма» + суффиксы виртуальной таблицы).
 */
export const BALANCE_AND_TURNOVERS = {
  name: "BalanceAndTurnovers",
  startParam: "StartPeriod",
  endParam: "EndPeriod",
  fields: {
    openingDr: "СуммаOpeningBalanceDr",
    openingCr: "СуммаOpeningBalanceCr",
    turnoverDr: "СуммаTurnoverDr",
    turnoverCr: "СуммаTurnoverCr",
    closingDr: "СуммаClosingBalanceDr",
    closingCr: "СуммаClosingBalanceCr",
  },
  /** Поле периода: появляется в строках, только если 1С разбила результат по периодичности. */
  periodField: "Period",
} as const;

/**
 * Следующий календарный день для YYYY-MM-DD — чистая арифметика по компонентам,
 * без Date/таймзоны (new Date("…") + toISOString() может «съехать» на сутки).
 * Учитывает конец месяца, года и високосный февраль. Несуществующую дату
 * (напр. 2025-02-30 — формат её пропускает) отвергает, а не «перекатывает».
 */
export function nextDay(ymd: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd);
  if (!m) throw new InputError(`Дата должна быть в формате YYYY-MM-DD: ${ymd}`);
  let y = Number(m[1]);
  let mo = Number(m[2]);
  let d = Number(m[3]);
  const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const dim = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mo - 1];
  if (dim === undefined || d < 1 || d > dim) throw new InputError(`Несуществующая дата: ${ymd}`);
  if (d < dim) d += 1;
  else if (mo < 12) {
    mo += 1;
    d = 1;
  } else {
    y += 1;
    mo = 1;
    d = 1;
  }
  const pad = (n: number, w = 2): string => String(n).padStart(w, "0");
  return `${pad(y, 4)}-${pad(mo)}-${pad(d)}`;
}

/**
 * Путь к виртуальной таблице «Остатки и обороты» за [from,to] (YYYY-MM-DD, оба включительно).
 * Параметры — path-параметрами (как Period у Balance), а не через $filter.
 * EndPeriod — to 23:59:59. В отличие от Period у Balance, EndPeriod в 1С ВКЛЮЧАЕТ свой момент
 * (проверено на базах УТ и БП): при начале следующего дня документ, датированный ровно 00:00:00,
 * входил в этот период и ещё раз в следующий — помесячная ОСВ ooo за март–апрель 2024 давала
 * 48,8 млн Дт-оборотов вместо 41,5. Секунда — шаг времени 1С, движения в 23:59:59 не теряются:
 *   AccountingRegister_Хозрасчетный/BalanceAndTurnovers(StartPeriod=datetime'2026-08-01T00:00:00',EndPeriod=datetime'2026-08-31T23:59:59')
 */
/** Дата как есть, если она существует: «2025-02-30» — отказ, как у nextDay. */
const checkedDay = (ymd: string): string => (nextDay(ymd), ymd);

export function balanceAndTurnoversPath(reg: string, from: string, to: string): string {
  const vt = BALANCE_AND_TURNOVERS;
  return (
    `${reg}/${vt.name}(` +
    `${vt.startParam}=datetime'${from}T00:00:00',` +
    `${vt.endParam}=datetime'${checkedDay(to)}T23:59:59')`
  );
}

/**
 * Проверяет, что строки виртуальной таблицы содержат все шесть ресурсных полей.
 * Иначе (другое имя таблицы/ресурса в конкретной базе) num() тихо дал бы нули,
 * и неизвестная структура выдала бы «ОСВ из нулей» с consistent:true.
 * В ошибку идут только ИМЕНА полей строки, не значения.
 */
export function assertTurnoverFields(rows: readonly ODataEntity[]): void {
  const expected = Object.values(BALANCE_AND_TURNOVERS.fields);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] as Record<string, unknown>;
    const missing = expected.filter((f) => !Object.prototype.hasOwnProperty.call(r, f));
    if (missing.length === 0) continue;
    throw new Error(
      `Виртуальная таблица ${BALANCE_AND_TURNOVERS.name} вернула строки несовместимой структуры ` +
        `(строка #${i + 1}): нет полей ${missing.join(", ")}. ` +
        `Фактические поля строки: ${Object.keys(r).join(", ") || "(нет)"}. ` +
        `Проверьте BALANCE_AND_TURNOVERS в src/odata/accounting.ts.`,
    );
  }
}

/** $filter по набору счетов и (необязательно) организации — фильтрует сама 1С. */
export function accountsOrgFilter(accountKeys: readonly string[], orgKey?: string): string | undefined {
  return (
    and(
      accountKeys ? or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))) : undefined,
      orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
    ) || undefined
  );
}

/**
 * Сколько GUID-ов счетов кладём в один $filter. Широкий префикс (напр. «9») даёт
 * десятки субсчетов — длинный OR упирается в лимит длины URL веб-сервера.
 */
export const ACCOUNT_FILTER_BATCH = 40;

/**
 * Строки «Остатков и оборотов» Хозрасчетного за период по набору счетов (+орг).
 *
 * Безопасность итога: каждая пачка счетов — через fetchAllForAggregation (громкое
 * переполнение), и суммарное число строк по всем пачкам тоже сверяется с тем же
 * потолком analyticsMaxRows. Частичной выборки наружу не уходит — только ошибка.
 */
export async function turnoversByAccounts(
  conn: Connection,
  accountKeys: readonly string[],
  from: string,
  to: string,
  orgKey?: string,
): Promise<{ rows: ODataEntity[]; meta: ScanMeta }> {
  if (accountKeys.length === 0) return { rows: [], meta: emptyMeta() };
  const reg = await requireEntity(conn, accountingRegister(conn.cfg), "Регистр бухгалтерии «Хозрасчётный»");
  const path = balanceAndTurnoversPath(reg, from, to);
  const cap = conn.behavior.analyticsMaxRows;
  const period = `${from}..${to}`;

  let rows: ODataEntity[] = [];
  let meta = emptyMeta();
  for (let i = 0; i < accountKeys.length; i += ACCOUNT_FILTER_BATCH) {
    const batch = accountKeys.slice(i, i + ACCOUNT_FILTER_BATCH);
    // Без $select: 1С группирует виртуальную таблицу по выбранным полям, и
    // $select=Account_Key,… свернул бы сальдо по субконто. Нужны строки по всем
    // измерениям/субконто — как развёрнутое сальдо в стандартной ОСВ.
    const part = await bySplittingAccounts(batch, (keys) =>
      fetchAllForAggregation(conn, path, { filter: accountsOrgFilter(keys, orgKey) }, period),
    );
    rows = rows.concat(part.rows);
    meta = addMeta(meta, part.meta);
    if (rows.length > cap) throw new AggregateOverflowError(path, cap, period);
  }

  // Структура ответа должна совпадать с ожидаемой — до любой агрегации.
  assertTurnoverFields(rows);

  // Защита от молча неверного сальдо: если 1С вернула разбивку по периодичности,
  // начальные/конечные остатки подпериодов сложились бы с задвоением.
  const periods = new Set(
    rows.map((r) => r[BALANCE_AND_TURNOVERS.periodField]).filter((p) => p !== undefined && p !== null),
  );
  if (periods.size > 1) {
    throw new Error(
      `Виртуальная таблица ${BALANCE_AND_TURNOVERS.name} вернула разбивку по периодам ` +
        `(${periods.size} значений ${BALANCE_AND_TURNOVERS.periodField}) — остатки нельзя сложить. ` +
        `Проверьте параметры виртуальной таблицы (periodicity) в src/odata/accounting.ts.`,
    );
  }
  return { rows, meta };
}

/**
 * Резолвит GUID-ы в наименования (Description) из справочника батчами.
 * Используется, чтобы показать имена контрагентов/номенклатуры вместо GUID.
 */
export async function resolveNames(
  conn: Connection,
  entitySet: string,
  keys: Iterable<string>,
): Promise<Map<string, string>> {
  const rows = await readByRefs(conn, entitySet, keys, ["Ref_Key", "Description"]);
  const result = new Map<string, string>();
  for (const [ref, r] of rows) result.set(ref, String(r["Description"] ?? ""));
  return result;
}

/**
 * Читает объекты набора по списку Ref_Key батчами (только выбранные поля). Набор не
 * опубликован — пустая карта: вызывающий сам решает, как показать «нет данных».
 */
export async function readByRefs(
  conn: Connection,
  entitySet: string,
  keys: Iterable<string>,
  select: string[],
): Promise<Map<string, ODataEntity>> {
  const unique = [...new Set([...keys].filter(Boolean))];
  const result = new Map<string, ODataEntity>();
  const available = await conn.available();
  if (!available.has(entitySet)) return result;

  const BATCH = 20;
  for (let i = 0; i < unique.length; i += BATCH) {
    const batch = unique.slice(i, i + BATCH);
    const filter = or(...batch.map((k) => cmp("Ref_Key", "eq", odataGuid(k))));
    const { rows } = await fetchAll(conn.client, entitySet, { filter, select }, BATCH, BATCH);
    for (const r of rows) result.set(String(r["Ref_Key"]), r);
  }
  return result;
}

/** Число из поля сальдо (OData может отдавать строкой). */
export function num(v: unknown): number {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}

export { and, cmp };

// ─── Проводки одного документа (регистратора) ─────────────────────────────────

/** GUID без фигурных скобок (Ref_Key документа). */
export const GUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;

/**
 * ЗАПИСИ регистра бухгалтерии (не виртуальные таблицы) — то, как их публикует 1С.
 *
 * Проверено на двух живых базах БП 3.0 (23.10.2026): записи опубликованы как EntitySet
 * «AccountingRegister_Хозрасчетный_RecordType» с полями Recorder / Recorder_Type / Period /
 * AccountDr_Key / AccountCr_Key / Сумма. Субконто (ExtDimension…) в нём НЕТ — их отдаёт
 * виртуальная таблица RecordsWithExtDimensions (см. fetchRegistrarPostings). Другие конфигурации
 * могут публиковать иначе — поэтому структура по-прежнему сверяется с $metadata конкретной базы.
 *
 * Защита от неверного предположения: перед запросом fetchRegistrarPostings сверяет
 * имена полей с $metadata конкретной базы (resolvePostingsSource) и при расхождении
 * падает с перечнем фактических ИМЁН полей — без запроса к регистру и без какого-либо
 * запасного пути с выкачкой всего регистра. Всё, что может понадобиться поправить
 * после живого теста, собрано здесь и в recorderFilter() — больше нигде.
 */
export const REGISTER_RECORDS = {
  /**
   * Кандидаты EntitySet с записями регистра — по порядку. Берётся первый,
   * опубликованный в $metadata И содержащий все обязательные поля + поле регистратора.
   * «<Регистр>_RecordType» — запись регистра по правилам именования стандартного
   * интерфейса OData; сам «<Регистр>» — набор записей (Recorder/Recorder_Type/RecordSet),
   * он пройдёт проверку, только если база публикует поля записи плоско.
   */
  entitySetSuffixes: ["_RecordType", ""],
  /** Поле регистратора (составной тип: значение GUID + отдельное поле _Type). */
  recorder: "Recorder",
  recorderType: "Recorder_Type",
  /** Без этих полей строка не является проводкой — их отсутствие = громкая ошибка. */
  required: {
    period: "Period",
    accountDr: "AccountDr_Key",
    accountCr: "AccountCr_Key",
    amount: "Сумма",
  },
  /** Используются, если есть в $metadata регистра. */
  optional: {
    lineNumber: "LineNumber",
    active: "Active",
    organization: "Организация_Key",
    quantityDr: "КоличествоDr",
    quantityCr: "КоличествоCr",
    currencyAmountDr: "ВалютнаяСуммаDr",
    currencyAmountCr: "ВалютнаяСуммаCr",
    divisionDr: "ПодразделениеDr_Key",
    divisionCr: "ПодразделениеCr_Key",
    content: "Содержание",
  },
  /** Субконто записи: ExtDimensionDr1..3 / ExtDimensionCr1..3 (+ поле <имя>_Type). */
  extDimensionRe: /^ExtDimension(Dr|Cr)(\d+)$/,
} as const;

/**
 * ЕДИНСТВЕННОЕ место с синтаксисом отбора по регистратору.
 *
 * Синтаксис
 *     Recorder eq cast(guid'…', 'Document_…')
 * проверен на двух живых базах БП 3.0 (23.10.2026): и на записях регистра (_RecordType), и на
 * RecordsWithExtDimensions 1С возвращает ровно записи этого документа.
 *
 * Если живая 1С отвергнет фильтр (400) или вернёт не те записи — инструмент обязан
 * упасть явно (ODataError / проверка assertPostingRows). Никакого альтернативного
 * синтаксиса и никакого отката к выборке всего регистра здесь нет намеренно:
 * реальное представление определяем по живой базе и правим только эту функцию.
 */
export function recorderFilter(documentEntity: string, ref: string): string {
  return cmp(REGISTER_RECORDS.recorder, "eq", `cast(${odataGuid(ref)}, ${odataString(documentEntity)})`);
}

/** Субконто одной стороны проводки, найденное в $metadata. */
export interface ExtDimensionField {
  side: "Dr" | "Cr";
  index: number;
  field: string;
  typeField?: string;
}

export interface PostingsSource {
  entitySet: string;
  properties: ReadonlySet<string>;
  extDimensions: ExtDimensionField[];
}

const propNames = (e: EntityMeta | undefined): Set<string> =>
  new Set((e?.properties ?? []).map((p) => p.name));

/**
 * Находит по $metadata EntitySet записей регистра бухгалтерии и проверяет, что в нём
 * есть поле регистратора и все обязательные поля проводки. Ничего не угадывает:
 * не нашлось — ошибка с ИМЕНАМИ полей кандидатов (значений в ней нет).
 */
export async function resolvePostingsSource(conn: Connection): Promise<PostingsSource> {
  const reg = await requireEntity(conn, accountingRegister(conn.cfg), "Регистр бухгалтерии «Хозрасчётный»");
  const meta = await conn.getMetadata();
  const needed = [REGISTER_RECORDS.recorder, ...Object.values(REGISTER_RECORDS.required)];
  const report: string[] = [];
  for (const suffix of REGISTER_RECORDS.entitySetSuffixes) {
    const set = `${reg}${suffix}`;
    const em = meta.entities.get(set);
    if (!em) {
      report.push(`${set}: не опубликован`);
      continue;
    }
    const props = propNames(em);
    const missing = needed.filter((f) => !props.has(f));
    if (missing.length > 0) {
      report.push(`${set}: нет полей ${missing.join(", ")}; есть: ${[...props].join(", ") || "(нет)"}`);
      continue;
    }
    const extDimensions: ExtDimensionField[] = [];
    for (const p of props) {
      const m = REGISTER_RECORDS.extDimensionRe.exec(p);
      if (!m) continue;
      const typeField = `${p}_Type`;
      extDimensions.push({
        side: m[1] as "Dr" | "Cr",
        index: Number(m[2]),
        field: p,
        ...(props.has(typeField) ? { typeField } : {}),
      });
    }
    extDimensions.sort((a, b) => a.index - b.index);
    return { entitySet: set, properties: props, extDimensions };
  }
  throw new Error(
    `В $metadata не найден набор записей регистра бухгалтерии с полем регистратора и полями проводки. ` +
      `Проверено: ${report.join(" | ")}. Проверьте REGISTER_RECORDS в src/odata/accounting.ts.`,
  );
}

/** Поля шапки документа, которые берём, если они есть в $metadata данного вида документа. */
export const DOCUMENT_HEADER_FIELDS = [
  "Ref_Key",
  "Number",
  "Date",
  "Posted",
  "DeletionMark",
  "Организация_Key",
  "ВидОперации",
  "Состояние",
  "Комментарий",
] as const;

/**
 * Проверяет, что documentEntity — именно документ из $metadata (не справочник, не
 * табличная часть документа, не произвольная сущность). Возвращает его описание.
 */
export async function requireDocumentEntity(conn: Connection, documentEntity: string): Promise<EntityMeta> {
  if (!documentEntity.startsWith("Document_")) {
    throw new InputError(
      `"${documentEntity}" — не документ: ожидается имя вида Document_<Имя> (см. list_entities, class=document).`,
    );
  }
  const meta = await conn.getMetadata();
  ensurePublished(new Set(meta.entities.keys()), documentEntity);
  const em = meta.entities.get(documentEntity)!;
  const props = propNames(em);
  // Табличная часть документа (Document_X_Товары) тоже class=document, но у неё
  // есть LineNumber и составной ключ — регистратором она быть не может.
  if (em.class !== "document" || !props.has("Ref_Key") || props.has("LineNumber")) {
    throw new InputError(
      `"${documentEntity}" — не документ (табличная часть или иной объект). Укажите сам документ, напр. Document_РеализацияТоваровУслуг.`,
    );
  }
  return em;
}

/**
 * Читает шапку документа (только GET). Несуществующий Ref_Key — явная ошибка
 * not_found, а не «пустые проводки». fields — какие поля запросить (берутся только
 * те, что есть в $metadata данного вида документа); по умолчанию DOCUMENT_HEADER_FIELDS.
 */
export async function readDocumentHeader(
  conn: Connection,
  em: EntityMeta,
  ref: string,
  fields: readonly string[] = DOCUMENT_HEADER_FIELDS,
): Promise<ODataEntity> {
  const props = propNames(em);
  const select = fields.filter((f) => props.has(f));
  const path = `${em.entitySet}(${odataGuid(ref)})${buildQuery({ select })}`;
  try {
    const doc = await conn.client.getEntity(path);
    if (!doc || typeof doc !== "object") throw notFound(em.entitySet, ref);
    return doc;
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found") throw notFound(em.entitySet, ref);
    throw e;
  }
}

const notFound = (entitySet: string, ref: string): ODataError =>
  new ODataError({
    kind: "not_found",
    message: `Документ ${entitySet} с Ref_Key ${ref} не найден в базе.`,
  });

const normGuid = (v: unknown): string =>
  typeof v === "string" ? v.replace(/[{}]/g, "").trim().toLowerCase() : "";

const stripNs = (t: string): string => t.slice(t.lastIndexOf(".") + 1);

/**
 * Проверяет строки регистра ДО любой агрегации:
 *  - все обязательные поля проводки присутствуют (иначе num() дал бы нули);
 *  - каждая строка принадлежит именно этому регистратору (защита от молча
 *    неприменённого серверного отбора — частичные/чужие проводки хуже ошибки);
 *  - номера строк не повторяются (иначе $skip-листание задвоило записи).
 * В ошибку идут только ИМЕНА полей и номер строки, не значения.
 */
export function assertPostingRows(
  rows: readonly ODataEntity[],
  documentEntity: string,
  ref: string,
  source: PostingsSource,
): void {
  const required = [
    REGISTER_RECORDS.recorder,
    ...Object.values(REGISTER_RECORDS.required),
    ...(source.properties.has(REGISTER_RECORDS.recorderType) ? [REGISTER_RECORDS.recorderType] : []),
  ];
  const want = normGuid(ref);
  const lineField = REGISTER_RECORDS.optional.lineNumber;
  const seenLines = new Set<string>();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] as Record<string, unknown>;
    const missing = required.filter((f) => r[f] === undefined || r[f] === null);
    if (missing.length > 0) {
      throw new Error(
        `Записи ${source.entitySet} несовместимой структуры (строка #${i + 1}): нет полей ${missing.join(", ")}. ` +
          `Фактические поля строки: ${Object.keys(r).join(", ") || "(нет)"}. ` +
          `Проверьте REGISTER_RECORDS в src/odata/accounting.ts.`,
      );
    }
    const recType = r[REGISTER_RECORDS.recorderType];
    const typeOk =
      recType === undefined || (typeof recType === "string" && stripNs(recType) === documentEntity);
    if (normGuid(r[REGISTER_RECORDS.recorder]) !== want || !typeOk) {
      throw new Error(
        `1С вернула запись ${source.entitySet} другого регистратора (строка #${i + 1}): серверный отбор ` +
          `по регистратору не применился как ожидалось. Проверьте recorderFilter() в src/odata/accounting.ts.`,
      );
    }
    if (r[lineField] !== undefined && r[lineField] !== null) {
      const ln = String(r[lineField]);
      if (seenLines.has(ln)) {
        throw new Error(
          `Повтор ${lineField} в записях ${source.entitySet} (строка #${i + 1}): постраничная выборка ` +
            `неустойчива, набор проводок может быть неполным.`,
        );
      }
      seenLines.add(ln);
    }
  }
}

/**
 * Все записи регистра Хозрасчетный по одному регистратору — отбор на стороне 1С.
 * Через fetchAllForAggregation: полная выборка с громким переполнением
 * (AggregateOverflowError) вместо частичного набора. Никакого «плана Б» с
 * выкачкой всего регистра: ошибка серверного отбора всплывает как есть.
 */
export async function fetchRegistrarPostings(
  conn: Connection,
  documentEntity: string,
  ref: string,
): Promise<{
  rows: ODataEntity[];
  meta: ScanMeta;
  source: PostingsSource;
  filter: string;
  /** Откуда взяты субконто: из записей регистра или из «Движений с субконто»; нет — почему. */
  dimensions: { from: "records" | "recordsWithExtDimensions" } | { unavailable: string };
}> {
  const source = await resolvePostingsSource(conn);
  const filter = recorderFilter(documentEntity, ref);
  const lineField = REGISTER_RECORDS.optional.lineNumber;
  // Устойчивый порядок для $skip-листания: у одного регистратора номер строки уникален.
  const orderby = source.properties.has(lineField) ? `${lineField} asc` : undefined;
  // Без $select: состав полей записи зависит от конфигурации (субконто, _Type-поля),
  // а лишнее поле в $select — это 400. Проверка структуры — assertPostingRows.
  const { rows, meta } = await fetchAllForAggregation(
    conn,
    source.entitySet,
    { filter, ...(orderby ? { orderby } : {}) },
    `документ ${documentEntity} ${ref}`,
  );
  assertPostingRows(rows, documentEntity, ref, source);
  if (source.extDimensions.length > 0 || rows.length === 0)
    return { rows, meta, source, filter, dimensions: { from: "records" } };
  // На БП 3.0 записи регистра (_RecordType) субконто не публикуют вовсе (проверено на живых базах):
  // их отдаёт виртуальная таблица «Движения с субконто» — та же, что у карточки счёта (#49).
  try {
    const enriched = await withExtDimensions(conn, documentEntity, ref, rows);
    if ("unavailable" in enriched) return { rows, meta, source, filter, dimensions: enriched };
    return {
      rows: enriched.rows,
      meta: addMeta(meta, enriched.meta),
      source: { ...source, extDimensions: enriched.extDimensions },
      filter,
      dimensions: { from: "recordsWithExtDimensions" },
    };
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    return { rows, meta, source, filter, dimensions: { unavailable: reason } };
  }
}

const RECORDS_WITH_EXT_DIMENSIONS = "RecordsWithExtDimensions";

/**
 * Субконто проводок документа из «Движений с субконто» (RecordsWithExtDimensions): за период его
 * проводок, отбор по тому же регистратору, сопоставление с записями регистра по LineNumber.
 * Не сошлось число строк или номера — субконто не подставляем (unavailable), проводки остаются.
 */
async function withExtDimensions(
  conn: Connection,
  documentEntity: string,
  ref: string,
  rows: readonly ODataEntity[],
): Promise<
  { rows: ODataEntity[]; meta: ScanMeta; extDimensions: ExtDimensionField[] } | { unavailable: string }
> {
  const reg = await requireEntity(conn, accountingRegister(conn.cfg), "Регистр бухгалтерии «Хозрасчётный»");
  const periods = rows
    .map((r) => String(r[REGISTER_RECORDS.required.period] ?? ""))
    .filter(Boolean)
    .sort();
  const from = periods[0]!.slice(0, 19);
  const to = periods[periods.length - 1]!.slice(0, 19);
  const path = `${reg}/${RECORDS_WITH_EXT_DIMENSIONS}(StartPeriod=datetime'${from}',EndPeriod=datetime'${to}')`;
  const { rows: dimRows, meta } = await fetchAllForAggregation(
    conn,
    path,
    { filter: recorderFilter(documentEntity, ref) },
    `субконто документа ${documentEntity} ${ref}`,
  );
  const line = REGISTER_RECORDS.optional.lineNumber;
  const byLine = new Map(dimRows.map((r) => [Number(r[line]), r]));
  if (dimRows.length !== rows.length || rows.some((r) => !byLine.has(Number(r[line]))))
    return {
      unavailable: `«Движения с субконто» вернули ${dimRows.length} строк(и) при ${rows.length} записях регистра — сопоставить по номеру строки нельзя`,
    };
  const fields = new Set<string>();
  for (const r of dimRows)
    for (const k of Object.keys(r)) if (REGISTER_RECORDS.extDimensionRe.test(k)) fields.add(k);
  const extDimensions: ExtDimensionField[] = [...fields]
    .map((field) => {
      const m = REGISTER_RECORDS.extDimensionRe.exec(field)!;
      return { side: m[1] as "Dr" | "Cr", index: Number(m[2]), field, typeField: `${field}_Type` };
    })
    .sort((a, b) => a.index - b.index);
  const merged = rows.map((r) => {
    const d = byLine.get(Number(r[line]))!;
    const extra: ODataEntity = {};
    for (const x of extDimensions) {
      // Вида субконто нет (ExtDimensionType…_Key = null) — у счёта такого субконто нет по плану
      // счетов, 1С отдаёт null и в значении. Это подтверждённая пустота, а не «не вернула»
      // (проверено на живой базе: счёт 84.01 без видов субконто).
      const kindKey = `ExtDimensionType${x.side}${x.index}_Key`;
      const noSuchSlot = d[x.field] == null && kindKey in d && d[kindKey] == null;
      extra[x.field] = noSuchSlot ? "" : d[x.field];
      if (x.typeField && x.typeField in d) extra[x.typeField] = d[x.typeField];
    }
    return { ...r, ...extra };
  });
  return { rows: merged, meta, extDimensions };
}

/**
 * Счета плана «Хозрасчётный» по Ref_Key — пачками (один запрос на ACCOUNT_FILTER_BATCH
 * счетов), а не запросом на каждую проводку. Несопоставленный счёт — ошибка.
 */
export async function accountsByRef(conn: Connection, keys: Iterable<string>): Promise<Map<string, Account>> {
  const unique = [...new Set([...keys].filter((k) => k && k !== EMPTY_GUID))];
  if (unique.length === 0) return new Map();
  // Ссылки пришли из данных 1С — промах значит устаревший кэш: перечитываем.
  const result = await fromChart(
    conn,
    (chart) => {
      const byKey = new Map(chart.map((a) => [normGuid(a.key), a]));
      return new Map(
        unique.filter((k) => byKey.has(normGuid(k))).map((k) => [normGuid(k), byKey.get(normGuid(k))!]),
      );
    },
    (r) => r.size < unique.length,
  );
  const missing = unique.filter((k) => !result.has(normGuid(k)));
  if (missing.length > 0) {
    throw new Error(`Счета не найдены в плане счетов «Хозрасчётный» по Ref_Key: ${missing.join(", ")}.`);
  }
  return result;
}

/** Ключ для поиска в карте accountsByRef (регистр GUID и скобки не важны). */
export const accountLookupKey = normGuid;

export { EMPTY_GUID };
