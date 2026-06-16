import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, odataGuid, odataString, or } from "./query.js";
import { CATALOGS, REGISTERS, resolveEntity } from "../config/mapping.js";
import { requireEntity } from "./publication.js";
import { buildQuery } from "./query.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Аналитика для 1С:Бухгалтерия 3.0 строится на регистре бухгалтерии
 * «Хозрасчетный» и его виртуальной таблице Balance. И дебиторка (сч. 62),
 * и остатки товаров (сч. 41) — это сальдо по соответствующему счёту,
 * с контрагентом/номенклатурой в ExtDimension1.
 */

const CHART_CANDIDATES = ["ChartOfAccounts_Хозрасчетный"] as const;

export interface Account {
  key: string;
  code: string;
  description: string;
}

/** Возвращает счета, код которых начинается с одного из префиксов (напр. "62", "41"). */
export async function resolveAccounts(conn: Connection, prefixes: readonly string[]): Promise<Account[]> {
  const chart = await requireEntity(conn, CHART_CANDIDATES, "План счетов «Хозрасчётный»");
  const filter = or(...prefixes.map((p) => `startswith(Code, ${odataString(p)})`));
  const { rows } = await fetchAll(
    conn.client,
    chart,
    { filter, select: ["Ref_Key", "Code", "Description"], orderby: "Code" },
    conn.behavior.pageSize,
    conn.behavior.maxRows,
  );
  return rows.map((r) => ({
    key: String(r["Ref_Key"] ?? ""),
    code: String(r["Code"] ?? ""),
    description: String(r["Description"] ?? ""),
  }));
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
  const result = new Map<string, string>();
  if (codes.length === 0) return result;
  const chart = await requireEntity(conn, CHART_CANDIDATES, "План счетов «Хозрасчётный»");
  const filter = or(...codes.map((c) => cmp("Code", "eq", odataString(c))));
  const { rows } = await fetchAll(
    conn.client,
    chart,
    { filter, select: ["Ref_Key", "Code"] },
    codes.length,
    codes.length,
  );
  for (const r of rows) result.set(String(r["Code"]), String(r["Ref_Key"]));
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
 * и (необязательно) по организации.
 */
export async function balanceByAccounts(
  conn: Connection,
  accountKeys: string[],
  maxRows: number,
  orgKey?: string,
  asOf?: string,
): Promise<ODataEntity[]> {
  if (accountKeys.length === 0) return [];
  const reg = await requireEntity(conn, REGISTERS.accounting, "Регистр бухгалтерии «Хозрасчётный»");
  const filter = and(
    or(...accountKeys.map((k) => cmp("Account_Key", "eq", odataGuid(k)))),
    orgKey ? cmp("Организация_Key", "eq", odataGuid(orgKey)) : undefined,
  );
  // Параметр Period — НЕ через $filter, а path-параметром у виртуальной таблицы:
  // .../AccountingRegister_Хозрасчетный/Balance(Period=datetime'YYYY-MM-DDT23:59:59').
  // Без него виртуальная таблица возвращает текущее сальдо.
  const balancePath = asOf ? `${reg}/Balance(Period=datetime'${asOf}T23:59:59')` : `${reg}/Balance`;
  const { rows } = await fetchAll(conn.client, balancePath, { filter }, conn.behavior.pageSize, maxRows);
  return rows;
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
  const unique = [...new Set([...keys].filter(Boolean))];
  const result = new Map<string, string>();
  const available = await conn.available();
  if (!available.has(entitySet)) return result;

  const BATCH = 20;
  for (let i = 0; i < unique.length; i += BATCH) {
    const batch = unique.slice(i, i + BATCH);
    const filter = or(...batch.map((k) => cmp("Ref_Key", "eq", odataGuid(k))));
    const { rows } = await fetchAll(
      conn.client,
      entitySet,
      { filter, select: ["Ref_Key", "Description"] },
      BATCH,
      BATCH,
    );
    for (const r of rows) result.set(String(r["Ref_Key"]), String(r["Description"] ?? ""));
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
