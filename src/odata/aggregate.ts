import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp } from "./query.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Единый безопасный путь для АГРЕГАЦИИ документов (суммы/итоги).
 *
 * Принципы (чтобы не повторять баг тихого недосчёта):
 *  1. Все агрегаты идут через эти функции — лимит не разбросан по инструментам.
 *  2. Никакой тихой обрезки: если данных больше потолка `analyticsMaxRows`,
 *     бросаем AggregateOverflowError (инструмент вернёт явную ошибку), а не
 *     неполную сумму под видом ответа.
 *  3. Большой период, упёршийся в потолок, автоматически режется по месяцам и
 *     суммируется ПОЛНОСТЬЮ — это не обрезка, а способ досчитать. Громкая ошибка
 *     только если даже один месяц не влезает (т.е. данных реально слишком много).
 *  4. Возвращаем мета-достоверности (rowsScanned, chunks) — видно, что выборка полная.
 */

/** Размер страницы для аналитики: крупнее обычного — меньше round-trip'ов. */
export const ANALYTICS_PAGE = 1000;

/** Бросается, когда точный итог нельзя посчитать в пределах потолка. */
export class AggregateOverflowError extends Error {
  constructor(
    readonly entitySet: string,
    readonly cap: number,
    readonly period?: string,
  ) {
    super(
      `Слишком много документов для точного итога${period ? ` за ${period}` : ""} ` +
        `(> ${cap}) по ${entitySet}. Сузьте период или добавьте фильтр ` +
        `(или увеличьте ODATA_ANALYTICS_MAX_ROWS).`,
    );
    this.name = "AggregateOverflowError";
  }
}

/** Мета-достоверности выборки: сколько строк просканировано и за сколько окон. */
export interface ScanMeta {
  rowsScanned: number;
  chunks: number;
}

export const emptyMeta = (): ScanMeta => ({ rowsScanned: 0, chunks: 0 });
export const addMeta = (a: ScanMeta, b: ScanMeta): ScanMeta => ({
  rowsScanned: a.rowsScanned + b.rowsScanned,
  chunks: a.chunks + b.chunks,
});

/**
 * Одна выборка под аналитику с ГРОМКИМ переполнением: тянем до cap+1 строки,
 * и если их больше потолка — бросаем (значит выборка была бы неполной).
 */
export async function fetchAllForAggregation(
  conn: Connection,
  entitySet: string,
  opts: { filter?: string | undefined; select?: string[]; orderby?: string },
  period?: string,
): Promise<{ rows: ODataEntity[]; meta: ScanMeta }> {
  const cap = conn.behavior.analyticsMaxRows;
  const { rows } = await fetchAll(conn.client, entitySet, opts, ANALYTICS_PAGE, cap + 1);
  if (rows.length > cap) throw new AggregateOverflowError(entitySet, cap, period);
  return { rows, meta: { rowsScanned: rows.length, chunks: 1 } };
}

/** Делит [from,to] (YYYY-MM-DD) на календарные месяцы, клампя к границам периода. */
export function splitByMonth(from: string, to: string): Array<{ from: string; to: string }> {
  const [fy, fm] = from.split("-").map(Number) as [number, number, number];
  const [ty, tm] = to.split("-").map(Number) as [number, number, number];
  const out: Array<{ from: string; to: string }> = [];
  let y = fy,
    m = fm;
  const pad = (n: number): string => String(n).padStart(2, "0");
  while (y < ty || (y === ty && m <= tm)) {
    const first = y === fy && m === fm ? from : `${y}-${pad(m)}-01`;
    const lastDay = new Date(y, m, 0).getDate(); // последний день месяца m (1-based)
    const last = y === ty && m === tm ? to : `${y}-${pad(m)}-${pad(lastDay)}`;
    out.push({ from: first, to: last });
    m += 1;
    if (m > 12) {
      m = 1;
      y += 1;
    }
  }
  return out;
}

/**
 * Собирает ВСЕ документы за период с авто-чанкингом по месяцам при переполнении.
 * baseFilter — условия БЕЗ даты (орг/проведён/контрагент/статья и т.п.); дату
 * добавляем здесь, чтобы корректно резать по окнам.
 */
export async function collectDocuments(
  conn: Connection,
  entitySet: string,
  params: { baseFilter?: string | undefined; dateField: string; from: string; to: string; select: string[] },
): Promise<{ rows: ODataEntity[]; meta: ScanMeta }> {
  const { baseFilter, dateField, from, to, select } = params;
  const filter =
    and(
      baseFilter,
      cmp(dateField, "ge", `datetime'${from}T00:00:00'`),
      cmp(dateField, "le", `datetime'${to}T23:59:59'`),
    ) || undefined;
  try {
    return await fetchAllForAggregation(
      conn,
      entitySet,
      { filter, select, orderby: `${dateField} asc` },
      `${from}..${to}`,
    );
  } catch (e) {
    if (!(e instanceof AggregateOverflowError)) throw e;
    const months = splitByMonth(from, to);
    if (months.length <= 1) throw e; // даже один месяц не влезает — действительно слишком много
    let rows: ODataEntity[] = [];
    let meta = emptyMeta();
    for (const win of months) {
      const part = await collectDocuments(conn, entitySet, {
        baseFilter,
        dateField,
        from: win.from,
        to: win.to,
        select,
      });
      rows = rows.concat(part.rows);
      meta = addMeta(meta, part.meta);
    }
    return { rows, meta };
  }
}
