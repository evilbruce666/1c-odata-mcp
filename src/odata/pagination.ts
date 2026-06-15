import type { ODataClient } from "./client.js";
import { buildQuery, type QueryOptions } from "./query.js";
import type { ODataEntity } from "../types/odata.js";

/**
 * Тянет страницы через $top/$skip, пока не достигнут конец данных
 * или защитный лимит maxRows. 1С не делает серверной автопагинации,
 * поэтому листаем сами.
 *
 * @param entitySet  имя набора, напр. "Catalog_Контрагенты"
 * @param opts       опции запроса (top трактуется как размер страницы)
 * @param maxRows    жёсткий максимум суммарно возвращаемых строк
 */
export async function fetchAll<T extends ODataEntity = ODataEntity>(
  client: ODataClient,
  entitySet: string,
  opts: QueryOptions,
  pageSize: number,
  maxRows: number,
): Promise<{ rows: T[]; truncated: boolean }> {
  const rows: T[] = [];
  let skip = opts.skip ?? 0;
  let truncated = false;

  for (;;) {
    const remaining = maxRows - rows.length;
    if (remaining <= 0) {
      truncated = true;
      break;
    }
    const top = Math.min(pageSize, remaining);
    const query = buildQuery({ ...opts, top, skip });
    const page = await client.getCollection<T>(`${entitySet}${query}`);
    const batch = page.value ?? [];
    rows.push(...batch);

    // Конец данных: вернулось меньше, чем просили.
    if (batch.length < top) break;
    skip += batch.length;
  }

  return { rows, truncated };
}
