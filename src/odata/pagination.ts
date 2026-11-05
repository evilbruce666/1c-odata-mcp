import type { ODataClient } from "./client.js";
import { buildQuery, type QueryOptions } from "./query.js";
import type { ODataEntity } from "../types/odata.js";

/** Таблица со ссылкой: Ref_Key делает порядок однозначным. Виртуальные таблицы и регистры его не имеют. */
const REF_TABLE = /^(Catalog|Document|ChartOfAccounts|ChartOfCharacteristicTypes)_[^/(]+$/;

/**
 * Порядок, при котором листание не теряет строк. Без однозначной сортировки 1С между страницами
 * отдаёт одни строки дважды и пропускает другие, а их число сходится: на живой базе страницы по 100
 * потеряли 255 строк плана счетов из 490 и 47 контрагентов из 148. Сортировка только по дате или
 * счёту тоже не спасает.
 */
export function stableOrderBy(orderby?: string): string {
  if (!orderby) return "Ref_Key asc";
  return /\bRef_Key\b/.test(orderby) ? orderby : `${orderby},Ref_Key asc`;
}

/**
 * Читает строки до конца данных или защитного лимита maxRows. 1С не делает серверной автопагинации.
 * Таблицы со ссылкой листаются через $top/$skip с однозначной сортировкой. Регистры и их виртуальные
 * таблицы ключа не имеют (сортировка по всем измерениям на 1111 строках заняла минуты), а $skip теряет
 * и на страницах по 1000, поэтому они читаются одним запросом.
 *
 * @param entitySet  имя набора, напр. "Catalog_Контрагенты", или путь виртуальной таблицы
 * @param opts       опции запроса
 * @param pageSize   размер страницы для таблиц со ссылкой
 * @param maxRows    жёсткий максимум суммарно возвращаемых строк
 */
export async function fetchAll<T extends ODataEntity = ODataEntity>(
  client: ODataClient,
  entitySet: string,
  opts: QueryOptions,
  pageSize: number,
  maxRows: number,
): Promise<{ rows: T[]; truncated: boolean }> {
  if (!REF_TABLE.test(entitySet)) {
    const query = buildQuery({ ...opts, top: maxRows, skip: opts.skip ?? 0 });
    const page = await client.getCollection<T>(`${entitySet}${query}`);
    const rows = page.value ?? [];
    return { rows, truncated: rows.length >= maxRows };
  }

  const orderby = stableOrderBy(opts.orderby);
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
    const query = buildQuery({ ...opts, orderby, top, skip });
    const page = await client.getCollection<T>(`${entitySet}${query}`);
    const batch = page.value ?? [];
    rows.push(...batch);

    // Конец данных: вернулось меньше, чем просили.
    if (batch.length < top) break;
    skip += batch.length;
  }

  return { rows, truncated };
}
