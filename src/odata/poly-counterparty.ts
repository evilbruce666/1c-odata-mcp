import type { Connection } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { ODataError } from "./errors.js";
import { buildQuery, cmp, odataGuid } from "./query.js";

/**
 * Полиморфный «Контрагент» банковских и кассовых документов БП (рядом «Контрагент_Type»): в нём бывает не только
 * контрагент, но и физлицо или сотрудник. Голый guid 1С с таким полем не сравнивает (500 «Нельзя сравнивать поля
 * неограниченной длины…»), поэтому раньше отбор шёл на стороне коннектора — по всем документам периода.
 *
 * Теперь отбор на сервере через cast к справочнику контрагентов. Он находит только контрагентов, поэтому ошибка
 * дешёвая в обе стороны:
 * - 1С отвергла отбор — этот вид документа перечитывается прежним способом;
 * - отбор нигде ничего не нашёл — одна проверка, контрагент ли это. Нет (физлицо, сотрудник, чужой guid) — прежний
 *   отбор на стороне коннектора, иначе такие платежи молча пропали бы.
 * Идея серверного отбора через cast — из форка 1c-odata-mcp-kz (MIT), код написан заново.
 */
export const POLY_COUNTERPARTY = "Контрагент";

/** Как ищется контрагент в полиморфном поле: на сервере (cast) или на стороне коннектора, как раньше. */
export type PolyMode = "server" | "client";

export const POLY_REJECTED_NOTE =
  "1С не приняла серверный отбор по полю «Контрагент» — для этого вида документов отбор на стороне коннектора.";
export const POLY_NOT_COUNTERPARTY_NOTE =
  "Ссылка не найдена среди контрагентов: в документах, где «Контрагент» может быть физлицом или сотрудником, отбор " +
  "на стороне коннектора.";

function counterpartySet(available: ReadonlySet<string>): string {
  return resolveEntity(CATALOGS.counterparties, available) ?? CATALOGS.counterparties[0];
}

/** Условие «в полиморфном поле этот контрагент». */
export function polyCounterpartyFilter(available: ReadonlySet<string>, ref: string): string {
  const set = counterpartySet(available).replace(/'/g, "");
  return cmp(POLY_COUNTERPARTY, "eq", `cast(${odataGuid(ref)}, '${set}')`);
}

/**
 * 1С отвергла сам запрос: 400 или 500, на которые повтор не поможет (в том числе «Нельзя сравнивать поля…»).
 * Сбой связи и перегрузка сюда не относятся — их, как раньше, видит пользователь.
 */
export function rejectedByOneC(e: unknown): boolean {
  return e instanceof ODataError && (e.kind === "bad_request" || (e.kind === "server" && !e.retryable));
}

/**
 * Ссылка — контрагент из справочника. Любая неудача проверки (нет объекта, не опубликован справочник, сбой) — «нет»:
 * тогда отбор идёт прежним способом, ответ верный, только медленнее.
 */
export async function isCounterparty(conn: Connection, ref: string): Promise<boolean> {
  const set = counterpartySet(await conn.available());
  try {
    await conn.client.getEntity(`${set}(${odataGuid(ref)})${buildQuery({ select: ["Ref_Key"] })}`);
    return true;
  } catch {
    return false;
  }
}
