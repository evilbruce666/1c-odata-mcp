/**
 * Проверка нового родителя объекта справочника — для общих write.entity.move_to_folder и update_entity с Parent_Key.
 * Перемещать группы по-прежнему можно; отказ — только на неверное перемещение. У части справочников нет IsFolder:
 * там родитель — обычный элемент, и само поле запрашивать нельзя (1С ответит 400).
 */
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "./errors.js";
import { buildQuery } from "./query.js";

const EMPTY = "00000000-0000-0000-0000-000000000000";
const GUID = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
/** Длиннее цепочку родителей не проходим: 1С такие справочники практически не строит, а цикл в данных бесконечен. */
const MAX_DEPTH = 64;

const clean = (v: unknown) => (typeof v === "string" ? v.replace(/[{}']/g, "").trim().toLowerCase() : "");

/** Текст отказа или undefined. parent пустой — перемещение в корень. */
export async function parentMoveRefusal(
  conn: Connection,
  entitySet: string,
  ref: string,
  parent: unknown,
): Promise<string | undefined> {
  const target = clean(parent);
  // Не ссылка и корень — решает 1С, как раньше.
  if (!entitySet.startsWith("Catalog_") || !GUID.test(target) || target === EMPTY) return undefined;
  const props = new Set(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => p.name),
  );
  if (!props.has("Parent_Key")) return undefined;
  const subject = clean(ref);
  if (target === subject) return "Нельзя переместить объект в себя. Ничего не записано.";

  const folders = props.has("IsFolder");
  const select = ["Parent_Key", ...["IsFolder", "DeletionMark"].filter((k) => props.has(k))];
  const read = async (key: string): Promise<ODataEntity | undefined> => {
    try {
      return await conn.client.getEntity<ODataEntity>(`${entitySet}(guid'${key}')${buildQuery({ select })}`);
    } catch (e) {
      if (e instanceof ODataError && e.kind === "not_found") return undefined;
      throw e;
    }
  };

  const destination = await read(target);
  let own: Promise<ODataEntity | undefined> | undefined;
  const current = () => (own ??= read(subject));
  const wrong = !destination
    ? `Новый родитель не найден в ${entitySet}. Ничего не записано.`
    : folders && destination["IsFolder"] === false
      ? "Новый родитель — не группа (папка), а обычный элемент справочника. Ничего не записано."
      : destination["DeletionMark"] === true
        ? "Новый родитель помечен на удаление. Ничего не записано."
        : undefined;
  // Родитель не меняется (update_entity с прежним Parent_Key рядом с другими полями) — это не перемещение.
  if (wrong || !destination) {
    // Не прочиталось — это не повод скрыть отказ: отвечаем причиной, как до проверки «родитель не меняется».
    const unchanged = await current().then(
      (subjectRow) => clean(subjectRow?.["Parent_Key"]) === target,
      () => false,
    );
    return unchanged ? undefined : wrong;
  }

  // Цикл возможен, только если перемещаемый объект сам может быть родителем.
  if (folders && (await current())?.["IsFolder"] === false) return undefined;
  const seen = new Set([target]);
  let row = destination;
  for (let depth = 0; depth < MAX_DEPTH; depth++) {
    const up = clean(row["Parent_Key"]);
    if (!up || up === EMPTY) return undefined;
    if (up === subject)
      return "Нельзя переместить объект в его же подгруппу или потомка: получится цикл. Ничего не записано.";
    if (seen.has(up)) break;
    seen.add(up);
    const next = await read(up);
    if (!next) break;
    row = next;
  }
  return (
    "Не удалось проверить цепочку родителей нового места (в ней цикл, разрыв или больше " +
    `${MAX_DEPTH} уровней). Ничего не записано.`
  );
}
