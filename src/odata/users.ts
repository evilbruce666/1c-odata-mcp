import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";
import { ODataError } from "./errors.js";
import { resolveEntity } from "../config/mapping.js";
import { fetchAll } from "./pagination.js";
import { odataGuid } from "./query.js";
import { resolveNames } from "./accounting.js";
import { requireEntity } from "./publication.js";

/**
 * Пользователи 1С (справочник «Пользователи», одинаковый в Бухгалтерии 3.0 и УТ 11.5): для выбора ответственного в
 * документах записи. Отдаются только имя, признаки и подразделение — без контактной информации, фото, физического лица
 * и идентификаторов пользователя ИБ: задача — выбрать ответственного, а не читать кадровые данные.
 */
export const USERS = "Catalog_Пользователи";
/** Подразделение пользователя: в Бухгалтерии — подразделения организаций, в УТ — структура предприятия. */
const DEPARTMENTS = ["Catalog_СтруктураПредприятия", "Catalog_ПодразделенияОрганизаций"] as const;
const SELECT = ["Ref_Key", "Description", "DeletionMark", "Недействителен", "Служебный", "Подразделение_Key"];
const EMPTY = "00000000-0000-0000-0000-000000000000";
/** Справочник пользователей небольшой (на копии рабочей УТ — около тысячи), читается целиком. */
const MAX_USERS = 20_000;

export interface UserRow {
  ref: string;
  name: string;
  inactive: boolean;
  service: boolean;
  deletionMark: boolean;
  departmentRef?: string;
}

export async function readUsers(conn: Connection): Promise<{ users: UserRow[]; truncated: boolean }> {
  const set = await requireEntity(conn, [USERS], "Справочник «Пользователи»");
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);
  const { rows, truncated } = await fetchAll(
    conn.client,
    set,
    { select: SELECT.filter((f) => props.has(f)) },
    conn.behavior.pageSize,
    MAX_USERS,
  );
  return {
    truncated,
    users: rows.map((r) => {
      const dep = r["Подразделение_Key"];
      return {
        ref: String(r["Ref_Key"] ?? ""),
        name: String(r["Description"] ?? ""),
        inactive: r["Недействителен"] === true,
        service: r["Служебный"] === true,
        deletionMark: r["DeletionMark"] === true,
        ...(typeof dep === "string" && dep && dep !== EMPTY ? { departmentRef: dep } : {}),
      };
    }),
  };
}

/** Названия подразделений; справочник не опубликован — пустая карта (подразделение тогда не показывается). */
export async function departmentNames(
  conn: Connection,
  refs: Iterable<string>,
): Promise<Map<string, string>> {
  const set = resolveEntity(DEPARTMENTS, await conn.available());
  const keys = [...new Set(refs)];
  if (!set || !keys.length) return new Map();
  return resolveNames(conn, set, keys).catch(() => new Map<string, string>());
}

/**
 * Ответственный документа (responsibleRef) до записи: пользователь существует, не помечен на удаление, не
 * недействителен и не служебный — иначе отказ. Раньше ссылка уходила в 1С как есть. Справочник не опубликован —
 * отказ с подсказкой: проверить ссылку нечем, а неверная испортила бы документ молча.
 */
export async function checkResponsible(
  conn: Connection,
  ref: string | undefined,
): Promise<string | undefined> {
  if (ref === undefined) return undefined;
  const guid = ref.trim().replace(/[{}']/g, "");
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(guid))
    throw new InputError(
      "responsibleRef — Ref_Key пользователя 1С (GUID): найдите его через read.catalog.list_users. Ничего не записано.",
    );
  const set = await requireEntity(conn, [USERS], "Справочник «Пользователи»");
  const hint = "Выберите действующего пользователя через read.catalog.list_users. Ничего не записано.";
  let u: ODataEntity;
  try {
    u = await conn.client.getEntity<ODataEntity>(
      `${set}(${odataGuid(guid)})?$format=json&$select=Ref_Key,Description,DeletionMark,Недействителен,Служебный`,
    );
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found")
      throw new InputError(`Пользователь responsibleRef не найден в справочнике «Пользователи». ${hint}`);
    throw e;
  }
  const name = String(u["Description"] ?? "");
  if (u["DeletionMark"] === true)
    throw new InputError(`Пользователь «${name}» помечен на удаление и не может быть ответственным. ${hint}`);
  if (u["Недействителен"] === true)
    throw new InputError(`Пользователь «${name}» недействителен и не может быть ответственным. ${hint}`);
  if (u["Служебный"] === true)
    throw new InputError(`Пользователь «${name}» — служебный и не может быть ответственным. ${hint}`);
  return guid;
}
