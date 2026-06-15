import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { contains } from "./query.js";
import { CATALOGS } from "../config/mapping.js";
import { requireEntity } from "./publication.js";

export interface Organization {
  ref: string;
  name: string;
  inn?: string;
}

const orgSet = (conn: Connection): Promise<string> =>
  requireEntity(conn, CATALOGS.organizations, "Справочник «Организации»");

/** Все организации (юрлица) базы. */
export async function listOrganizations(conn: Connection): Promise<Organization[]> {
  const set = await orgSet(conn);
  const { rows } = await fetchAll(
    conn.client,
    set,
    { select: ["Ref_Key", "Description", "ИНН"], orderby: "Description" },
    conn.behavior.pageSize,
    conn.behavior.maxRows,
  );
  return rows.map((r) => ({
    ref: String(r["Ref_Key"] ?? ""),
    name: String(r["Description"] ?? ""),
    inn: r["ИНН"] ? String(r["ИНН"]) : undefined,
  }));
}

/**
 * Резолвит организацию по части названия в Ref_Key.
 * Бросает понятную ошибку, если ничего не найдено или найдено неоднозначно.
 */
export async function resolveOrganization(conn: Connection, query: string): Promise<Organization> {
  const set = await orgSet(conn);
  const { rows } = await fetchAll(
    conn.client,
    set,
    { filter: contains("Description", query), select: ["Ref_Key", "Description", "ИНН"] },
    20,
    20,
  );
  if (rows.length === 0) {
    throw new Error(`Организация "${query}" не найдена. Список — в list_organizations.`);
  }
  if (rows.length > 1) {
    const names = rows.map((r) => String(r["Description"])).join(", ");
    throw new Error(`Под "${query}" подходит несколько организаций: ${names}. Уточните название.`);
  }
  const r = rows[0] as Record<string, unknown>;
  return {
    ref: String(r["Ref_Key"] ?? ""),
    name: String(r["Description"] ?? ""),
    inn: r["ИНН"] ? String(r["ИНН"]) : undefined,
  };
}
