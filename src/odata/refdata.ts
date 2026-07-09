import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, contains, odataGuid, odataString } from "./query.js";
import { resolveEntity } from "../config/mapping.js";
import { requireEntity } from "./publication.js";

/**
 * Резолверы справочных данных, нужных для записи «богатой» карточки контрагента:
 * виды контактной информации, банк по БИК, доп.реквизиты.
 */

export const KI_KINDS = ["Catalog_ВидыКонтактнойИнформации"] as const;
const BANKS = ["Catalog_Банки"] as const;
const ADD_PROPS = ["ChartOfCharacteristicTypes_ДополнительныеРеквизитыИСведения"] as const;

export interface ContactKind {
  key: string;
  тип: string;
}
export interface ContactKinds {
  phone?: ContactKind;
  email?: ContactKind;
  address?: ContactKind;
}

/** Виды КИ для контрагентов: телефон / email / (юридический) адрес. */
export async function contactKindsForCounterparties(conn: Connection): Promise<ContactKinds> {
  const set = resolveEntity(KI_KINDS, await conn.available());
  if (!set) return {};
  const { rows: folders } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(contains("Description", "Контрагенты"), cmp("IsFolder", "eq", "true")),
      select: ["Ref_Key"],
    },
    1,
    1,
  );
  const folder = folders[0];
  if (!folder) return {};
  const { rows } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(
        cmp("Parent_Key", "eq", odataGuid(String(folder["Ref_Key"]))),
        cmp("IsFolder", "eq", "false"),
      ),
      select: ["Ref_Key", "Description", "Тип"],
    },
    50,
    50,
  );
  const out: ContactKinds = {};
  for (const r of rows) {
    const t = String(r["Тип"]);
    const d = String(r["Description"]);
    const kind = { key: String(r["Ref_Key"]), тип: t };
    if (t === "Телефон" && !out.phone) out.phone = kind;
    else if (t === "АдресЭлектроннойПочты" && !out.email) out.email = kind;
    else if (t === "Адрес" && d.toLowerCase().includes("юридическ")) out.address = kind;
  }
  if (!out.address) {
    const a = rows.find((r) => String(r["Тип"]) === "Адрес");
    if (a) out.address = { key: String(a["Ref_Key"]), тип: "Адрес" };
  }
  return out;
}

/** Банк по БИК (Code в справочнике Банки; берём не-папку). */
export async function resolveBankByBik(
  conn: Connection,
  bik: string,
): Promise<{ ref: string; name: string }> {
  const set = await requireEntity(conn, BANKS, "Справочник «Банки»");
  const { rows } = await fetchAll(
    conn.client,
    set,
    {
      filter: and(cmp("Code", "eq", odataString(bik)), cmp("IsFolder", "eq", "false")),
      select: ["Ref_Key", "Description"],
    },
    3,
    3,
  );
  const first = rows[0];
  if (!first) throw new Error(`Банк с БИК ${bik} не найден в справочнике «Банки».`);
  return { ref: String(first["Ref_Key"]), name: String(first["Description"] ?? "") };
}

/** Ref доп.реквизита по точному наименованию (напр. «ОГРН»), либо undefined если не настроен. */
export async function resolveAdditionalProperty(conn: Connection, name: string): Promise<string | undefined> {
  const set = resolveEntity(ADD_PROPS, await conn.available());
  if (!set) return undefined;
  const { rows } = await fetchAll(
    conn.client,
    set,
    { filter: cmp("Description", "eq", odataString(name)), select: ["Ref_Key"] },
    1,
    1,
  );
  return rows[0] ? String(rows[0]["Ref_Key"]) : undefined;
}
