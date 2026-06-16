import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, contains, odataString } from "./query.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { requireEntity } from "./publication.js";

/**
 * Хелперы для аналитики: поднимают «срезы» справочников (контрагенты по типу,
 * статьи ДДС по имени) в Set/Map, чтобы тяжёлые вопросы вроде «сколько заплатили
 * ИП в прошлом году» или «приход по статье ДДС "Аренда"» решались одним проходом
 * по документам с клиентским фильтром.
 */

/** Категории контрагентов: чем фильтровать справочник. */
export type CounterpartyKind = "ИП" | "ЮрЛицо" | "ФизЛицо" | "Нерезидент" | "Госорган";

/** Возвращает Set<Ref_Key> контрагентов выбранной категории. Пустой Set, если их нет. */
export async function counterpartyRefsByKind(conn: Connection, kind: CounterpartyKind): Promise<Set<string>> {
  const set = resolveEntity(CATALOGS.counterparties, await conn.available());
  if (!set) return new Set();
  // Фильтры собираются под особенности справочника контрагентов БП 3.0:
  //  ИП        = ФизЛицо + ИндивидуальныйПредприниматель=true
  //  ЮрЛицо    = ЮридическоеФизическоеЛицо="ЮридическоеЛицо"
  //  ФизЛицо   = ФизическоеЛицо без признака ИП
  //  Нерезидент= заполнено НаименованиеНерезидентаРус
  //  Госорган  = заполнен ВидГосударственногоОргана
  const filter = (() => {
    switch (kind) {
      case "ИП":
        return and(
          cmp("ЮридическоеФизическоеЛицо", "eq", odataString("ФизическоеЛицо")),
          cmp("ИндивидуальныйПредприниматель", "eq", "true"),
          cmp("IsFolder", "eq", "false"),
        );
      case "ЮрЛицо":
        return and(
          cmp("ЮридическоеФизическоеЛицо", "eq", odataString("ЮридическоеЛицо")),
          cmp("IsFolder", "eq", "false"),
        );
      case "ФизЛицо":
        return and(
          cmp("ЮридическоеФизическоеЛицо", "eq", odataString("ФизическоеЛицо")),
          cmp("ИндивидуальныйПредприниматель", "eq", "false"),
          cmp("IsFolder", "eq", "false"),
        );
      case "Нерезидент":
        return and(cmp("НаименованиеНерезидентаРус", "ne", odataString("")), cmp("IsFolder", "eq", "false"));
      case "Госорган":
        return and(cmp("ВидГосударственногоОргана", "ne", odataString("")), cmp("IsFolder", "eq", "false"));
    }
  })();
  const { rows } = await fetchAll(
    conn.client,
    set,
    { filter, select: ["Ref_Key"] },
    conn.behavior.pageSize,
    conn.behavior.maxRows,
  );
  return new Set(rows.map((r) => String(r["Ref_Key"]).toLowerCase()));
}

/**
 * Резолвит статью ДДС: по Ref_Key (GUID), коду (Code) или подстроке наименования.
 * Возвращает массив подходящих — один точный или несколько по подстроке.
 */
export async function resolveCashflowItems(
  conn: Connection,
  query: string,
): Promise<Array<{ ref: string; code?: string; name?: string }>> {
  const set = await requireEntity(conn, CATALOGS.cashflowItems, "Справочник «Статьи ДДС»");
  const trimmed = query.trim();
  const isGuid = /^\{?[0-9a-fA-F-]{32,38}\}?$/.test(trimmed);
  const filter = isGuid
    ? cmp("Ref_Key", "eq", `guid'${trimmed.replace(/[{}]/g, "")}'`)
    : and(
        cmp("IsFolder", "eq", "false"),
        `(${cmp("Code", "eq", odataString(trimmed))} or ${contains("Description", trimmed)})`,
      );
  const { rows } = await fetchAll(
    conn.client,
    set,
    { filter, select: ["Ref_Key", "Code", "Description"] },
    20,
    20,
  );
  return rows.map((r) => ({
    ref: String(r["Ref_Key"]),
    code: r["Code"] ? String(r["Code"]) : undefined,
    name: r["Description"] ? String(r["Description"]) : undefined,
  }));
}
