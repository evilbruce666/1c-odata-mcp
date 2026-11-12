import { InputError } from "../errors.js";
import type { Connection } from "../context.js";
import { fetchAll } from "./pagination.js";
import { and, cmp, contains, odataString } from "./query.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import { requireEntity } from "./publication.js";
import { fetchAllForAggregation } from "./aggregate.js";

/**
 * Хелперы для аналитики: поднимают «срезы» справочников (контрагенты по типу,
 * статьи ДДС по имени) в Set/Map, чтобы тяжёлые вопросы вроде «сколько заплатили
 * ИП в прошлом году» или «приход по статье ДДС "Аренда"» решались одним проходом
 * по документам с клиентским фильтром.
 */

/** Категории контрагентов: чем фильтровать справочник. */
export type CounterpartyKind = "ИП" | "ЮрЛицо" | "ФизЛицо" | "Нерезидент" | "Госорган";

/** Возвращает Set<Ref_Key> контрагентов выбранной категории. Пустой Set, если их нет. */
/** Статей ДДС в одном отборе. */
const MATCHES_LIMIT = 20;

/** Отбор категории в справочнике контрагентов БП 3.0 (см. комментарий в counterpartyRefsByKind). */
function bpKindFilter(kind: CounterpartyKind): string {
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
}

/**
 * Отбор категории в справочнике контрагентов Управления торговлей 11: тип лица — одно поле ЮрФизЛицо (ЮрЛицо,
 * ИндивидуальныйПредприниматель, ФизЛицо, ЮрЛицоНеРезидент — значения проверены на базе УТ 11.5), групп нет, признака
 * госоргана нет.
 */
function utKindFilter(kind: CounterpartyKind, props: ReadonlySet<string>): string {
  if (!props.has("ЮрФизЛицо")) {
    throw new InputError(
      "В справочнике контрагентов этой базы нет признака типа лица: отбор по категории невозможен.",
    );
  }
  switch (kind) {
    case "ИП":
      return cmp("ЮрФизЛицо", "eq", odataString("ИндивидуальныйПредприниматель"));
    case "ЮрЛицо":
      // Как в БП: юрлица вместе с нерезидентами.
      return `(${cmp("ЮрФизЛицо", "eq", odataString("ЮрЛицо"))} or ${cmp("ЮрФизЛицо", "eq", odataString("ЮрЛицоНеРезидент"))})`;
    case "ФизЛицо":
      return cmp("ЮрФизЛицо", "eq", odataString("ФизЛицо"));
    case "Нерезидент":
      return cmp("ЮрФизЛицо", "eq", odataString("ЮрЛицоНеРезидент"));
    case "Госорган":
      throw new InputError("В этой базе у контрагентов нет признака госоргана: отбор «Госорган» невозможен.");
  }
}

export async function counterpartyRefsByKind(conn: Connection, kind: CounterpartyKind): Promise<Set<string>> {
  const set = resolveEntity(CATALOGS.counterparties, await conn.available());
  if (!set) return new Set();
  // Фильтры собираются под особенности справочника контрагентов БП 3.0:
  //  ИП        = ФизЛицо + ИндивидуальныйПредприниматель=true
  //  ЮрЛицо    = ЮридическоеФизическоеЛицо="ЮридическоеЛицо"
  //  ФизЛицо   = ФизическоеЛицо без признака ИП
  //  Нерезидент= заполнено НаименованиеНерезидентаРус
  //  Госорган  = заполнен ВидГосударственногоОргана
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));
  const filter = props.has("ИндивидуальныйПредприниматель") ? bpKindFilter(kind) : utKindFilter(kind, props);
  // Срез справочника должен быть ПОЛНЫМ (иначе часть контрагентов «не узнается»
  // при клиентском фильтре) — через безопасную выборку с громким переполнением.
  const { rows } = await fetchAllForAggregation(
    conn,
    set,
    { filter, select: ["Ref_Key"] },
    `справочник ${kind}`,
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
    MATCHES_LIMIT + 1,
    MATCHES_LIMIT + 1,
  );
  // Часть статей не вошла бы в отбор, и сумма оказалась бы заниженной без предупреждения.
  if (rows.length > MATCHES_LIMIT) {
    const variants = rows
      .slice(0, 10)
      .map(
        (r) =>
          `${String(r["Description"] ?? "")} (${r["Code"] ? `код ${String(r["Code"]).trim()}, ` : ""}${String(r["Ref_Key"])})`,
      )
      .join("; ");
    throw new InputError(
      `Под «${query}» подходит больше ${MATCHES_LIMIT} статей ДДС, например: ${variants}. Укажите Ref_Key нужной статьи или её код.`,
    );
  }
  return rows.map((r) => ({
    ref: String(r["Ref_Key"]),
    code: r["Code"] ? String(r["Code"]) : undefined,
    name: r["Description"] ? String(r["Description"]) : undefined,
  }));
}
