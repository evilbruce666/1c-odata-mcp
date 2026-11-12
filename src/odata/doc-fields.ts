import { cmp, odataGuid } from "./query.js";

/**
 * Поля документов, которые в Бухгалтерии и Управлении торговлей называются по-разному при одном смысле. Инструмент
 * выбирает поле по схеме базы: запрос несуществующего поля 1С отвергает целиком, а пропуск поля давал тихий ноль.
 */

/** Вид операции документа: ВидОперации в БП, ХозяйственнаяОперация в УТ. */
export const operationField = (props: ReadonlySet<string>): string | undefined =>
  ["ВидОперации", "ХозяйственнаяОперация"].find((f) => props.has(f));

/**
 * Значение вида операции «уплата налога» для поля: в БП ПеречислениеНалога, в УТ ПеречислениеВБюджет (проверено на
 * базе УТ 11.5: 270 списаний за 2026 год).
 */
export const TAX_PAYMENT_OPERATION: Readonly<Record<string, string>> = {
  ВидОперации: "ПеречислениеНалога",
  ХозяйственнаяОперация: "ПеречислениеВБюджет",
};

export interface ContractField {
  /** Поле для $select и чтения значения. */
  field: string;
  /** Условие «документ по этому договору». */
  filter: (ref: string) => string;
}

/**
 * Договор документа: ДоговорКонтрагента_Key в БП, Договор_Key в документах продаж и закупок УТ, составное Договор
 * (с Договор_Type) в банковских и кассовых документах УТ — его 1С сравнивает только через cast.
 */
export function contractField(
  props: ReadonlySet<string>,
  contractSet: string | undefined,
): ContractField | undefined {
  for (const field of ["ДоговорКонтрагента_Key", "Договор_Key"]) {
    if (props.has(field)) return { field, filter: (ref) => cmp(field, "eq", odataGuid(ref)) };
  }
  if (props.has("Договор") && props.has("Договор_Type") && contractSet) {
    const type = contractSet.replace(/'/g, "");
    return { field: "Договор", filter: (ref) => cmp("Договор", "eq", `cast(${odataGuid(ref)}, '${type}')`) };
  }
  return undefined;
}
