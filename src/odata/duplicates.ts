import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { and, buildQuery, cmp, odataString } from "./query.js";

export interface DuplicateCheck {
  /** Текст отказа: такой контрагент уже есть. */
  refusal?: string;
  /** Предупреждение: ИНН совпал, но КПП другой — обособленное подразделение. */
  note?: string;
}

/**
 * Контрагент с тем же ИНН уже есть (без пометки удаления) — отказ со ссылкой (решение владельца, 20.11.2026).
 * Исключение — обособленное подразделение: КПП указан, и контрагента с этой парой ИНН+КПП (или с пустым КПП) нет — так
 * же сверяет и сама 1С. Повтор подтверждения той же операции сюда не доходит:
 * журнал отвечает раньше, чем инструмент читает 1С.
 */
export async function checkDuplicateInn(
  conn: Connection,
  entitySet: string,
  inn: string | undefined,
  kpp: string | undefined,
  /** Изменяемый объект: при правке ИНН он сам себе не дубль. */
  exclude?: string,
): Promise<DuplicateCheck> {
  const value = inn?.trim();
  if (!value) return {};
  const props = new Set(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => p.name),
  );
  if (!props.has("ИНН")) return {};
  const hasKpp = props.has("КПП");
  const select = ["Ref_Key", "Description", "ИНН", ...(hasKpp ? ["КПП"] : [])];
  const self = exclude?.replace(/[{}']/g, "").toLowerCase();
  const find = async (extra?: string) =>
    (
      await conn.client.getCollection<ODataEntity>(
        `${entitySet}${buildQuery({
          filter: and(cmp("ИНН", "eq", odataString(value)), "DeletionMark eq false", extra),
          select,
          top: 5,
        })}`,
      )
    ).value.filter(
      (r) => String(r["ИНН"] ?? "").trim() === value && String(r["Ref_Key"]).toLowerCase() !== self,
    );
  const kppValue = hasKpp ? kpp?.trim() : undefined;
  // С КПП — точная пара (или КПП пуст: подразделение не отличить), а не вывод по первым строкам: у сети филиалов
  // их больше страницы.
  const hits = await find(
    kppValue ? `(${cmp("КПП", "eq", odataString(kppValue))} or ${cmp("КПП", "eq", "''")})` : undefined,
  );
  const kppOf = (r: ODataEntity) => String(r["КПП"] ?? "").trim();
  const describe = (rows: ODataEntity[]) =>
    rows
      .map(
        (r) =>
          `«${String(r["Description"] ?? "")}» (ref ${String(r["Ref_Key"])}${kppOf(r) ? `, КПП ${kppOf(r)}` : ""})`,
      )
      .join("; ");
  const list = describe(hits);
  if (hits.length === 0) {
    if (!kppValue) return {};
    const others = await find();
    return others.length
      ? {
          note: `С ИНН ${value} уже есть контрагенты с другим КПП: ${describe(others)}. Создаётся обособленное подразделение с КПП ${kppValue}.`,
        }
      : {};
  }
  return {
    refusal:
      `Контрагент с ИНН ${value} уже есть: ${list}. Ничего не записано. Используйте существующего ` +
      "(read.counterparty.get_counterparty, write.counterparty.update_counterparty)" +
      (hasKpp ? " или, если это обособленное подразделение, укажите его КПП." : "."),
  };
}
