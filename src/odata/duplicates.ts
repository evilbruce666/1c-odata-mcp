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
 * Исключение — обособленное подразделение: КПП указан и не совпадает ни с одним найденным непустым КПП (так же
 * сверяет и сама 1С — по паре ИНН+КПП). Повтор подтверждения той же операции сюда не доходит:
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
  const { value: rows } = await conn.client.getCollection<ODataEntity>(
    `${entitySet}${buildQuery({
      filter: and(cmp("ИНН", "eq", odataString(value)), "DeletionMark eq false"),
      select,
      top: 20,
    })}`,
  );
  const self = exclude?.replace(/[{}']/g, "").toLowerCase();
  const hits = rows.filter(
    (r) => String(r["ИНН"] ?? "").trim() === value && String(r["Ref_Key"]).toLowerCase() !== self,
  );
  if (hits.length === 0) return {};
  const kppValue = kpp?.trim();
  const kppOf = (r: ODataEntity) => String(r["КПП"] ?? "").trim();
  const list = hits
    .slice(0, 5)
    .map(
      (r) =>
        `«${String(r["Description"] ?? "")}» (ref ${String(r["Ref_Key"])}${kppOf(r) ? `, КПП ${kppOf(r)}` : ""})`,
    )
    .join("; ");
  if (hasKpp && kppValue && hits.every((r) => kppOf(r) && kppOf(r) !== kppValue))
    return {
      note: `С ИНН ${value} уже есть контрагенты с другим КПП: ${list}. Создаётся обособленное подразделение с КПП ${kppValue}.`,
    };
  return {
    refusal:
      `Контрагент с ИНН ${value} уже есть: ${list}. Ничего не создано. Используйте существующего ` +
      "(read.counterparty.get_counterparty, write.counterparty.update_counterparty)" +
      (hasKpp ? " или, если это обособленное подразделение, укажите его КПП." : "."),
  };
}
