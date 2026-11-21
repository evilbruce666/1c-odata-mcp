import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";
import { accountingChart } from "../config/accounting.js";
import { accountsByCode } from "./accounting.js";
import { buildQuery, odataGuid } from "./query.js";

/**
 * Прочие операции денежных документов Бухгалтерии: корреспондирующий счёт — в шапке
 * (СчетУчетаРасчетовСКонтрагентом), субконто — по видам субконто этого счёта. Без счёта 1С документ не проводит
 * («Не заполнен счёт кредита/дебета»), а угадать его нельзя: на рабочих базах БП 3.0 у таких документов 76.03,
 * 91.02, 66.01, 55.03, 57.03 — счёт задаёт пользователь (corrAccount), значения субконто — из аргументов.
 */
export const OTHER_OPERATIONS: Readonly<Record<string, "Дт" | "Кт">> = {
  ПрочийПриход: "Кт",
  ПрочийРасход: "Дт",
  ПрочееСписание: "Дт",
};

const EMPTY = "00000000-0000-0000-0000-000000000000";

/** Вид субконто (предопределённое имя) → значение из аргументов и тип для поля _Type. */
const KIND_SOURCES: Readonly<Record<string, { arg: keyof OtherArgs; type: string; label: string }>> = {
  Контрагенты: { arg: "counterpartyRef", type: "Catalog_Контрагенты", label: "контрагент (counterpartyRef)" },
  Договоры: { arg: "contractRef", type: "Catalog_ДоговорыКонтрагентов", label: "договор (contractRef)" },
  ПрочиеДоходыИРасходы: {
    arg: "otherItemRef",
    type: "Catalog_ПрочиеДоходыИРасходы",
    label: "статья прочих доходов и расходов (otherItemRef)",
  },
  СтатьиДвиженияДенежныхСредств: {
    arg: "cashflowItemRef",
    type: "Catalog_СтатьиДвиженияДенежныхСредств",
    label: "статья движения денежных средств (cashflowItemRef)",
  },
};

export interface OtherArgs {
  corrAccount?: string | undefined;
  counterpartyRef?: string | undefined;
  contractRef?: string | undefined;
  otherItemRef?: string | undefined;
  cashflowItemRef?: string | undefined;
}

/**
 * Поля шапки для прочей операции: счёт и субконто 1–3. undefined — операция не прочая (и corrAccount не задан).
 * Отказ: прочая операция без счёта, счёт не найден, у субконто нет значения в аргументах.
 */
export async function otherOperationFields(
  conn: Connection,
  operationKind: string,
  a: OtherArgs,
): Promise<{ fields: Record<string, unknown>; notes: string[] } | undefined> {
  const side = OTHER_OPERATIONS[operationKind];
  if (!side) {
    if (a.corrAccount)
      throw new InputError(
        `corrAccount — только для прочих операций (${Object.keys(OTHER_OPERATIONS).join(", ")}); у «${operationKind}» счёт расчётов подбирается по договору. Ничего не записано.`,
      );
    return undefined;
  }
  if (!a.corrAccount)
    throw new InputError(
      `Для операции «${operationKind}» укажите корреспондирующий счёт corrAccount (код, напр. «91.02», «76.09») — без него 1С документ не проведёт. Ничего не записано.`,
    );
  const code = a.corrAccount.trim();
  const ref = (await accountsByCode(conn, [code])).get(code);
  if (!ref) throw new InputError(`Счёт «${code}» не найден в плане счетов. Ничего не записано.`);
  const chart = accountingChart(conn.cfg)[0]!;
  const acc = await conn.client.getEntity<ODataEntity>(
    `${chart}(${odataGuid(ref)})${buildQuery({ select: ["Ref_Key", "ЗапретитьИспользоватьВПроводках", "ExtDimensionTypes"] })}`,
  );
  if (acc["ЗапретитьИспользоватьВПроводках"] === true)
    throw new InputError(
      `Счёт «${code}» — группа: в проводках используются его субсчета. Укажите субсчёт. Ничего не записано.`,
    );
  const kinds = ((acc["ExtDimensionTypes"] as ODataEntity[] | undefined) ?? [])
    .slice()
    .sort((x, y) => Number(x["LineNumber"]) - Number(y["LineNumber"]));
  const fields: Record<string, unknown> = { СчетУчетаРасчетовСКонтрагентом_Key: ref };
  const notes: string[] = [];
  const missing: string[] = [];
  for (const [i, k] of kinds.slice(0, 3).entries()) {
    const kindRef = String(k["ExtDimensionType_Key"] ?? "");
    const kind = await conn.client
      .getEntity<ODataEntity>(
        `ChartOfCharacteristicTypes_ВидыСубконтоХозрасчетные(${odataGuid(kindRef)})${buildQuery({ select: ["PredefinedDataName", "Description"] })}`,
      )
      .catch(() => undefined);
    const name = String(kind?.["PredefinedDataName"] ?? "");
    const src = KIND_SOURCES[name];
    const field = `Субконто${side}${i + 1}`;
    if (!src) {
      notes.push(
        `Субконто ${i + 1} счёта ${code} («${String(kind?.["Description"] ?? name) || "неизвестный вид"}») коннектор не заполняет — при необходимости заполните в 1С.`,
      );
      continue;
    }
    const value = a[src.arg];
    if (!value || value === EMPTY) {
      missing.push(src.label);
      continue;
    }
    fields[field] = value;
    fields[`${field}_Type`] = `StandardODATA.${src.type}`;
  }
  if (missing.length)
    throw new InputError(
      `Счёт ${code} ведётся по субконто: укажите ${missing.join(", ")}. Ничего не записано.`,
    );
  return { fields, notes };
}
