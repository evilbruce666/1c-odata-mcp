import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";
import { accountingChart } from "../config/accounting.js";
import { accountsByCode } from "./accounting.js";
import { buildQuery, odataGuid } from "./query.js";

/**
 * Операции денежных документов Бухгалтерии с корреспондирующим счётом в шапке (СчетУчетаРасчетовСКонтрагентом) и
 * субконто по видам субконто этого счёта. Угадать счёт нельзя — на рабочих базах БП 3.0 у таких документов 76.03,
 * 91.02, 66.01, 55.03, 57.01, 68.90, 69.06 и другие: его задаёт пользователь (corrAccount), субконто — из аргументов.
 * required — без счёта 1С документ не проводит («Не заполнен счёт дебета/кредита»: прочие операции, налог);
 * проверено проведением на рабочей БП 3.0 21.11.2026. Комиссию банка (91.02) и депозит (55.03) 1С проводит со своим
 * счётом — corrAccount у них можно задать, чтобы указать счёт и статью.
 */
/**
 * Виды списания с расчётного счёта, которые коннектор не может собрать верно: на рабочей БП 3.0 (21.11.2026) 1С
 * проводила их без единой проводки или с проводкой Дт 51 Кт 51 — и со счётом в шапке тоже. Отказ с причиной, а не
 * документ, который проведётся неверно.
 */
export const WRITEOFF_REFUSED: Readonly<Record<string, string>> = {
  ПеречислениеЗП:
    "Перечисление зарплаты 1С проводит по ведомости на выплату зарплаты: без неё документ проведётся без проводок. Создайте списание по ведомости в 1С.",
  СнятиеНаличных:
    "Снятие наличных 1С проводит без проводок, если документ записан через OData (проверено и со счётом 57.01). Оформите его в 1С.",
  ПереводНаДругойСчет:
    "Перевод на другой счёт требует счёта-получателя: без него 1С ставит в документ счёт 51 и проводит Дт 51 Кт 51. Оформите перевод в 1С.",
};

export const OTHER_OPERATIONS: Readonly<Record<string, { side: "Дт" | "Кт"; required: boolean }>> = {
  ПрочийПриход: { side: "Кт", required: true },
  ПрочийРасход: { side: "Дт", required: true },
  ПрочееСписание: { side: "Дт", required: true },
  ПеречислениеНалога: { side: "Дт", required: true },
  КомиссияБанка: { side: "Дт", required: false },
  Депозит: { side: "Дт", required: false },
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
  /** Что документ и так пишет в шапку (статья ДДС — всегда; у списания — ещё контрагент и договор). */
  inHeader: ReadonlyArray<keyof OtherArgs> = ["cashflowItemRef"],
): Promise<{ fields: Record<string, unknown>; notes: string[] } | undefined> {
  const spec = OTHER_OPERATIONS[operationKind];
  if (!spec) {
    if (a.corrAccount || a.otherItemRef)
      throw new InputError(
        `corrAccount и otherItemRef — только для операций ${Object.keys(OTHER_OPERATIONS).join(", ")}; у «${operationKind}» счёт 1С подбирает сама (по договору или по виду операции). Ничего не записано.`,
      );
    return undefined;
  }
  const side = spec.side;
  if (!a.corrAccount) {
    if (!spec.required && !a.otherItemRef) return undefined;
    throw new InputError(
      spec.required
        ? `Для операции «${operationKind}» укажите корреспондирующий счёт corrAccount (код, напр. «91.02», «68.90», «57.01») — без него 1С документ не проведёт или проведёт без верных проводок. Ничего не записано.`
        : `otherItemRef для «${operationKind}» — вместе с corrAccount (счёт, к субконто которого относится статья). Ничего не записано.`,
    );
  }
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
    // Вид субконто не прочитался — отказ: иначе значение субконто тихо пропало бы (ревью выпуска 5).
    const kind = await conn.client
      .getEntity<ODataEntity>(
        `ChartOfCharacteristicTypes_ВидыСубконтоХозрасчетные(${odataGuid(kindRef)})${buildQuery({ select: ["PredefinedDataName", "Description"] })}`,
      )
      .catch((e: unknown) => {
        throw new InputError(
          `Не удалось прочитать вид субконто ${i + 1} счёта ${code} (${e instanceof Error ? e.message : String(e)}): без него субконто не заполнить. Проверьте, что в составе OData опубликован план видов характеристик «Виды субконто хозрасчётные». Ничего не записано.`,
        );
      });
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
      // Статья ДДС стоит и в шапке документа: субконто по ней в документах рабочих баз пустое (57.01, 55.03) — не
      // требуем (проверено на рабочей БП 3.0, ревью выпуска 5).
      if (src.arg !== "cashflowItemRef") missing.push(src.label);
      continue;
    }
    fields[field] = value;
    fields[`${field}_Type`] = `StandardODATA.${src.type}`;
  }
  // Значение, которому у счёта нет субконто и которого нет в шапке, в документ не попадёт — сказать (ревью выпуска 5).
  const used = new Set(Object.values(fields).filter((v) => typeof v === "string"));
  for (const src of Object.values(KIND_SOURCES)) {
    const v = a[src.arg];
    if (v && v !== EMPTY && !used.has(v) && !inHeader.includes(src.arg))
      notes.push(`У счёта ${code} нет субконто «${src.label}» — это значение в документ не записано.`);
  }
  if (missing.length)
    throw new InputError(
      `Счёт ${code} ведётся по субконто: укажите ${missing.join(", ")}. Ничего не записано.`,
    );
  return { fields, notes };
}
