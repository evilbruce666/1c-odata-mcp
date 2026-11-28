import type { Connection } from "../context.js";
import { resolveEntity } from "../config/mapping.js";
import { SYNONYMS } from "../config/synonyms.js";
import { tradeRefusal, tradeWithoutAccounting } from "./configuration.js";

/**
 * Требуемый объект отсутствует в доступной OData-схеме. guard() превращает ошибку
 * в подсказку по «Составу», имени объекта, конфигурации и правам.
 */
export class NotPublishedError extends Error {
  readonly missing: Array<{ label: string; candidates: readonly string[] }>;
  /**
   * Текст вместо совета опубликовать объект: когда объекта нет в самой конфигурации, а данные есть в другом
   * инструменте (в УТ нет Хозрасчётного — остатки в регистре склада).
   */
  readonly instead: string | undefined;
  constructor(missing: Array<{ label: string; candidates: readonly string[] }>, instead?: string) {
    super("Требуемые объекты отсутствуют в доступной OData-схеме");
    this.name = "NotPublishedError";
    this.missing = missing;
    this.instead = instead;
  }
}

/**
 * Возвращает имя опубликованного EntitySet среди кандидатов или бросает
 * NotPublishedError с человекочитаемым ярлыком.
 */
export async function requireEntity(
  conn: Connection,
  candidates: readonly string[],
  label: string,
): Promise<string> {
  const known = resolveEntity(candidates, await conn.available());
  if (known) return known;
  // Регистр бухгалтерии и план счетов в Управлении торговлей отсутствуют в самой конфигурации: совет опубликовать их
  // неверен, отказ ведёт к инструментам, которые в такой базе работают.
  const accounting = candidates.some((c) => /^(AccountingRegister|ChartOfAccounts)_/.test(c));
  if (accounting && (await tradeWithoutAccounting(conn))) {
    throw new NotPublishedError(
      [{ label, candidates }],
      tradeRefusal("Этот инструмент работает по бухгалтерскому учёту."),
    );
  }
  const added = await afterRecheck(conn, candidates);
  if (added) return added;
  throw new NotPublishedError([{ label, candidates }]);
}

/** Не чаще раза в минуту: столько после чтения схемы повторный отказ строится на той же. */
export const RECHECK_AFTER_MS = 60_000;

/** Объект, появившийся в «Составе OData» после того, как схему прочитали. */
async function afterRecheck(conn: Connection, candidates: readonly string[]): Promise<string | undefined> {
  if (!(await conn.client?.refreshIfOlderThan?.(RECHECK_AFTER_MS))) return undefined;
  return resolveEntity(candidates, await conn.available());
}

/** Как requireEntity — для имени объекта, которое задал пользователь. */
export async function requirePublished(conn: Connection, entitySet: string, label?: string): Promise<void> {
  const available = await conn.available();
  if (available.has(entitySet) || (await afterRecheck(conn, [entitySet]))) return;
  ensurePublished(available, entitySet, label);
}

/**
 * Отказ инструмента, который построен на объектах Бухгалтерии, в торговой базе: вместо совета опубликовать их — что
 * в такой базе работает. what — одна фраза: на чём построен инструмент.
 */
export async function refuseOnTrade(conn: Connection, label: string, what: string): Promise<void> {
  if (await tradeWithoutAccounting(conn)) {
    throw new NotPublishedError([{ label, candidates: [] }], tradeRefusal(what));
  }
}

/** Отказ, если конкретного EntitySet нет в переданном наборе; с перечитыванием схемы — requirePublished. */
function ensurePublished(available: ReadonlySet<string>, entitySet: string, label?: string): void {
  if (!available.has(entitySet)) {
    throw new NotPublishedError([{ label: label ?? entitySet, candidates: [entitySet] }]);
  }
}

const KINDS: Readonly<Record<string, string>> = {
  Catalog: "справочник",
  Document: "документ",
  InformationRegister: "регистр сведений",
  AccumulationRegister: "регистр накопления",
  AccountingRegister: "регистр бухгалтерии",
  ChartOfAccounts: "план счетов",
  ChartOfCharacteristicTypes: "план видов характеристик",
};

/**
 * Объект так, как его искать в «Составе OData»: вид и название, в скобках техническое имя. Таблица записей регистра
 * («…_RecordType») в составе отдельно не отмечается — называется сам регистр. Без проверенного названия — только имя.
 */
export function describeObject(entitySet: string): string {
  const name = entitySet.replace(/_RecordType$/, "");
  const synonym = SYNONYMS[name];
  const kind = KINDS[name.slice(0, name.indexOf("_"))];
  if (!synonym || !kind) return name;
  const title =
    synonym.bp && synonym.ut
      ? `«${synonym.bp}» (в Управлении торговлей — «${synonym.ut}»)`
      : `«${synonym.bp ?? synonym.ut}»`;
  return `${kind} ${title} [${name}]`;
}

/** Собирает текст вежливого фоллбэка по списку недостающих объектов. */
export function publicationHelp(missing: NotPublishedError["missing"]): string {
  const lines = missing
    .map((m) => `  • ${m.label}: ${[...new Set(m.candidates.map(describeObject))].join(" или ")}`)
    .join("\n");
  return (
    "Эти объекты 1С отсутствуют в доступной OData-схеме, поэтому операция недоступна. " +
    "Добавьте их в 1С: «Все функции → Обработки → Настройка стандартного интерфейса OData → вкладка Состав», затем сохраните:\n" +
    lines +
    "\nНазвания — как в типовых конфигурациях; в вашем релизе они могут отличаться, в скобках — техническое имя." +
    "\nПосле сохранения подождите минуту и повторите запрос: перед отказом коннектор перечитывает схему." +
    "\nВозможны также другое имя объекта, его отсутствие в этой конфигурации или ограничения прав."
  );
}
