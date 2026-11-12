import type { Connection } from "../context.js";
import { resolveEntity } from "../config/mapping.js";
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
  const set = resolveEntity(candidates, await conn.available());
  if (set) return set;
  // Регистр бухгалтерии и план счетов в Управлении торговлей отсутствуют в самой конфигурации: совет опубликовать их
  // неверен, отказ ведёт к инструментам, которые в такой базе работают.
  const accounting = candidates.some((c) => /^(AccountingRegister|ChartOfAccounts)_/.test(c));
  if (accounting && (await tradeWithoutAccounting(conn))) {
    throw new NotPublishedError(
      [{ label, candidates }],
      tradeRefusal("Этот инструмент работает по бухгалтерскому учёту."),
    );
  }
  throw new NotPublishedError([{ label, candidates }]);
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

/** Проверяет, что конкретный EntitySet опубликован (для имён, заданных пользователем). */
export function ensurePublished(available: ReadonlySet<string>, entitySet: string, label?: string): void {
  if (!available.has(entitySet)) {
    throw new NotPublishedError([{ label: label ?? entitySet, candidates: [entitySet] }]);
  }
}

/** Собирает текст вежливого фоллбэка по списку недостающих объектов. */
export function publicationHelp(missing: NotPublishedError["missing"]): string {
  const lines = missing.map((m) => `  • ${m.label} (${m.candidates.join(" или ")})`).join("\n");
  return (
    "Эти объекты 1С отсутствуют в доступной OData-схеме, поэтому операция недоступна. " +
    "Добавьте их в 1С: «Все функции → Обработки → Настройка стандартного интерфейса OData → вкладка Состав», затем сохраните:\n" +
    lines +
    "\nВозможны также другое имя объекта, его отсутствие в этой конфигурации или ограничения прав."
  );
}
