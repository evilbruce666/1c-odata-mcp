import { AsyncLocalStorage } from "node:async_hooks";

/**
 * Ответственный создаваемого документа (responsibleRef) на время одного вызова инструмента записи: обёртка сервера
 * проверяет пользователя и кладёт ссылку сюда, createCore ставит её в поле документа. Так параметр одинаков у всех
 * инструментов создания документов и не протаскивается через каждый сборщик.
 */
const store = new AsyncLocalStorage<{ ref: string }>();

export function withResponsible<T>(ref: string | undefined, fn: () => Promise<T>): Promise<T> {
  return ref ? store.run({ ref }, fn) : fn();
}

export function currentResponsible(): string | undefined {
  return store.getStore()?.ref;
}

/**
 * Поле ответственного документа по его схеме: «Ответственный» (Бухгалтерия; деньги и склад УТ), иначе «Менеджер»
 * (продажи, закупки и заказы УТ). «Автор» не заполняется — его ставит 1С (пользователь OData).
 */
export const RESPONSIBLE_FIELDS = ["Ответственный_Key", "Менеджер_Key"] as const;
