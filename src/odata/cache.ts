/**
 * Кэш справочных данных подключения (план счетов, список организаций): они меняются редко, а
 * запрашивались почти в каждом инструменте — на облачных 1С это по секунде на запрос, и запросы
 * одной учётной записи идут по очереди. Ключ владельца — объект подключения (WeakMap: кэш уходит
 * вместе с ним). Хранится промис: параллельные вызовы ждут одну загрузку. Ошибка не кэшируется.
 */

import { detachedFromCall, untilCallAborted } from "./call-context.js";

export const REFERENCE_TTL_MS = 10 * 60_000;

interface Entry {
  at: number;
  value: Promise<unknown>;
}

const store = new WeakMap<object, Map<string, Entry>>();

export function cached<T>(
  owner: object,
  key: string,
  load: () => Promise<T>,
  ttlMs: number = REFERENCE_TTL_MS,
  now: () => number = Date.now,
): Promise<T> {
  let map = store.get(owner);
  if (!map) {
    map = new Map();
    store.set(owner, map);
  }
  const hit = map.get(key);
  if (hit && now() - hit.at < ttlMs) return untilCallAborted(hit.value as Promise<T>);
  // Загрузку ждут и другие вызовы: отмена начавшего её прерывает только его ожидание.
  const value = detachedFromCall(load);
  const entry: Entry = { at: now(), value };
  map.set(key, entry);
  value.catch(() => {
    if (map.get(key) === entry) map.delete(key);
  });
  return untilCallAborted(value);
}

/** Сбросить запись — например, когда нужного объекта в кэше нет (его могли только что создать в 1С). */
export function invalidate(owner: object, key: string): void {
  store.get(owner)?.delete(key);
}

/** Перечитывать при промахе не чаще: несуществующий код не должен каждый раз сбрасывать кэш. */
export const REFRESH_MIN_AGE_MS = 30_000;

/**
 * Поиск в кэшированном списке с одним перечитыванием при промахе: объект, заведённый в 1С после
 * загрузки кэша, не должен давать ложное «не найдено». Свежий кэш (моложе REFRESH_MIN_AGE_MS) не
 * перечитывается — промах в нём настоящий.
 */
export async function lookupWithRefresh<T, R>(
  owner: object,
  key: string,
  load: () => Promise<T>,
  pick: (value: T) => R,
  missing: (result: R) => boolean,
  now: () => number = Date.now,
): Promise<R> {
  const first = pick(await cached(owner, key, load, REFERENCE_TTL_MS, now));
  const entry = store.get(owner)?.get(key);
  if (!missing(first) || !entry || now() - entry.at < REFRESH_MIN_AGE_MS) return first;
  invalidate(owner, key);
  return pick(await cached(owner, key, load, REFERENCE_TTL_MS, now));
}
