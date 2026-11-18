/**
 * Кэш справочных данных подключения (план счетов, список организаций): они меняются редко, а
 * запрашивались почти в каждом инструменте — на облачных 1С это по секунде на запрос, и запросы
 * одной учётной записи идут по очереди. Ключ владельца — объект подключения (WeakMap: кэш уходит
 * вместе с ним). Хранится промис: параллельные вызовы ждут одну загрузку. Ошибка не кэшируется.
 */

import { callAbortError, detachedFromCall, untilCallAborted } from "./call-context.js";

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

/** Сколько держать чтение для следующих страниц того же вызова: за это время клиент дочитывает список. */
export const PAGE_READ_TTL_MS = 60_000;

/**
 * Сколько разных чтений держать на подключение. Это полные выборки регистра: без предела бот с разными периодами
 * копил бы их в памяти.
 */
export const PAGE_READ_ENTRIES = 4;

const pageStore = new WeakMap<object, Map<string, Entry>>();

/**
 * Чтение для страниц одного списка. Первая страница (offset=0) всегда читает свежие данные и кладёт их сюда, следующие
 * в течение PAGE_READ_TTL_MS берут те же строки: иначе каждая страница заново читала бы весь регистр (6000 позиций —
 * 28 страниц и 28 чтений). Так страницы ещё и согласованы между собой: все из одного чтения. Окно истекло — чтение
 * заново, как раньше.
 *
 * Чтение идёт в контексте своего вызова, а не отдельно от него, как в cached: отмена клиентом и ODATA_TOOL_DEADLINE_MS
 * останавливают запросы к 1С (ревью). Готовое чтение другого вызова только переиспользуется; если тот вызов прервали,
 * продолжение читает само.
 */
export async function pageCached<T>(
  owner: object,
  key: string,
  firstPage: boolean,
  load: () => Promise<T>,
  now: () => number = Date.now,
): Promise<T> {
  let map = pageStore.get(owner);
  if (!map) {
    map = new Map();
    pageStore.set(owner, map);
  }
  for (const [k, e] of map) if (now() - e.at >= PAGE_READ_TTL_MS) map.delete(k);
  const hit = map.get(key);
  if (hit && !firstPage) {
    try {
      return await untilCallAborted(hit.value as Promise<T>);
    } catch (e) {
      // Своя отмена — наружу; чужое прерванное или неудачное чтение — читаем сами.
      if (callAbortError()) throw e;
    }
  }
  const value = load();
  const entry: Entry = { at: now(), value };
  // Новая запись — в конец: при переполнении уходит самая старая.
  map.delete(key);
  map.set(key, entry);
  while (map.size > PAGE_READ_ENTRIES) map.delete(map.keys().next().value!);
  value.catch(() => {
    if (map.get(key) === entry) map.delete(key);
  });
  return value;
}
