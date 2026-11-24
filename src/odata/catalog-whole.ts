import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { cached } from "./cache.js";
import { buildQuery, cmp } from "./query.js";

/**
 * Справочник целиком или по ссылкам — по цене в запросах к 1С. Документ на 300 строк при справочнике на 60 тыс.
 * позиций читал его целиком (13 страниц по 5000) дороже, чем пачками ссылок (6 запросов названий), а перемещение —
 * дважды: для названий и для проверки серий (ревью пачки Б). Теперь число позиций — одним запросом ($inlinecount,
 * кэш 10 минут), и справочник целиком читается, только если страниц меньше, чем пачек; прочитанный справочник общий
 * для названий и серий одного предпросмотра (кэш 60 секунд). Число не узнать — по ссылкам: ошибка оценки стоит
 * нескольких запросов, а не отказа.
 */

/** Позиций на странице чтения справочника целиком: строки короткие (3–4 поля). */
export const WHOLE_CATALOG_PAGE = 5000;
const COUNT_TTL_MS = 10 * 60_000;
/**
 * Справочник общий для названий и серий предпросмотра и для вызовов в ближайшие 60 с (подтверждение, следующий
 * документ): прочитанное только показывается (названия, заметка о сериях), позиции вне него — по ссылкам.
 */
const BRIEF_TTL_MS = 60_000;

async function propsOf(conn: Connection, set: string): Promise<Set<string>> {
  return new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);
}

/** Позиций (не папок) в справочнике; undefined — 1С не сказала. */
export function catalogItemCount(conn: Connection, set: string): Promise<number | undefined> {
  return cached(
    conn,
    `catalog-count:${set}`,
    async () => {
      try {
        const props = await propsOf(conn, set);
        const page = await conn.client.getCollection<ODataEntity>(
          `${set}${buildQuery({
            ...(props.has("IsFolder") ? { filter: cmp("IsFolder", "eq", "false") } : {}),
            select: ["Ref_Key"],
            top: 1,
            count: true,
          })}`,
        );
        const n = Number(page["odata.count"]);
        return Number.isFinite(n) && page["odata.count"] !== undefined ? n : undefined;
      } catch {
        return undefined;
      }
    },
    COUNT_TTL_MS,
  );
}

/** Когда справочник прочитан целиком (для кэша readCatalogBrief): свежий — бесплатен. */
const briefAt = new WeakMap<object, Map<string, number>>();

/**
 * Справочник целиком дешевле пачек ссылок по chunk: он уже прочитан (кэш 60 с) или страниц меньше, чем пачек, и
 * справочник помещается в предел.
 */
export async function wholeCatalogCheaper(
  conn: Connection,
  set: string,
  refs: number,
  chunk: number,
): Promise<boolean> {
  const at = briefAt.get(conn)?.get(set);
  if (at !== undefined && Date.now() - at < BRIEF_TTL_MS) return true;
  const n = await catalogItemCount(conn, set);
  if (n === undefined || n > conn.behavior.analyticsMaxRows) return false;
  return Math.floor(n / WHOLE_CATALOG_PAGE) + 1 < Math.ceil(refs / chunk);
}

/**
 * Позиции справочника кратко (Ref_Key, Description, Code, ВидНоменклатуры_Key — что опубликовано); ключ — Ref_Key
 * в нижнем регистре. Переполнение предела выборки — AggregateOverflowError (вызывающий читает по ссылкам).
 */
export function readCatalogBrief(conn: Connection, set: string): Promise<Map<string, ODataEntity>> {
  return cached(
    conn,
    `catalog-brief:${set}`,
    async () => {
      const props = await propsOf(conn, set);
      const select = ["Ref_Key", "Description", "Code", "ВидНоменклатуры_Key"].filter((f) => props.has(f));
      const { rows } = await fetchAllForAggregation(
        conn,
        set,
        { ...(props.has("IsFolder") ? { filter: cmp("IsFolder", "eq", "false") } : {}), select },
        undefined,
        WHOLE_CATALOG_PAGE,
      );
      const marks = briefAt.get(conn) ?? new Map<string, number>();
      marks.set(set, Date.now());
      briefAt.set(conn, marks);
      return new Map(rows.map((r) => [String(r["Ref_Key"]).toLowerCase(), r]));
    },
    BRIEF_TTL_MS,
  );
}
