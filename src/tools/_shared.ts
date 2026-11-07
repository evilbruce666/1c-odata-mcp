import { createHash } from "node:crypto";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { z } from "zod";
import { ODataError } from "../odata/errors.js";
import { NotPublishedError, publicationHelp } from "../odata/publication.js";
import { AggregateOverflowError } from "../odata/aggregate.js";
import { callResponseLimit } from "../odata/call-context.js";
import { InputError } from "../errors.js";
import { logger } from "../logger.js";

/** Общее поле выбора базы — добавляется во все инструменты. */
export const databaseField = z
  .string()
  .optional()
  .describe(
    "Имя базы 1С из list_databases. Без указания — база по умолчанию; если баз несколько и неясно какая, спросите пользователя.",
  );

/** Общее поле фильтра по организации (юрлицу) внутри базы. */
export const organizationField = z
  .string()
  .optional()
  .describe(
    "Организация (юрлицо) для фильтра: часть названия, ИНН или Ref_Key из list_organizations. Без указания — все организации базы.",
  );

// Формат + диапазон месяца (01-12) и дня (01-31): ловим «2025-13-99» на валидации.
const DATE_RE = /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/;

/**
 * Обязательное поле даты YYYY-MM-DD с валидацией формата (кривая дата → битый
 * $filter → 400/мусор). Для необязательного — `dateField(label).optional()`.
 */
export function dateField(label: string): z.ZodString {
  return z.string().regex(DATE_RE, "Дата должна быть в формате YYYY-MM-DD").describe(`${label} (YYYY-MM-DD)`);
}

/**
 * Текст ответа и расчёты размера страницы: единое место, где задан вид JSON в ответах. Компактный: отступы
 * занимают около 28% токенов (замер на отчёте о должниках), а точность ответов моделей от них не выигрывает.
 */
export const renderJson = (data: unknown): string => JSON.stringify(data);

/**
 * Аннотации инструментов (подсказки клиенту MCP о характере операции).
 * openWorldHint — обращение к внешней системе (1С). readOnlyHint — только чтение.
 * destructiveHint — операция меняет/портит данные (удаление, проведение).
 */
export const READ_HINTS = { readOnlyHint: true, openWorldHint: true } as const;
export const WRITE_HINTS = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
export const DESTRUCTIVE_HINTS = {
  readOnlyHint: false,
  destructiveHint: true,
  openWorldHint: true,
} as const;

/**
 * Успешный результат инструмента: JSON-данные в текстовом блоке + structuredContent
 * для инструментов с outputSchema (SDK валидирует structuredContent против неё —
 * см. src/schemas/output.ts). Все вызовы ok() в кодовой базе передают объект.
 */
export function ok(data: unknown): CallToolResult {
  return {
    content: [{ type: "text", text: renderJson(data) }],
    structuredContent: data as Record<string, unknown>,
  };
}

/**
 * Ключ _meta с видом ошибки: клиент различает отказы без разбора текста, а текст ошибки может содержать
 * сообщение 1С с данными. Не в structuredContent: там ответ должен соответствовать outputSchema.
 */
export const ERROR_KIND_META = "1c-odata-mcp/errorKind";

/**
 * Ошибка инструмента: помечаем isError, отдаём понятный текст вместо падения. kind: виды ODataError либо
 * refused (отказ инструмента по содержанию запроса), input, not_published, overflow, internal.
 */
export function fail(message: string, kind = "refused"): CallToolResult {
  return {
    content: [{ type: "text", text: message }],
    isError: true,
    _meta: { [ERROR_KIND_META]: kind },
  };
}

/**
 * Оборачивает тело инструмента: ловит ODataError/прочие и превращает
 * в аккуратный isError-результат. Сервер не должен падать из-за одного вызова.
 */
export async function guard(toolName: string, fn: () => Promise<CallToolResult>): Promise<CallToolResult> {
  try {
    return await fn();
  } catch (e) {
    if (e instanceof NotPublishedError) {
      logger.warn({ tool: toolName, missing: e.missing.map((m) => m.label) }, "not published");
      return fail(publicationHelp(e.missing), "not_published");
    }
    if (e instanceof InputError) {
      // Неверные аргументы вызова — отдаём текст как есть: это чинится уточнением
      // запроса, а не разбирательством с сервером.
      logger.info({ tool: toolName }, e.message);
      return fail(e.message, "input");
    }
    if (e instanceof AggregateOverflowError) {
      // Громкий, но аккуратный отказ — лучше явная ошибка, чем неполная сумма.
      logger.warn({ tool: toolName, entitySet: e.entitySet, cap: e.cap }, "aggregate overflow");
      return fail(e.message, "overflow");
    }
    if (e instanceof ODataError) {
      logger.warn({ tool: toolName, kind: e.kind }, e.message);
      return fail(`[${e.kind}] ${e.message}`, e.kind);
    }
    const msg = e instanceof Error ? e.message : String(e);
    logger.error({ tool: toolName }, msg);
    return fail(`Внутренняя ошибка инструмента ${toolName}: ${msg}`, "internal");
  }
}

/**
 * Предел текста ответа — с запасом к норме 50 000 символов (MAX_RESPONSE_CHARS в тестах): больший ответ
 * клиент уносит в файл или обрезает, и конец теряется уже у него.
 */
export const MAX_RESPONSE_CHARS = 45_000;

/** Норма ответа: больше клиент уносит в файл и теряет конец (проверяет test/fake-1c.test.ts). */
export const RESPONSE_CHARS_NORM = 50_000;

/** Предел текста текущего вызова: настройка ODATA_MAX_RESPONSE_CHARS, по умолчанию MAX_RESPONSE_CHARS. */
export const responseCharsLimit = (): number => callResponseLimit() ?? MAX_RESPONSE_CHARS;

/** Символов, которые значение занимает в ответе ok() как элемент массива верхнего уровня. */
export const elementChars = (v: unknown): number => {
  return renderJson(v).length + 1;
};

/** Сколько первых строк поместится в остаток предела. Хотя бы одна — иначе offset не сдвинется. */
export function fitToChars<T>(items: readonly T[], usedChars: number, maxChars = responseCharsLimit()): T[] {
  let used = usedChars;
  let n = 0;
  for (const item of items) {
    used += elementChars(item);
    if (used > maxChars && n > 0) break;
    n += 1;
  }
  return items.slice(0, n);
}

/**
 * Метка итогов ответа. Следующая страница передаёт её обратно: если итоги изменились, в 1С за это время
 * провели документы, и смещения сдвинулись. Склеивать такие страницы нельзя, читать надо с начала.
 */
export function snapshotFields(
  head: unknown,
  previous?: string,
): { fields: Record<string, unknown>; note?: string } {
  const snapshot = createHash("sha256").update(JSON.stringify(head)).digest("hex").slice(0, 16);
  const dataChanged = previous !== undefined && previous !== snapshot;
  return {
    fields: { snapshot, ...(dataChanged ? { dataChanged: true } : {}) },
    ...(dataChanged
      ? { note: "Данные в 1С изменились после прошлой страницы: прочитайте заново с offset=0." }
      : {}),
  };
}

export const snapshotField = z
  .string()
  .optional()
  .describe("snapshot прошлой страницы: если данные изменились, в ответе будет dataChanged");

/**
 * Страница ответа: не больше limit строк и столько, сколько поместится рядом с head. truncated значит
 * «есть ещё», nextOffset есть, только если продолжение можно прочитать; capHit — данные упёрлись в
 * лимит выборки, дальше читать нечего, и об этом нужно сказать.
 */
export function pageOf<T>(
  all: readonly T[],
  o: {
    offset: number;
    limit?: number;
    head: unknown;
    capHit?: boolean;
    unit?: string;
    /** snapshot прошлой страницы; undefined у инструментов, чей head меняется от страницы к странице. */
    snapshot?: string | undefined;
    tracksChanges?: boolean;
  },
): { items: T[]; page: Record<string, unknown>; note?: string } {
  const candidates = all.slice(o.offset, o.limit === undefined ? undefined : o.offset + o.limit);
  const items = fitToChars(candidates, renderJson(o.head).length);
  const next = o.offset + items.length;
  const more = next < all.length;
  const unit = o.unit ?? "строки";
  const note = more
    ? `Показаны ${unit} ${o.offset + 1}–${next}. Следующая страница: offset=${next}.`
    : o.capHit
      ? `Выборка упёрлась в лимит строк, продолжения нет: показаны не все. Уточните фильтр или период.`
      : o.offset > 0 && o.offset >= all.length
        ? `offset=${o.offset} за пределами: всего ${all.length}.`
        : undefined;
  const changes = o.tracksChanges ? snapshotFields(o.head, o.snapshot) : undefined;
  const fullNote = [changes?.note, note].filter(Boolean).join(" ");
  return {
    items,
    page: {
      offset: o.offset,
      returned: items.length,
      truncated: more || !!o.capHit,
      ...(more ? { nextOffset: next } : {}),
      ...(changes?.fields ?? {}),
    },
    ...(fullNote ? { note: fullNote } : {}),
  };
}

/** Поле offset для инструментов со страницами. */
export const offsetField = z
  .number()
  .int()
  .min(0)
  .default(0)
  .describe("С какой строки начать (0 — с начала); значение из nextOffset прошлого ответа");

/** Помечает усечённый по лимиту результат, чтобы модель это видела. */
export function withTruncationNote<T>(
  rows: T[],
  truncated: boolean,
  maxRows: number,
): { rows: T[]; count: number; truncated: boolean; note?: string } {
  return {
    rows,
    count: rows.length,
    truncated,
    ...(truncated ? { note: `Результат усечён до ${maxRows} строк. Уточните фильтр или период.` } : {}),
  };
}
