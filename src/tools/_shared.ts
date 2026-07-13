import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { ODataError } from "../odata/errors.js";
import { NotPublishedError, publicationHelp } from "../odata/publication.js";
import { AggregateOverflowError } from "../odata/aggregate.js";
import { logger } from "../logger.js";

/** Общее поле выбора базы — добавляется во все инструменты. */
export const databaseField = z
  .string()
  .optional()
  .describe("Имя базы 1С из list_databases. Без указания — база по умолчанию.");

/** Общее поле фильтра по организации (юрлицу) внутри базы. */
export const organizationField = z
  .string()
  .optional()
  .describe("Название организации (юрлица) для фильтра. Без указания — все организации базы.");

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
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
    structuredContent: data as Record<string, unknown>,
  };
}

/** Ошибка инструмента: помечаем isError, отдаём понятный текст вместо падения. */
export function fail(message: string): CallToolResult {
  return {
    content: [{ type: "text", text: message }],
    isError: true,
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
      return fail(publicationHelp(e.missing));
    }
    if (e instanceof AggregateOverflowError) {
      // Громкий, но аккуратный отказ — лучше явная ошибка, чем неполная сумма.
      logger.warn({ tool: toolName, entitySet: e.entitySet, cap: e.cap }, "aggregate overflow");
      return fail(e.message);
    }
    if (e instanceof ODataError) {
      logger.warn({ tool: toolName, kind: e.kind }, e.message);
      return fail(`[${e.kind}] ${e.message}`);
    }
    const msg = e instanceof Error ? e.message : String(e);
    logger.error({ tool: toolName }, msg);
    return fail(`Внутренняя ошибка инструмента ${toolName}: ${msg}`);
  }
}

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
