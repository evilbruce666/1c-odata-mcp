import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { ODataError } from "../odata/errors.js";
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

/** Успешный результат инструмента: JSON-данные в текстовом блоке. */
export function ok(data: unknown): CallToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
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
export async function guard(
  toolName: string,
  fn: () => Promise<CallToolResult>,
): Promise<CallToolResult> {
  try {
    return await fn();
  } catch (e) {
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
    ...(truncated
      ? { note: `Результат усечён до ${maxRows} строк. Уточните фильтр или период.` }
      : {}),
  };
}
