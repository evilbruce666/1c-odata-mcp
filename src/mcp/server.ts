import { readFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { createResultSchema } from "../schemas/output.js";
import { fingerprintWriteInput, withWriteOperation } from "../odata/write-operation-context.js";
import { registerMetaTools } from "../tools/meta.js";
import { registerCounterpartyTools } from "../tools/counterparties.js";
import { registerDocumentTools } from "../tools/documents.js";
import { registerRegisterTools } from "../tools/registers.js";
import { registerCashflowTools } from "../tools/cashflow.js";
import { registerSalesTools } from "../tools/sales.js";
import { registerOrganizationTools } from "../tools/organization.js";
import { registerWriteTools } from "../tools/write.js";
import { READ_HINTS, WRITE_HINTS, DESTRUCTIVE_HINTS, fail, guard, ok } from "../tools/_shared.js";

/** Версия берётся из package.json (в собранном пакете он на два уровня выше dist/mcp/). */
function readVersion(): string {
  try {
    const url = new URL("../../package.json", import.meta.url);
    const pkg = JSON.parse(readFileSync(url, "utf8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

const INSTRUCTIONS =
  "Доступ к данным 1С:Предприятие через OData. По умолчанию только чтение " +
  "(аналитика: дебиторка, остатки, продажи, движение денег; справочники и документы). " +
  "Запись включается отдельно и по умолчанию работает в режиме предпросмотра (dry-run): " +
  "сначала показывайте пользователю, что будет создано/изменено, и выполняйте запись " +
  "только после явного согласия (confirm=true). У инструментов есть параметр database " +
  "(см. read.system.list_databases) и organization (см. read.system.list_organizations). " +
  "Для создания объекта при подтверждении передавайте operationId из предпросмотра; " +
  "при таймауте повторяйте подтверждение с тем же id и теми же аргументами, а не с новым предпросмотром. " +
  "Это защищает от дубликата, если ответ 1С потерялся.";

/**
 * Правки документа, которые нельзя слепо повторять: инструмент перечитывает документ, и повтор
 * добавил бы ещё строку или удалил соседнюю. Как и создание, защищены operationId и журналом.
 */
const LINE_OPERATION_TOOLS = new Set([
  "write.document.add_document_line",
  "write.document.remove_document_line",
]);

/** Подсказки клиенту о характере инструмента — по имени (одна точка вместо правок 55 конфигов). */
function annotationsFor(name: string) {
  if (name === "write.document.post_document" || name === "write.entity.mark_for_deletion")
    return DESTRUCTIVE_HINTS;
  if (name.startsWith("read.")) return READ_HINTS;
  return WRITE_HINTS;
}

/** Создаёт MCP-сервер и регистрирует все инструменты (чтение + гейтованная запись). */
export function createServer(ctx: ServerContext): McpServer {
  const server = new McpServer(
    { name: "1c-odata-mcp", version: readVersion() },
    { instructions: INSTRUCTIONS },
  );

  // Оборачиваем registerTool, чтобы каждому инструменту проставить аннотации по имени.
  const original = server.registerTool.bind(server) as unknown as (
    name: string,
    config: Record<string, unknown>,
    cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
  ) => unknown;
  const registerTool = (
    name: string,
    config: Record<string, unknown>,
    cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
  ) => {
    const annotatedConfig = {
      ...config,
      annotations: {
        ...annotationsFor(name),
        ...(config.annotations as Record<string, unknown> | undefined),
      },
    };
    const isCreateTool = name.startsWith("write.") && config.outputSchema === createResultSchema;
    if (!isCreateTool && !LINE_OPERATION_TOOLS.has(name)) return original(name, annotatedConfig, cb);

    const inputSchema = {
      ...(config.inputSchema as Record<string, unknown>),
      operationId: z
        .string()
        .uuid()
        .optional()
        .describe("UUID из предпросмотра. Передайте его при confirm=true, чтобы повтор не создал дубликат."),
    };
    const wrapped = async (args: Record<string, unknown>, extra: unknown): Promise<CallToolResult> => {
      const operationId = typeof args.operationId === "string" ? args.operationId : undefined;
      if (args.confirm === true && !operationId) {
        return fail("Сначала выполните dry-run. При confirm=true передайте operationId из его результата.");
      }
      const token = operationId ?? randomUUID();
      const requestHash = fingerprintWriteInput(name, args as Record<string, unknown>);
      // Повтор уже подтверждённой операции: ответ из журнала — до того, как инструмент заново
      // прочитает 1С (иначе правка строк применилась бы к уже изменённому документу).
      if (args.confirm === true && operationId) {
        let replay: CallToolResult | undefined;
        const checked = await guard(name, async () => {
          const conn = ctx.db(typeof args.database === "string" ? args.database : undefined);
          const settled = await conn.client.operationSettled(operationId, requestHash);
          if (settled) {
            replay = ok({
              ...(isCreateTool ? { created: true } : { updated: true }),
              replayed: true,
              note: "Операция с этим operationId уже выполнена; повторный запрос в 1С не отправлялся.",
              database: conn.cfg.name,
              entitySet: settled.entitySet,
              ref: settled.result["Ref_Key"],
              ...(settled.result["Code"] ? { code: settled.result["Code"] } : {}),
              operationId,
            });
          }
          return ok({});
        });
        if (checked.isError) return checked;
        if (replay) return replay;
      }
      return withWriteOperation(token, requestHash, async () => {
        const result = await cb(args, extra);
        if (result?.isError || !result?.structuredContent || typeof result.structuredContent !== "object")
          return result;
        const structuredContent = { ...result.structuredContent, operationId: token };
        const content = result.content.map((item) =>
          item.type === "text" ? { ...item, text: JSON.stringify(structuredContent, null, 2) } : item,
        );
        return { ...result, content, structuredContent };
      });
    };
    return original(name, { ...annotatedConfig, inputSchema }, wrapped);
  };
  Reflect.set(server, "registerTool", (name: string, config: unknown, cb: unknown) =>
    registerTool(
      name,
      config as Record<string, unknown>,
      cb as (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
    ),
  );

  registerMetaTools(server, ctx);
  registerCounterpartyTools(server, ctx);
  registerDocumentTools(server, ctx);
  registerRegisterTools(server, ctx);
  registerCashflowTools(server, ctx);
  registerSalesTools(server, ctx);
  registerOrganizationTools(server, ctx);
  registerWriteTools(server, ctx);

  return server;
}
