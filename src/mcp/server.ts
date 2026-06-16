import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ServerContext } from "../context.js";
import { registerMetaTools } from "../tools/meta.js";
import { registerCounterpartyTools } from "../tools/counterparties.js";
import { registerDocumentTools } from "../tools/documents.js";
import { registerRegisterTools } from "../tools/registers.js";
import { registerWriteTools } from "../tools/write.js";
import { READ_HINTS, WRITE_HINTS, DESTRUCTIVE_HINTS } from "../tools/_shared.js";

const INSTRUCTIONS =
  "Доступ к данным 1С:Предприятие через OData. По умолчанию только чтение " +
  "(аналитика: дебиторка, остатки, продажи, движение денег; справочники и документы). " +
  "Запись включается отдельно и по умолчанию работает в режиме предпросмотра (dry-run): " +
  "сначала показывайте пользователю, что будет создано/изменено, и выполняйте запись " +
  "только после явного согласия (confirm=true). У инструментов есть параметр database " +
  "(см. list_databases) и organization (см. list_organizations).";

/** Подсказки клиенту о характере инструмента — по имени (одна точка вместо правок 33 конфигов). */
function annotationsFor(name: string) {
  if (name === "post_document" || name === "mark_for_deletion") return DESTRUCTIVE_HINTS;
  if (/^(get_|list_|find_|describe_|search_|health_)/.test(name)) return READ_HINTS;
  return WRITE_HINTS;
}

/** Создаёт MCP-сервер и регистрирует все инструменты (чтение + гейтованная запись). */
export function createServer(ctx: ServerContext): McpServer {
  const server = new McpServer({ name: "1c-odata-mcp", version: "0.1.0" }, { instructions: INSTRUCTIONS });

  // Оборачиваем registerTool, чтобы каждому инструменту проставить аннотации по имени.
  const original = server.registerTool.bind(server);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (server as any).registerTool = (name: string, config: any, cb: any) =>
    original(name, { ...config, annotations: { ...annotationsFor(name), ...config.annotations } }, cb);

  registerMetaTools(server, ctx);
  registerCounterpartyTools(server, ctx);
  registerDocumentTools(server, ctx);
  registerRegisterTools(server, ctx);
  registerWriteTools(server, ctx);

  return server;
}
