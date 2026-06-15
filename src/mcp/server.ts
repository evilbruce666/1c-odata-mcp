import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ServerContext } from "../context.js";
import { registerMetaTools } from "../tools/meta.js";
import { registerCounterpartyTools } from "../tools/counterparties.js";
import { registerDocumentTools } from "../tools/documents.js";
import { registerRegisterTools } from "../tools/registers.js";

/** Создаёт MCP-сервер и регистрирует все инструменты (read-only). */
export function createServer(ctx: ServerContext): McpServer {
  const server = new McpServer({
    name: "1c-odata-mcp",
    version: "0.1.0",
  });

  registerMetaTools(server, ctx);
  registerCounterpartyTools(server, ctx);
  registerDocumentTools(server, ctx);
  registerRegisterTools(server, ctx);

  return server;
}
