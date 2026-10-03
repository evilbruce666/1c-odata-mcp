#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadConfig } from "./config/env.js";
import { ServerContext } from "./context.js";
import { createServer } from "./mcp/server.js";
import { logger } from "./logger.js";

async function main(): Promise<void> {
  const cfg = loadConfig();
  const ctx = new ServerContext(cfg);
  const server = createServer(ctx);

  const transport = new StdioServerTransport();
  await server.connect(transport);

  // Логи только в stderr — stdout занят JSON-RPC.
  logger.info(
    {
      databases: cfg.connections.map((c) => c.name),
      default: cfg.defaultName,
      readOnly: cfg.behavior.readOnly,
    },
    "1c-odata-mcp запущен (stdio)",
  );
}

main().catch((err) => {
  const message = err instanceof Error ? err.message : String(err);
  logger.error({ err: message }, "Фатальная ошибка старта");
  // Под MCP-клиентом лог идёт в файл — причину отказа дублируем в stderr: клиент показывает её
  // пользователю (процесс всё равно завершается, «чистый stderr» здесь уже не нужен).
  process.stderr.write(`1c-odata-mcp: ${message}\n`);
  process.exit(1);
});
