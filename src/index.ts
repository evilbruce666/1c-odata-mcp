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
    { baseUrl: cfg.ODATA_BASE_URL, readOnly: cfg.READ_ONLY },
    "1c-odata-mcp запущен (stdio)",
  );
}

main().catch((err) => {
  logger.error({ err: err instanceof Error ? err.message : String(err) }, "Фатальная ошибка старта");
  process.exit(1);
});
