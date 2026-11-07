#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";
import { loadConfig, type HttpTransport } from "./config/env.js";
import { ServerContext } from "./context.js";
import { createServer } from "./mcp/server.js";
import { logger } from "./logger.js";

const SHUTDOWN_TIMEOUT_MS = 8_000;

async function main(): Promise<void> {
  const cfg = loadConfig();
  const ctx = new ServerContext(cfg);
  if (cfg.http) return runHttp(ctx, cfg.http);
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

async function runHttp(ctx: ServerContext, http: HttpTransport): Promise<void> {
  const { startHttpServer } = await import("./http.js");
  const { port, close } = await startHttpServer(ctx, http);
  logger.info({ host: http.host, port }, "1c-odata-mcp запущен (http, только чтение)");
  for (const signal of ["SIGTERM", "SIGINT"] as const) {
    process.once(signal, () => {
      // Запрос к 1С может идти минуты: контейнер не должен ждать его дольше обычного срока остановки.
      setTimeout(() => process.exit(0), SHUTDOWN_TIMEOUT_MS).unref();
      void close().then(() => process.exit(0));
    });
  }
}

main().catch((err) => {
  const message = err instanceof Error ? err.message : String(err);
  logger.error({ err: message }, "Фатальная ошибка старта");
  // Под MCP-клиентом лог идёт в файл — причину отказа дублируем в stderr: клиент показывает её
  // пользователю (процесс всё равно завершается, «чистый stderr» здесь уже не нужен).
  process.stderr.write(`1c-odata-mcp: ${message}\n`);
  process.exit(1);
});
