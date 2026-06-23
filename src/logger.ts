import { tmpdir } from "node:os";
import { join } from "node:path";
import pino from "pino";

/**
 * Логгер НИКОГДА не пишет в stdout — там JSON-RPC stdio-транспорт; любая запись
 * в stdout сломает протокол.
 *
 * Куда идёт вывод:
 *  - в терминале (stdin = tty) — в stderr, как обычно (удобно при отладке);
 *  - под MCP-клиентом (stdin = pipe) — в файл `<tmpdir>/1c-odata-mcp/server.log`.
 *
 * Причина файла: некоторые MCP-клиенты (Kilo Code, OpenCode и др.) трактуют ЛЮБОЙ
 * вывод в stderr как фатальную ошибку — оставляем их stderr чистым. Принудительно
 * вернуть вывод в stderr: переменная `MCP_LOG_STDERR=1`.
 */
function makeDestination(): pino.DestinationStream {
  const forceStderr = process.env.MCP_LOG_STDERR === "1" || process.env.MCP_LOG_STDERR === "true";
  const underClient = !process.stdin.isTTY;
  if (!forceStderr && underClient) {
    try {
      return pino.destination({ dest: join(tmpdir(), "1c-odata-mcp", "server.log"), mkdir: true });
    } catch {
      // Не удалось открыть файл — падать из-за логов нельзя, откатываемся на stderr.
    }
  }
  return pino.destination(2);
}

/**
 * Логгер пишет в stderr (терминал) или в файл (под MCP-клиентом) — см. makeDestination.
 */
export const logger = pino(
  {
    level: process.env.LOG_LEVEL ?? "info",
    base: undefined,
    redact: {
      // Пароли и заголовки авторизации никогда не попадают в логи.
      paths: ["password", "headers.authorization", "headers.Authorization", "*.password", "*.authorization"],
      censor: "[redacted]",
    },
  },
  makeDestination(),
);

export type Logger = typeof logger;
