import pino from "pino";

/**
 * Логгер пишет ТОЛЬКО в stderr.
 *
 * Критично для MCP: stdout зарезервирован под JSON-RPC транспорт (stdio).
 * Любая запись в stdout сломает протокол, поэтому destination = 2 (stderr).
 */
export const logger = pino(
  {
    level: process.env.LOG_LEVEL ?? "info",
    base: undefined,
    redact: {
      // Пароли и заголовки авторизации никогда не попадают в логи.
      paths: [
        "password",
        "headers.authorization",
        "headers.Authorization",
        "*.password",
        "*.authorization",
      ],
      censor: "[redacted]",
    },
  },
  pino.destination(2),
);

export type Logger = typeof logger;
