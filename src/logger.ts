import { chmodSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
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
  // В HTTP-режиме клиента на stdin нет, а журнал контейнера — это stderr.
  const underClient = !process.stdin.isTTY && process.env.MCP_TRANSPORT !== "http";
  if (!forceStderr && underClient) {
    try {
      // sync: файл открывается сразу. Асинхронный пишет только после открытия — ошибка старта,
      // записанная перед process.exit(1), терялась, а pino падал с «sonic boom is not ready yet».
      return logFile(join(tmpdir(), "1c-odata-mcp", "server.log"));
    } catch {
      // Не удалось открыть файл — падать из-за логов нельзя, откатываемся на stderr.
    }
  }
  return pino.destination(2);
}

/**
 * Файл лога только для владельца: каталог 0700, файл 0600. В логе — тексты ошибок 1С (названия, ИНН), а на Linux
 * с общим /tmp путь предсказуем, и с правами по umask (обычно 0644) его читали другие пользователи. Права
 * выставляются и уже существующим каталогу и файлу; на Windows chmod ничего не делает — там tmpdir у каждого свой.
 */
export function logFile(path: string): pino.DestinationStream {
  const dir = dirname(path);
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  try {
    chmodSync(dir, 0o700);
  } catch {
    // Каталог чужой (другой пользователь создал его раньше) — открыть файл всё равно попробуем.
  }
  const dest = pino.destination({ dest: path, sync: true, mode: 0o600 });
  try {
    chmodSync(path, 0o600);
  } catch {
    // Права не сменить — пишем как есть, падать из-за лога нельзя.
  }
  return dest;
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
