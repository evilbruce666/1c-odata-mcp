import { createHash, randomUUID, timingSafeEqual } from "node:crypto";
import { mkdir, open, rm } from "node:fs/promises";
import { join } from "node:path";
import { createServer as createHttpServer, type Server } from "node:http";
import type { AddressInfo } from "node:net";
import {
  createMcpHandler,
  hostHeaderValidationResponse,
  localhostAllowedHostnames,
  originValidationResponse,
} from "@modelcontextprotocol/server";
import type { HttpTransport } from "./config/env.js";
import type { ServerContext } from "./context.js";
import { createServer } from "./mcp/server.js";
import { redactInResponses } from "./tools/_shared.js";
import { withConnection } from "./odata/call-context.js";
import { logger } from "./logger.js";

/**
 * HTTP-транспорт для серверной установки. Без сессий: каждый запрос обслуживает новый McpServer над общим
 * ServerContext (сервер ничего не шлёт клиенту сам, состояния сессии у него нет). Защита: обязательный токен,
 * проверка Host и Origin от подмены DNS и запросов из браузера. TLS здесь нет: снаружи внутренней сети нужен
 * прокси. По умолчанию только чтение; запись — при READ_ONLY=false с журналом на томе и инструментами поимённо
 * (условия — в env.ts, доступность журнала — при старте ниже). Кто подтверждает запись — ответственность интеграции.
 */

/**
 * Каталог журнала записи доступен на запись: создать и удалить пробный файл. Без этого первое подтверждение
 * упало бы уже после предпросмотра, а журнал без тома не переживает рестарт (#67).
 */
export async function assertWriteJournalWritable(dir: string): Promise<void> {
  try {
    await mkdir(dir, { recursive: true, mode: 0o700 });
    const probe = join(dir, `.probe-${randomUUID()}`);
    await (await open(probe, "wx", 0o600)).close();
    await rm(probe, { force: true });
  } catch (e) {
    throw new Error(
      `Запись по HTTP: каталог журнала ODATA_WRITE_JOURNAL_DIR недоступен на запись (${e instanceof Error ? e.message : String(e)}). ` +
        "Подключите том с правом записи для пользователя контейнера или работайте только для чтения: READ_ONLY=true.",
      { cause: e },
    );
  }
}

const MCP_PATH = "/mcp";
/** Аргументы инструментов чтения невелики; больший запрос отклоняется до разбора. */
const MAX_BODY_BYTES = 1024 * 1024;

const digest = (value: string) => createHash("sha256").update(value).digest();

/** Сравнение за постоянное время: хэши одной длины, а время ответа не выдаёт совпавшее начало токена. */
function validToken(header: string | null, token: string): boolean {
  const presented = /^Bearer (.+)$/i.exec(header ?? "")?.[1];
  return !!presented && timingSafeEqual(digest(presented), digest(token));
}

const json = (status: number, body: unknown, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...headers },
  });

/** Обработчик запроса: проверки до передачи в MCP. Отдельно от сокета, чтобы проверять его напрямую. */
export function createHttpHandler(ctx: ServerContext, cfg: HttpTransport) {
  const mcp = createMcpHandler(() => createServer(ctx), {
    maxRequestBodySize: MAX_BODY_BYTES,
    onerror: (e) => logger.warn({ err: e.message }, "http: запрос отклонён"),
  });
  const allowedHosts = cfg.allowedHosts ?? localhostAllowedHostnames();
  return {
    close: () => mcp.close(),
    fetch: async (request: Request): Promise<Response> => {
      const { pathname } = new URL(request.url);
      // Для проверки живости контейнера: без токена, без обращения к 1С и без сведений о сервере.
      if (pathname === "/healthz") {
        return request.method === "GET"
          ? json(200, { status: "ok" })
          : json(405, { error: "method not allowed" });
      }
      if (pathname !== MCP_PATH) return json(404, { error: "not found" });
      const rejected =
        hostHeaderValidationResponse(request, allowedHosts) ??
        originValidationResponse(request, allowedHosts);
      if (rejected) return rejected;
      if (!validToken(request.headers.get("authorization"), cfg.token)) {
        return json(401, { error: "unauthorized" }, { "WWW-Authenticate": "Bearer" });
      }
      return withConnection(request.signal, () => mcp.fetch(request));
    },
  };
}

/** Запускает HTTP-сервер; возвращает порт и остановку с ожиданием текущих запросов. */
export async function startHttpServer(
  ctx: ServerContext,
  cfg: HttpTransport,
): Promise<{ port: number; close: () => Promise<void> }> {
  // Адаптер нужен только этому режиму: установкам со stdio его зависимости не загружаются.
  const { toNodeHandler } = await import("@modelcontextprotocol/node");
  // Ошибка настройки (опечатка в TOOLS_ALLOW, журнал без права записи) должна остановить запуск, а не первый запрос.
  createServer(ctx);
  if (!ctx.behavior.readOnly) await assertWriteJournalWritable(ctx.behavior.writeJournalDir);
  // Путь к журналу клиенту по сети не отдаётся, в том числе в текстах ошибок файловой системы.
  redactInResponses(ctx.behavior.writeJournalDir);
  const handler = createHttpHandler(ctx, cfg);
  const node = toNodeHandler(handler, {
    // Адаптер читает тело до проверки токена: без этого чужой запрос занял бы 4 МиБ.
    maxRequestBodySize: MAX_BODY_BYTES,
    onerror: (e) => logger.error({ err: e.message }, "http: сбой обработки запроса"),
  });
  const server: Server = createHttpServer((req, res) => void node(req, res));
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(cfg.port, cfg.host, () => {
      server.off("error", reject);
      resolve();
    });
  });
  return {
    port: (server.address() as AddressInfo).port,
    close: async () => {
      await handler.close();
      await new Promise<void>((resolve) => {
        server.close(() => resolve());
        server.closeIdleConnections();
      });
    },
  };
}
