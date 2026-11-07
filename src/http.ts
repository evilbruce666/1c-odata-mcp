import { createHash, timingSafeEqual } from "node:crypto";
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
import { withConnection } from "./odata/call-context.js";
import { logger } from "./logger.js";

/**
 * HTTP-транспорт для серверной установки. Без сессий: каждый запрос обслуживает новый McpServer над общим
 * ServerContext (сервер ничего не шлёт клиенту сам, состояния сессии у него нет). Защита: обязательный токен,
 * проверка Host и Origin от подмены DNS и запросов из браузера. TLS здесь нет: снаружи внутренней сети нужен
 * прокси. Только чтение: журнал записи и подтверждение человеком здесь не гарантированы (проверка в env.ts).
 */

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
  // Ошибка настройки (опечатка в TOOLS_ALLOW) должна остановить запуск, а не первый запрос.
  createServer(ctx);
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
