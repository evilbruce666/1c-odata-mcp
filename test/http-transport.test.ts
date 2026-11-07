import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, StreamableHTTPClientTransport } from "@modelcontextprotocol/client";
import { parseConfig, type HttpTransport } from "../src/config/env.js";
import { createHttpHandler, startHttpServer } from "../src/http.js";
import { fake1C } from "./helpers/fake-1c.js";

/** Настоящий fetch: fake1C подменяет глобальный, а клиенту нужен настоящий сокет. */
const realFetch = globalThis.fetch;
const TOKEN = "t".repeat(32);
const http: HttpTransport = { host: "127.0.0.1", port: 0, token: TOKEN };
const base = { ODATA_BASE_URL: "http://1c.test/secret/", ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };
const INIT = {
  jsonrpc: "2.0",
  id: 1,
  method: "initialize",
  params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "t", version: "1" } },
};

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

function post(handler: ReturnType<typeof createHttpHandler>, headers: Record<string, string>, path = "/mcp") {
  return handler.fetch(
    new Request(`http://localhost:8787${path}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json, text/event-stream",
        host: "localhost:8787",
        ...headers,
      },
      body: JSON.stringify(INIT),
    }),
  );
}

describe("настройка HTTP", () => {
  const good = { ...base, MCP_TRANSPORT: "http", MCP_HTTP_TOKEN: TOKEN };
  it("по умолчанию stdio", () => {
    const c = parseConfig(base);
    expect(c.http).toBeUndefined();
    expect(c.behavior.networkTransport).toBe(false);
  });
  it("значения по умолчанию: loopback, порт 8787", () => {
    expect(parseConfig(good).http).toEqual({
      host: "127.0.0.1",
      port: 8787,
      token: TOKEN,
      allowedHosts: undefined,
    });
  });
  it.each([
    ["без токена", { MCP_HTTP_TOKEN: undefined }, /MCP_HTTP_TOKEN/],
    ["короткий токен", { MCP_HTTP_TOKEN: "короткий" }, /MCP_HTTP_TOKEN/],
    ["запись включена", { READ_ONLY: "false" }, /READ_ONLY=true/],
    ["адрес не loopback без списка имён", { MCP_HTTP_HOST: "0.0.0.0" }, /MCP_HTTP_ALLOWED_HOSTS/],
    ["неверный транспорт", { MCP_TRANSPORT: "sse" }, /MCP_TRANSPORT/],
    ["неверный порт", { MCP_HTTP_PORT: "70000" }, /MCP_HTTP_PORT/],
    ["пробелы в токене", { MCP_HTTP_TOKEN: " ".repeat(40) }, /без пробелов/],
    [
      "комментарий шаблона вместо токена",
      { MCP_HTTP_TOKEN: "# не короче 32 символов, например openssl rand" },
      /без пробелов/,
    ],
    ["профиль all", { TOOLS_PROFILE: "all" }, /только для чтения/],
    ["запись в TOOLS_ALLOW", { TOOLS_ALLOW: "read.a.b,write.sales.create_invoice" }, /только для чтения/],
    ["порт в списке имён", { MCP_HTTP_ALLOWED_HOSTS: "mcp.local:8787" }, /без схемы и порта/],
    ["схема в списке имён", { MCP_HTTP_ALLOWED_HOSTS: "https://mcp.local" }, /без схемы и порта/],
  ])("%s — отказ при старте", (_name, extra, message) => {
    expect(() => parseConfig({ ...good, ...extra } as Record<string, string>)).toThrow(message);
  });
  it("токен не попадает в текст ошибки настройки", () => {
    expect(() => parseConfig({ ...good, READ_ONLY: "false" })).toThrow();
    try {
      parseConfig({ ...good, READ_ONLY: "false" });
    } catch (e) {
      expect(String(e)).not.toContain(TOKEN);
    }
  });
  it("в HTTP по умолчанию только чтение, имена хостов без учёта регистра", () => {
    const c = parseConfig({ ...good, MCP_HTTP_ALLOWED_HOSTS: "MCP.Local, [::1]" });
    expect(c.behavior.toolsProfile).toBe("read");
    expect(c.http?.allowedHosts).toEqual(["mcp.local", "[::1]"]);
    expect(parseConfig(base).behavior.toolsProfile).toBe("all");
  });
  it("адрес не loopback со списком имён", () => {
    const c = parseConfig({
      ...good,
      MCP_HTTP_HOST: "0.0.0.0",
      MCP_HTTP_ALLOWED_HOSTS: "1c-odata, mcp.local",
    });
    expect(c.http?.allowedHosts).toEqual(["1c-odata", "mcp.local"]);
  });
  it("без MCP_TRANSPORT=http токен не обязателен", () => {
    expect(parseConfig({ ...base, READ_ONLY: "false" }).http).toBeUndefined();
  });
});

describe("проверки запроса", () => {
  const handler = () => createHttpHandler(fake1C({ READ_ONLY: "true" }).ctx, http);
  const auth = { Authorization: `Bearer ${TOKEN}` };

  it("без токена и с неверным — 401 с WWW-Authenticate", async () => {
    for (const headers of [
      {},
      { Authorization: "Bearer wrong-token" },
      { Authorization: `Basic ${TOKEN}` },
    ]) {
      const r = await post(handler(), headers as Record<string, string>);
      expect(r.status).toBe(401);
      expect(r.headers.get("www-authenticate")).toBe("Bearer");
    }
  });
  it("токен в адресе не принимается", async () => {
    expect((await post(handler(), {}, `/mcp?token=${TOKEN}`)).status).toBe(401);
  });
  it("верный токен — ответ MCP", async () => {
    const r = await post(handler(), auth);
    expect(r.status).toBe(200);
    expect(await r.text()).toContain("1c-odata-mcp");
  });
  it("чужой Host и чужой Origin — отказ, даже с токеном", async () => {
    expect((await post(handler(), { ...auth, host: "evil.example" })).status).toBe(403);
    expect((await post(handler(), { ...auth, Origin: "https://evil.example" })).status).toBe(403);
    expect((await post(handler(), { ...auth, Origin: "null" })).status).toBe(403);
  });
  it("Host из разрешённого списка проходит, localhost при этом уже нет", async () => {
    const h = createHttpHandler(fake1C({ READ_ONLY: "true" }).ctx, { ...http, allowedHosts: ["1c-odata"] });
    expect((await post(h, { ...auth, host: "1c-odata:8787" })).status).toBe(200);
    expect((await post(h, auth)).status).toBe(403);
  });
  it("/healthz без токена, без 1С и без сведений о сервере", async () => {
    const f = fake1C({ READ_ONLY: "true" });
    const fetchSpy = vi.fn(globalThis.fetch);
    vi.stubGlobal("fetch", fetchSpy);
    const r = await createHttpHandler(f.ctx, http).fetch(new Request("http://x/healthz"));
    expect(r.status).toBe(200);
    expect(await r.json()).toEqual({ status: "ok" });
    expect(fetchSpy).not.toHaveBeenCalled();
    const wrong = await createHttpHandler(f.ctx, http).fetch(
      new Request("http://x/healthz", { method: "POST" }),
    );
    expect(wrong.status).toBe(405);
  });
  it("другой адрес — 404", async () => {
    expect((await post(handler(), auth, "/other")).status).toBe(404);
  });
  it("слишком большой запрос отклоняется", async () => {
    const r = await handler().fetch(
      new Request("http://localhost:8787/mcp", {
        method: "POST",
        headers: {
          ...auth,
          "Content-Type": "application/json",
          Accept: "application/json, text/event-stream",
          host: "localhost",
        },
        body: JSON.stringify({ ...INIT, pad: "x".repeat(1024 * 1024 + 1) }),
      }),
    );
    expect(r.status).toBe(413);
  });
});

describe("настоящий клиент по настоящему сокету", () => {
  async function serve(env: Record<string, string> = {}) {
    const f = fake1C({ READ_ONLY: "true", MCP_TRANSPORT: "http", MCP_HTTP_TOKEN: TOKEN, ...env });
    const server = await startHttpServer(f.ctx, http);
    const url = new URL(`http://127.0.0.1:${server.port}/mcp`);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(
      new StreamableHTTPClientTransport(url, {
        fetch: realFetch,
        requestInit: { headers: { Authorization: `Bearer ${TOKEN}` } },
      }),
    );
    return { f, server, client, url };
  }

  it("список инструментов, вызов и параллельные вызовы", async () => {
    const { server, client } = await serve();
    const { tools } = await client.listTools();
    expect(tools).toHaveLength(34);
    expect(tools.every((t) => t.name.startsWith("read."))).toBe(true);
    const calls = await Promise.all(
      Array.from({ length: 5 }, () => client.callTool({ name: "read.system.list_databases", arguments: {} })),
    );
    expect(calls.every((r) => !r.isError)).toBe(true);
    await client.close();
    await server.close();
  });

  it("health_check не показывает адрес базы", async () => {
    const { server, client } = await serve();
    const r = (await client.callTool({ name: "read.system.health_check", arguments: {} })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    expect(r.isError).toBeFalsy();
    expect(r.structuredContent["baseUrl"]).toBeUndefined();
    expect(r.content[0]!.text).not.toContain("baseUrl");
    await client.close();
    await server.close();
  });

  /** 1С, которая не отвечает на запрос данных, пока его не прервут. Схема отвечает как обычно. */
  function hang1C() {
    const answer = globalThis.fetch;
    const stats = { calls: 0, aborted: 0 };
    vi.stubGlobal("fetch", (url: string | URL, init?: RequestInit) => {
      if (String(url).includes("$metadata")) return answer(url, init);
      stats.calls += 1;
      return new Promise((_resolve, reject) => {
        init?.signal?.addEventListener("abort", () => {
          stats.aborted += 1;
          reject(new DOMException("aborted", "AbortError"));
        });
      });
    });
    return stats;
  }
  const FIND = {
    name: "read.counterparty.find_counterparty",
    arguments: { query: "Ромашка" },
  };

  it("разрыв соединения клиентом прерывает запрос к 1С, повторов нет", async () => {
    const { server, client, url } = await serve();
    const stats = hang1C();
    const controller = new AbortController();
    const call = realFetch(url, {
      method: "POST",
      signal: controller.signal,
      headers: {
        Authorization: `Bearer ${TOKEN}`,
        "Content-Type": "application/json",
        Accept: "application/json, text/event-stream",
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 7, method: "tools/call", params: FIND }),
    });
    await vi.waitFor(() => expect(stats.calls).toBe(1));
    controller.abort();
    await expect(call).rejects.toThrow();
    await vi.waitFor(() => expect(stats.aborted).toBe(1));
    await new Promise((r) => setTimeout(r, 1_200));
    expect(stats.calls).toBe(1);
    await client.close();
    await server.close();
  });

  it("лимит времени работает по HTTP без участия клиента", async () => {
    const f = fake1C({ READ_ONLY: "true", ODATA_TOOL_DEADLINE_MS: "100" });
    const server = await startHttpServer(f.ctx, http);
    const stats = hang1C();
    const client = new Client({ name: "t", version: "1" });
    await client.connect(
      new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${server.port}/mcp`), {
        fetch: realFetch,
        requestInit: { headers: { Authorization: `Bearer ${TOKEN}` } },
      }),
    );
    const r = (await client.callTool(FIND)) as { isError?: boolean; content: { text: string }[] };
    expect(r.isError).toBe(true);
    expect(r.content[0]!.text).toContain("[deadline]");
    expect(stats.aborted).toBe(1);
    await client.close();
    await server.close();
  });

  it("клиент без токена не подключается", async () => {
    const { server, client, url } = await serve();
    const stranger = new Client({ name: "x", version: "1" });
    await expect(
      stranger.connect(new StreamableHTTPClientTransport(url, { fetch: realFetch })),
    ).rejects.toThrow();
    await client.close();
    await server.close();
  });

  it("опечатка в TOOLS_ALLOW останавливает запуск", async () => {
    const f = fake1C({ READ_ONLY: "true", TOOLS_ALLOW: "read.nope.none" });
    await expect(startHttpServer(f.ctx, http)).rejects.toThrow(/read\.nope\.none/);
  });
});
