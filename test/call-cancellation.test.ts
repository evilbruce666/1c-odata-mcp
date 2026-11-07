import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { cached } from "../src/odata/cache.js";
import { callAbortError, withToolCall } from "../src/odata/call-context.js";
import { ODataError } from "../src/odata/errors.js";
import { RequestLimiter } from "../src/odata/limiter.js";
import { fake1C } from "./helpers/fake-1c.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** fetch 1С, который не отвечает, пока запрос не прервут. Считает обращения и прерывания. */
function hangingFetch() {
  const stats = { calls: 0, aborted: 0 };
  vi.stubGlobal(
    "fetch",
    vi.fn(
      (_url: string | URL, init?: { signal?: AbortSignal }) =>
        new Promise((_resolve, reject) => {
          stats.calls += 1;
          init?.signal?.addEventListener("abort", () => {
            stats.aborted += 1;
            reject(new DOMException("aborted", "AbortError"));
          });
        }),
    ),
  );
  return stats;
}

async function connect(env: Record<string, string> = {}) {
  const fake = fake1C({ READ_ONLY: "true", ODATA_RETRIES: "3", ...env });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  return { fake, client };
}

describe("отмена вызова чтения", () => {
  it("настоящий клиент отменяет вызов: запрос прерван, повторов нет", async () => {
    const { client } = await connect();
    // Схема отвечает как обычно (её загрузка общая и отменой не прерывается), данные висят.
    const answer = globalThis.fetch;
    const stats = hangingFetch();
    const hang = globalThis.fetch;
    vi.stubGlobal("fetch", (url: string | URL, init?: RequestInit) =>
      String(url).includes("$metadata") ? answer(url, init) : hang(url, init),
    );
    const controller = new AbortController();
    const call = client.callTool(
      { name: "read.counterparty.find_counterparty", arguments: { query: "Ромашка" } },
      { signal: controller.signal },
    );
    await vi.waitFor(() => expect(stats.calls).toBe(1));
    controller.abort();
    await expect(call).rejects.toThrow();
    await vi.waitFor(() => expect(stats.aborted).toBe(1));
    await sleep(1_200); // пауза первого повтора — 1 с
    expect(stats.calls).toBe(1);
    await client.close();
  });

  it("лимит времени: ошибка [deadline] без частичных данных и без новых запросов", async () => {
    const { client } = await connect({ ODATA_TOOL_DEADLINE_MS: "50" });
    const stats = hangingFetch();
    const r = (await client.callTool({ name: "read.system.health_check", arguments: {} })) as {
      isError?: boolean;
      structuredContent?: unknown;
      content: { text: string }[];
    };
    expect(r.isError).toBe(true);
    expect(r.structuredContent).toBeUndefined();
    expect(r.content[0]!.text).toContain("[deadline]");
    expect(r.content[0]!.text).toContain("ODATA_TOOL_DEADLINE_MS");
    await sleep(1_200);
    expect(stats.calls).toBe(1);
    await client.close();
  });

  it("лимит 0: сигнала нет, контекст вызова не создаётся", async () => {
    await withToolCall({ deadlineMs: 0 }, async () => {
      expect(callAbortError()).toBeUndefined();
    });
  });

  it("таймаут запроса без отмены по-прежнему повторяется", async () => {
    const f = fake1C({ READ_ONLY: "true", ODATA_RETRIES: "1", ODATA_TIMEOUT_MS: "20" });
    const stats = hangingFetch();
    await expect(f.conn.client.getCollection("Catalog_Контрагенты")).rejects.toMatchObject({
      kind: "timeout",
    });
    expect(stats.calls).toBe(2);
  });

  it("запись не получает сигнал вызова", async () => {
    const f = fake1C({ READ_ONLY: "false", ODATA_DB_NAME: "x", ODATA_WRITABLE: "true" });
    let writeSignal: AbortSignal | undefined;
    let readSignal: AbortSignal | undefined;
    vi.stubGlobal(
      "fetch",
      vi.fn(async (_u: string | URL, init?: { method?: string; signal?: AbortSignal }) => {
        if (init?.method === "PATCH") writeSignal = init.signal;
        else readSignal = init?.signal;
        return new Response("{}", { status: 200, headers: { "Content-Type": "application/json" } });
      }),
    );
    const controller = new AbortController();
    await withToolCall({ signal: controller.signal, deadlineMs: 0 }, async () => {
      await f.conn.client.getEntity("Catalog_Контрагенты(guid'1')");
      const write = f.conn.client.patch("Catalog_Контрагенты(guid'1')", {});
      controller.abort();
      await write;
    });
    expect(readSignal).toBeDefined();
    expect(writeSignal?.aborted).toBe(false);
  });
});

describe("общие загрузки и отмена", () => {
  it("отмена начавшего загрузку кэша не задевает соседний вызов", async () => {
    const owner = {};
    let finish!: (v: string) => void;
    const load = vi.fn(() => new Promise<string>((r) => (finish = r)));
    const controller = new AbortController();
    const first = withToolCall({ signal: controller.signal, deadlineMs: 0 }, () => cached(owner, "k", load));
    const second = withToolCall({ signal: new AbortController().signal, deadlineMs: 0 }, () =>
      cached(owner, "k", load),
    );
    controller.abort();
    await expect(first).rejects.toMatchObject({ kind: "cancelled" });
    finish("данные");
    await expect(second).resolves.toBe("данные");
    expect(load).toHaveBeenCalledTimes(1);
    await expect(cached(owner, "k", load)).resolves.toBe("данные");
  });

  it("загрузка схемы переживает отмену начавшего её вызова", async () => {
    const f = fake1C({ READ_ONLY: "true" });
    const stats = hangingFetch();
    const controller = new AbortController();
    const first = withToolCall({ signal: controller.signal, deadlineMs: 0 }, () => f.conn.getMetadata());
    await vi.waitFor(() => expect(stats.calls).toBe(1));
    controller.abort();
    await expect(first).rejects.toMatchObject({ kind: "cancelled" });
    expect(stats.aborted).toBe(0);
  });
});

describe("ожидание места под запрос", () => {
  it("отмена убирает из очереди, следующий получает место", async () => {
    const limiter = new RequestLimiter(1, 10_000);
    const release = await limiter.acquire();
    const controller = new AbortController();
    const cancelled = limiter.acquire(controller.signal);
    const next = limiter.acquire();
    controller.abort();
    await expect(cancelled).rejects.toMatchObject({ name: "AbortError" });
    release();
    const releaseNext = await next;
    expect(limiter.inFlight).toBe(1);
    releaseNext();
    expect(limiter.inFlight).toBe(0);
  });

  it("вызов, прерванный в очереди, получает свою ошибку, а не busy", async () => {
    const f = fake1C({ READ_ONLY: "true", ODATA_MAX_CONCURRENT_REQUESTS: "1" });
    hangingFetch();
    const blocking = f.conn.client.getCollection("Catalog_A").catch((e: unknown) => e);
    const controller = new AbortController();
    const waiting = withToolCall({ signal: controller.signal, deadlineMs: 0 }, () =>
      f.conn.client.getCollection("Catalog_B"),
    );
    await sleep(10);
    controller.abort();
    const err = await waiting.catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ODataError);
    expect((err as ODataError).kind).toBe("cancelled");
    void blocking;
  });
});
