import { afterEach, describe, expect, it, vi } from "vitest";
import { METADATA_TTL_MS } from "../src/odata/client.js";
import { logger } from "../src/logger.js";
import { fake1C } from "./helpers/fake-1c.js";

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

/** fake1C подменяет fetch; поверх считаем обращения к $metadata. */
function countMetadata() {
  const answer = globalThis.fetch;
  const stats = { calls: 0 };
  vi.stubGlobal("fetch", (url: string | URL, init?: RequestInit) => {
    if (String(url).includes("$metadata")) stats.calls += 1;
    return answer(url, init);
  });
  return stats;
}

describe("срок жизни схемы", () => {
  it("свежая не перечитывается, устаревшая обновляется один раз, ответ не ждёт", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const f = fake1C({ READ_ONLY: "true" });
    const stats = countMetadata();
    const first = await f.conn.getMetadata();
    await f.conn.getMetadata();
    expect(stats.calls).toBe(1);
    vi.setSystemTime(Date.now() + METADATA_TTL_MS + 1);
    const [stale] = await Promise.all([f.conn.getMetadata(), f.conn.getMetadata()]);
    expect(stale).toBe(first);
    await vi.waitFor(() => expect(stats.calls).toBe(2));
    await vi.waitFor(async () => expect(await f.conn.getMetadata()).not.toBe(first));
    expect(stats.calls).toBe(2);
  });

  it("сбой обновления: прежняя схема и предупреждение, повтор не сразу", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    const warn = vi.spyOn(logger, "warn").mockImplementation(() => undefined);
    const f = fake1C({ READ_ONLY: "true", ODATA_RETRIES: "0" });
    const before = await f.conn.getMetadata();
    vi.setSystemTime(Date.now() + METADATA_TTL_MS + 1);
    const fetchMock = vi.fn(async () => new Response("{}", { status: 503 }));
    vi.stubGlobal("fetch", fetchMock);
    await expect(f.conn.getMetadata()).resolves.toBe(before);
    await vi.waitFor(() =>
      expect(warn.mock.calls.some((c) => String(c[1]).includes("обновить не удалось"))).toBe(true),
    );
    const failed = fetchMock.mock.calls.length;
    await expect(f.conn.getMetadata()).resolves.toBe(before);
    expect(fetchMock.mock.calls.length).toBe(failed);
  });
});
