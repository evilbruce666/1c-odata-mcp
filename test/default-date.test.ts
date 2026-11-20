import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

/**
 * Дата по умолчанию — момент предпросмотра. Журнал сравнивает у дат только день, поэтому «сейчас» при подтверждении
 * через несколько секунд не мешало, а после полуночи давало другой день: подтверждение отказывало («не совпадают с
 * dry-run»). Прежние тесты не ловили: предпросмотр и подтверждение шли в одну секунду.
 */

const id = "11111111-1111-4111-8111-111111111111";
const line = { nomenclatureRef: id, quantity: 1, price: 100, vatRate: "БезНДС" };

afterEach(() => vi.useRealTimers());

describe("дата по умолчанию в Бухгалтерии — момент предпросмотра", () => {
  it.each([
    ["write.catalog.create_contract", { counterpartyRef: id, kind: "СПокупателем" }, "Дата"],
    ["write.sales.create_invoice", { counterpartyRef: id, lines: [line] }, "Date"],
    ["write.sales.create_shipment", { counterpartyRef: id, lines: [line] }, "Date"],
  ] as const)(
    "%s: подтверждение после полуночи проходит, дата — как в предпросмотре",
    async (name, args, field) => {
      vi.useFakeTimers({ toFake: ["Date"] });
      vi.setSystemTime(new Date(2026, 10, 20, 23, 59, 58));
      const posts: Record<string, unknown>[] = [];
      const fake = fake1C({
        READ_ONLY: "false",
        ODATA_WRITABLE: "true",
        ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-date-")),
      });
      const inner = globalThis.fetch;
      vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
        if ((init?.method ?? "GET").toUpperCase() === "POST")
          posts.push(JSON.parse(String(init?.body)) as Record<string, unknown>);
        return inner(input, init);
      });
      const call = async (a: Record<string, unknown>) => {
        const [x, y] = InMemoryTransport.createLinkedPair();
        await createServer(fake.ctx).connect(x);
        const client = new Client({ name: "t", version: "1" });
        await client.connect(y);
        return (await client.callTool({ name, arguments: a })) as {
          isError?: boolean;
          structuredContent: Record<string, unknown>;
          content: { text: string }[];
        };
      };
      const preview = await call(args);
      expect(preview.isError, preview.content[0]?.text).toBeFalsy();
      expect((preview.structuredContent["payload"] as Record<string, unknown>)[field]).toBe(
        "2026-11-20T23:59:58",
      );
      vi.setSystemTime(new Date(2026, 10, 21, 0, 0, 5));
      const done = await call({
        ...args,
        confirm: true,
        operationId: preview.structuredContent["operationId"],
      });
      expect(done.isError, done.content[0]?.text.slice(0, 300)).toBeFalsy();
      expect(posts.at(-1)?.[field]).toBe("2026-11-20T23:59:58");
      // Явная дата — как была: полночь заданного дня.
      vi.setSystemTime(new Date(2026, 10, 21, 10, 0, 0));
      const explicit = await call({ ...args, date: "2026-11-15" });
      expect((explicit.structuredContent["payload"] as Record<string, unknown>)[field]).toBe(
        "2026-11-15T00:00:00",
      );
    },
  );
});
