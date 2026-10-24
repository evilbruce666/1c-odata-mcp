import { describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";

// Простое проведение (без expectedDataVersion): успех — только по состоянию документа после действия.
const SET = "Document_РеализацияТоваровУслуг";
const ref = "33333333-3333-4333-8333-333333333333";

type Doc = { Posted: boolean; DataVersion: string };

function tool(opts: { before: Doc; after: Doc | Error; response?: unknown }) {
  let acted = false;
  const action = vi.fn(async () => {
    acted = true;
    return opts.response;
  });
  const getEntity = vi.fn(async (path: string) => {
    if (!path.includes("$select=Posted,DataVersion")) return {}; // проверка договора: полей нет в схеме
    if (!acted) return opts.before;
    if (opts.after instanceof Error) throw opts.after;
    return opts.after;
  });
  const connection = {
    cfg: { name: "default", writable: true },
    available: async () => new Set([SET]),
    getMetadata: async () => ({ entities: new Map([[SET, { properties: [] }]]) }),
    client: { action, getEntity, getCollection: async () => ({ value: [] }) },
  };
  const t = (
    createServer({ db: () => connection } as never) as unknown as {
      _registeredTools: Record<
        string,
        { handler: (a: Record<string, unknown>, e: unknown) => Promise<CallToolResult> }
      >;
    }
  )._registeredTools["write.document.post_document"]!;
  return {
    action,
    run: (post = true) => t.handler({ database: "default", entitySet: SET, ref, post, confirm: true }, {}),
  };
}

const text = (r: CallToolResult) => JSON.stringify(r.content);

describe("post_document: успех по Posted до и после действия", () => {
  it("1С ответила 200, документ проведён — успех с verified", async () => {
    const { run } = tool({
      before: { Posted: false, DataVersion: "1" },
      after: { Posted: true, DataVersion: "2" },
    });
    const r = await run();
    expect(r.isError).toBeFalsy();
    expect(r.structuredContent).toMatchObject({ done: true, verified: true, action: "Post" });
  });

  it("1С ответила 200, но Posted=false — ошибка с телом ответа 1С", async () => {
    const { run } = tool({
      before: { Posted: false, DataVersion: "1" },
      after: { Posted: false, DataVersion: "1" },
      response: { message: "Не заполнен склад" },
    });
    const r = await run();
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("не проведён");
    expect(text(r)).toContain("Не заполнен склад");
  });

  it("повторное проведение без смены версии — не успех", async () => {
    const { run } = tool({
      before: { Posted: true, DataVersion: "5" },
      after: { Posted: true, DataVersion: "5" },
    });
    const r = await run();
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("версия не изменилась");
  });

  it("повторное проведение со сменой версии — успех", async () => {
    const { run } = tool({
      before: { Posted: true, DataVersion: "5" },
      after: { Posted: true, DataVersion: "6" },
    });
    expect((await run()).isError).toBeFalsy();
  });

  it("отмена проведения: Posted остался true — ошибка; стал false — успех", async () => {
    const stuck = tool({
      before: { Posted: true, DataVersion: "5" },
      after: { Posted: true, DataVersion: "6" },
    });
    const r = await stuck.run(false);
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("остался проведённым");
    const done = tool({
      before: { Posted: true, DataVersion: "5" },
      after: { Posted: false, DataVersion: "6" },
    });
    expect((await done.run(false)).isError).toBeFalsy();
  });

  it("перечитать после действия не удалось — ошибка «проверьте в 1С», без повтора действия", async () => {
    const { run, action } = tool({
      before: { Posted: false, DataVersion: "1" },
      after: new Error("сеть"),
    });
    const r = await run();
    expect(r.isError).toBe(true);
    expect(text(r)).toContain("Проверьте в 1С");
    expect(action).toHaveBeenCalledOnce();
  });
});
