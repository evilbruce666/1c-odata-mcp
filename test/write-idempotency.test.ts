import { describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { createServer } from "../src/mcp/server.js";
import { CATALOGS } from "../src/config/mapping.js";
import { currentWriteOperationId, currentWriteRequestHash } from "../src/odata/write-operation-context.js";

describe("create tool operationId flow", () => {
  it("returns a preview token and requires/reuses it for confirmation", async () => {
    const prepareCreate = vi.fn(async () => undefined);
    const create = vi.fn(async () => {
      expect(currentWriteOperationId()).toMatch(/^[0-9a-f-]{36}$/i);
      expect(currentWriteRequestHash()).toMatch(/^[0-9a-f]{64}$/);
      return { Ref_Key: "ref-1", Code: "00001", Description: "Example" };
    });
    const connection = {
      cfg: { name: "default", writable: true },
      behavior: { writeOperationMarker: false },
      available: async () => new Set(CATALOGS.counterparties),
      client: { prepareCreate, create },
    };
    const server = createServer({ db: () => connection } as never) as unknown as {
      _registeredTools: Record<
        string,
        {
          inputSchema: { safeParse: (input: unknown) => { success: boolean } };
          outputSchema: { safeParse: (input: unknown) => { success: boolean } };
          handler: (args: Record<string, unknown>, extra: Record<string, unknown>) => Promise<CallToolResult>;
        }
      >;
    };
    const tool = server._registeredTools["write.counterparty.create_counterparty"]!;
    const previewArgs = { database: "default", name: "Example", confirm: false };
    const preview = await tool.handler(previewArgs, {});
    const operationId = (preview.structuredContent as Record<string, unknown>).operationId;

    expect(typeof operationId).toBe("string");
    expect(prepareCreate).toHaveBeenCalledOnce();
    expect(create).not.toHaveBeenCalled();
    expect(tool.inputSchema.safeParse({ ...previewArgs, operationId }).success).toBe(true);
    expect(tool.outputSchema.safeParse(preview.structuredContent).success).toBe(true);

    const missingToken = await tool.handler({ ...previewArgs, confirm: true }, {});
    expect(missingToken.isError).toBe(true);
    expect(create).not.toHaveBeenCalled();

    const confirmed = await tool.handler({ ...previewArgs, confirm: true, operationId }, {});
    expect(create).toHaveBeenCalledOnce();
    expect(confirmed.structuredContent).toMatchObject({ created: true, operationId });
  });
});

type Tool = {
  handler: (args: Record<string, unknown>, extra: Record<string, unknown>) => Promise<CallToolResult>;
};
const toolsOf = (connection: unknown): Record<string, Tool> =>
  (createServer({ db: () => connection } as never) as unknown as { _registeredTools: Record<string, Tool> })
    ._registeredTools;

describe("operation marker and write.operation.status", () => {
  const meta = (withComment: boolean) => ({
    entities: new Map([
      [CATALOGS.counterparties[0]!, { properties: withComment ? [{ name: "Комментарий" }] : [] }],
    ]),
  });

  const counterpartyPreview = async (enabled: boolean, withComment: boolean) => {
    const connection = {
      cfg: { name: "default", writable: true },
      behavior: { writeOperationMarker: enabled },
      getMetadata: async () => meta(withComment),
      available: async () => new Set(CATALOGS.counterparties),
      client: { prepareCreate: vi.fn(), create: vi.fn() },
    };
    const res = await toolsOf(connection)["write.counterparty.create_counterparty"]!.handler(
      { database: "default", name: "Example", confirm: false },
      {},
    );
    return res.structuredContent as { operationId: string; payload: Record<string, unknown> };
  };

  it("appends the operation marker to the comment when the entity supports it", async () => {
    const sc = await counterpartyPreview(true, true);
    expect(sc.payload["Комментарий"]).toBe(`[op:${sc.operationId}]`);
  });

  it("does not touch the comment when the marker is disabled or unsupported", async () => {
    expect((await counterpartyPreview(false, true)).payload["Комментарий"]).toBeUndefined();
    expect((await counterpartyPreview(true, false)).payload["Комментарий"]).toBeUndefined();
  });

  it("reconciles an unknown operation when the marked object is found, searching without brackets", async () => {
    const id = "22222222-2222-4222-8222-222222222222";
    const reconcileOperation = vi.fn(async () => undefined);
    const getCollection = vi.fn(async () => ({ value: [{ Ref_Key: "ref-9", Number: "0000-000009" }] }));
    const connection = {
      cfg: { name: "default" },
      getMetadata: async () => meta(true),
      client: {
        operationEntry: async () => ({
          state: "outcome_unknown",
          entitySet: CATALOGS.counterparties[0]!,
          updatedAt: new Date().toISOString(),
        }),
        getCollection,
        reconcileOperation,
      },
    };
    const res = await toolsOf(connection)["write.operation.status"]!.handler(
      { database: "default", operationId: id },
      {},
    );
    expect(res.structuredContent).toMatchObject({ status: "found_reconciled", ref: "ref-9" });
    expect(reconcileOperation).toHaveBeenCalledWith(id, expect.objectContaining({ Ref_Key: "ref-9" }));
    const url = decodeURIComponent((getCollection.mock.calls[0] as unknown as [string])[0]);
    expect(url).toContain(`substringof('op:${id}',Комментарий)`);
    expect(url).not.toContain("[");
  });

  it("reports not_found / unverifiable / not_in_journal without reconciling", async () => {
    const reconcileOperation = vi.fn();
    const mk = (entry: unknown, withComment: boolean, value: unknown[]) => ({
      cfg: { name: "default" },
      getMetadata: async () => meta(withComment),
      client: {
        operationEntry: async () => entry,
        getCollection: async () => ({ value }),
        reconcileOperation,
      },
    });
    const unknown = {
      state: "executing",
      entitySet: CATALOGS.counterparties[0]!,
      updatedAt: new Date(Date.now() - 10 * 60_000).toISOString(),
    };
    const id = "33333333-3333-4333-8333-333333333333";
    const status = async (conn: unknown) =>
      (await toolsOf(conn)["write.operation.status"]!.handler({ database: "default", operationId: id }, {}))
        .structuredContent;
    expect(await status(mk(unknown, true, []))).toMatchObject({ status: "not_found" });
    expect(await status(mk(unknown, false, []))).toMatchObject({ status: "unverifiable" });
    expect(await status(mk(undefined, true, []))).toMatchObject({ status: "not_in_journal" });
    expect(reconcileOperation).not.toHaveBeenCalled();
  });
});
