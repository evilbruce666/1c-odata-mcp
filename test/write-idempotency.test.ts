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
