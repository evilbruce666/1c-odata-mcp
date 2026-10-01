import { describe, expect, it, vi } from "vitest";
import { InputError } from "../src/errors.js";
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
      client: { prepareCreate, create, operationSettled: async () => undefined },
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

describe("line operations (add/remove document line) and journal replay", () => {
  const DOC = "Document_СчетНаОплатуПокупателю";
  const ref = "44444444-4444-4444-8444-444444444444";
  const nom = "55555555-5555-4555-8555-555555555555";
  const row = (n: number) => ({ LineNumber: n, Номенклатура_Key: nom, Количество: 1, Цена: 10, Сумма: 10 });
  const docWith = (lines: number) => ({
    Ref_Key: ref,
    Posted: false,
    Организация_Key: "org",
    Товары: Array.from({ length: lines }, (_, i) => row(i + 1)),
  });

  it("add_document_line: preview gives operationId, confirm goes through patchOnce with a line-count check", async () => {
    const patchOnce = vi.fn(async () => ({ Ref_Key: ref }));
    const prepareCreate = vi.fn(async () => undefined);
    const connection = {
      cfg: { name: "default", writable: true },
      behavior: { writeOperationMarker: true },
      available: async () => new Set([DOC]),
      getMetadata: async () => ({ entities: new Map() }),
      client: {
        prepareCreate,
        patchOnce,
        operationSettled: async () => undefined,
        getEntity: async () => docWith(2),
        getCollection: async () => ({ value: [] }),
      },
    };
    const tool = toolsOf(connection)["write.document.add_document_line"]!;
    const args = {
      database: "default",
      entitySet: DOC,
      ref,
      line: { nomenclatureRef: nom, quantity: 1, price: 10, vatRate: "БезНДС" },
    };
    const preview = await tool.handler({ ...args, confirm: false }, {});
    expect(preview.isError, JSON.stringify(preview.content)).toBeFalsy();
    const operationId = (preview.structuredContent as { operationId: string }).operationId;
    expect(operationId).toMatch(/^[0-9a-f-]{36}$/);
    expect(prepareCreate).toHaveBeenCalledWith(DOC, { ref }, { withRefKey: false });
    const confirmed = await tool.handler({ ...args, confirm: true, operationId }, {});
    expect(confirmed.isError).toBeFalsy();
    expect(patchOnce).toHaveBeenCalledOnce();
    const call = patchOnce.mock.calls[0] as unknown as [
      string,
      unknown,
      string,
      unknown,
      Record<string, unknown>,
    ];
    expect(call[3]).toEqual({ ref });
    expect(call[4]).toMatchObject({ kind: "lineCount", ref, before: 2, expected: 3 });
  });

  it("a confirmed operation is answered from the journal without re-reading 1C or calling the tool", async () => {
    const getEntity = vi.fn();
    const patchOnce = vi.fn();
    const connection = {
      cfg: { name: "default", writable: true },
      behavior: { writeOperationMarker: true },
      available: async () => new Set([DOC]),
      client: {
        operationSettled: async () => ({ entitySet: DOC, result: { Ref_Key: ref } }),
        getEntity,
        patchOnce,
      },
    };
    const tool = toolsOf(connection)["write.document.remove_document_line"]!;
    const res = await tool.handler(
      {
        database: "default",
        entitySet: DOC,
        ref,
        lineNumber: 1,
        confirm: true,
        operationId: "66666666-6666-4666-8666-666666666666",
      },
      {},
    );
    expect(res.structuredContent).toMatchObject({ updated: true, replayed: true, ref, entitySet: DOC });
    expect(getEntity).not.toHaveBeenCalled();
    expect(patchOnce).not.toHaveBeenCalled();
  });

  it("an unknown outcome blocks the retry before the tool runs", async () => {
    const getEntity = vi.fn();
    const connection = {
      cfg: { name: "default", writable: true },
      client: {
        operationSettled: async () => {
          throw new InputError("Результат записи неизвестен");
        },
        getEntity,
      },
    };
    const res = await toolsOf(connection)["write.document.remove_document_line"]!.handler(
      {
        database: "default",
        entitySet: DOC,
        ref,
        lineNumber: 1,
        confirm: true,
        operationId: "77777777-7777-4777-8777-777777777777",
      },
      {},
    );
    expect(res.isError).toBe(true);
    expect(getEntity).not.toHaveBeenCalled();
  });

  describe("write.operation.status by line count", () => {
    const status = async (lines: number, ageMs: number) => {
      const reconcileOperation = vi.fn(async () => undefined);
      const markOperationNotApplied = vi.fn(async () => undefined);
      const connection = {
        cfg: { name: "default" },
        client: {
          operationEntry: async () => ({
            state: "outcome_unknown",
            entitySet: DOC,
            updatedAt: new Date(Date.now() - ageMs).toISOString(),
            check: { kind: "lineCount", ref, section: "Товары", before: 2, expected: 3 },
          }),
          getEntity: async () => docWith(lines),
          reconcileOperation,
          markOperationNotApplied,
        },
      };
      const res = await toolsOf(connection)["write.operation.status"]!.handler(
        { database: "default", operationId: "88888888-8888-4888-8888-888888888888" },
        {},
      );
      return { sc: res.structuredContent as { status: string }, reconcileOperation, markOperationNotApplied };
    };

    it("expected count → found_reconciled", async () => {
      const r = await status(3, 1000);
      expect(r.sc.status).toBe("found_reconciled");
      expect(r.reconcileOperation).toHaveBeenCalledOnce();
    });
    it("unchanged after 2+ minutes → not_applied (journal closed)", async () => {
      const r = await status(2, 5 * 60_000);
      expect(r.sc.status).toBe("not_applied");
      expect(r.markOperationNotApplied).toHaveBeenCalledOnce();
    });
    it("unchanged right after the attempt → not_found, journal untouched", async () => {
      const r = await status(2, 10_000);
      expect(r.sc.status).toBe("not_found");
      expect(r.markOperationNotApplied).not.toHaveBeenCalled();
    });
    it("some other count → unverifiable", async () => {
      const r = await status(5, 5 * 60_000);
      expect(r.sc.status).toBe("unverifiable");
      expect(r.reconcileOperation).not.toHaveBeenCalled();
      expect(r.markOperationNotApplied).not.toHaveBeenCalled();
    });
  });
});

describe("write.operation.status по назначенному Ref_Key", () => {
  const refKey = "aaaaaaaa-bbbb-1ccc-8ddd-eeeeeeeeeeee";
  const status = async (exists: boolean, ageMs: number) => {
    const reconcileOperation = vi.fn(async () => undefined);
    const markOperationNotApplied = vi.fn(async () => undefined);
    const { ODataError } = await import("../src/odata/errors.js");
    const connection = {
      cfg: { name: "default" },
      client: {
        operationEntry: async () => ({
          state: "outcome_unknown",
          entitySet: "Catalog_Контрагенты",
          refKey,
          updatedAt: new Date(Date.now() - ageMs).toISOString(),
        }),
        getEntity: async () => {
          if (exists) return { Ref_Key: refKey };
          throw new ODataError({ kind: "not_found", message: "нет" });
        },
        reconcileOperation,
        markOperationNotApplied,
      },
    };
    const res = await toolsOf(connection)["write.operation.status"]!.handler(
      { database: "default", operationId: "99999999-9999-4999-8999-999999999999" },
      {},
    );
    return {
      sc: res.structuredContent as { status: string; ref?: string },
      reconcileOperation,
      markOperationNotApplied,
    };
  };
  it("объект есть → found_reconciled (справочник тоже, без «Комментария»)", async () => {
    const r = await status(true, 1000);
    expect(r.sc).toMatchObject({ status: "found_reconciled", ref: refKey });
    expect(r.reconcileOperation).toHaveBeenCalledOnce();
  });
  it("объекта нет 2+ минуты → not_applied", async () => {
    const r = await status(false, 5 * 60_000);
    expect(r.sc.status).toBe("not_applied");
    expect(r.markOperationNotApplied).toHaveBeenCalledOnce();
  });
  it("объекта нет сразу после попытки → not_found без закрытия", async () => {
    const r = await status(false, 5_000);
    expect(r.sc.status).toBe("not_found");
    expect(r.markOperationNotApplied).not.toHaveBeenCalled();
  });
});
