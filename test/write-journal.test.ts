import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import { afterEach, describe, expect, it } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import { fingerprintWriteInput } from "../src/odata/write-operation-context.js";
import { WriteOperationJournal } from "../src/odata/write-journal.js";

const database = "buh";
const baseUrl = "https://1c.example/odata/standard.odata/";
const entitySet = "Document_ПлатежноеПоручение";
const payload = { Date: "2026-09-29T00:00:00", СуммаДокумента: 125 };
const requestHash = "request-hash";

describe("WriteOperationJournal", () => {
  const roots: string[] = [];

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
  });

  async function createJournal(): Promise<WriteOperationJournal> {
    const root = await mkdtemp(join(tmpdir(), "1c-odata-write-journal-"));
    roots.push(root);
    return new WriteOperationJournal(root, database, baseUrl);
  }

  it("reuses the stored result for a repeated confirmation without sending another POST", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    let sends = 0;
    await journal.prepare(operationId, entitySet, payload, requestHash);

    const created = await journal.execute(operationId, entitySet, payload, requestHash, async () => {
      sends += 1;
      return { Ref_Key: "ref-1", Code: "00001", Description: "do not persist" };
    });
    const restartedJournal = new WriteOperationJournal(roots[0]!, database, baseUrl);
    const replayed = await restartedJournal.execute(
      operationId,
      entitySet,
      payload,
      requestHash,
      async () => {
        sends += 1;
        throw new Error("duplicate POST must not be sent");
      },
    );

    expect(sends).toBe(1);
    expect(created.Ref_Key).toBe("ref-1");
    expect(replayed).toMatchObject({ Ref_Key: "ref-1", Code: "00001", _operation_replayed: true });
    expect(replayed.Description).toBeUndefined();
  });

  it("blocks a retry when the first POST ended with an unknown outcome", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    let sends = 0;
    await journal.prepare(operationId, entitySet, payload, requestHash);

    await expect(
      journal.execute(operationId, entitySet, payload, requestHash, async () => {
        sends += 1;
        throw new ODataError({ kind: "timeout", message: "timeout after dispatch" });
      }),
    ).rejects.toThrow(/результат записи/i);

    const restartedJournal = new WriteOperationJournal(roots[0]!, database, baseUrl);
    await expect(
      restartedJournal.execute(operationId, entitySet, payload, requestHash, async () => {
        sends += 1;
        return { Ref_Key: "duplicate" };
      }),
    ).rejects.toThrow(/неизвестен/i);
    expect(sends).toBe(1);
  });

  it("rejects reuse of a preview token for different data", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    await journal.prepare(operationId, entitySet, payload, requestHash);

    await expect(
      journal.execute(operationId, entitySet, { ...payload, СуммаДокумента: 250 }, requestHash, async () => ({
        Ref_Key: "must-not-create",
      })),
    ).rejects.toThrow(/не совпадают с dry-run/i);
  });

  it("treats object key order as the same preview payload", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    await journal.prepare(
      operationId,
      entitySet,
      { Date: payload.Date, СуммаДокумента: payload.СуммаДокумента },
      requestHash,
    );

    const created = await journal.execute(
      operationId,
      entitySet,
      { СуммаДокумента: payload.СуммаДокумента, Date: payload.Date },
      requestHash,
      async () => ({ Ref_Key: "ref-1" }),
    );

    expect(created.Ref_Key).toBe("ref-1");
  });

  it("accepts a seconds-only timestamp change while rejecting changed tool arguments", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    const args = { database, amount: 125, confirm: false };
    const sameOperationArgs = { amount: 125, database, confirm: true, operationId };
    const fingerprint = fingerprintWriteInput("write.money.create_payment", args);
    const confirmationFingerprint = fingerprintWriteInput("write.money.create_payment", sameOperationArgs);
    await journal.prepare(operationId, entitySet, { ...payload, Date: "2026-09-29T11:30:01" }, fingerprint);

    const created = await journal.execute(
      operationId,
      entitySet,
      { ...payload, Date: "2026-09-29T11:30:58" },
      confirmationFingerprint,
      async () => ({ Ref_Key: "ref-1" }),
    );

    expect(created.Ref_Key).toBe("ref-1");
    expect(
      fingerprintWriteInput("write.money.create_payment", { ...sameOperationArgs, amount: 250 }),
    ).not.toBe(fingerprint);
  });

  it("does not repeat a request that received a definite 400 response", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    let sends = 0;
    await journal.prepare(operationId, entitySet, payload, requestHash);

    await expect(
      journal.execute(operationId, entitySet, payload, requestHash, async () => {
        sends += 1;
        throw new ODataError({ kind: "bad_request", message: "invalid payload", status: 400 });
      }),
    ).rejects.toThrow("invalid payload");
    await expect(
      journal.execute(operationId, entitySet, payload, requestHash, async () => {
        sends += 1;
        return { Ref_Key: "must-not-create" };
      }),
    ).rejects.toThrow(/ранее была отклонена/i);
    expect(sends).toBe(1);
  });

  it("atomically allows only one concurrent POST for the same operationId", async () => {
    const journal = await createJournal();
    const operationId = randomUUID();
    let sends = 0;
    let releaseSend: (() => void) | undefined;
    let markSendStarted: (() => void) | undefined;
    const sendStarted = new Promise<void>((resolve) => {
      markSendStarted = resolve;
    });
    const holdSend = new Promise<void>((resolve) => {
      releaseSend = resolve;
    });
    await journal.prepare(operationId, entitySet, payload, requestHash);

    const first = journal.execute(operationId, entitySet, payload, requestHash, async () => {
      sends += 1;
      markSendStarted?.();
      await holdSend;
      return { Ref_Key: "ref-1" };
    });
    await sendStarted;
    const second = journal.execute(operationId, entitySet, payload, requestHash, async () => {
      sends += 1;
      return { Ref_Key: "duplicate" };
    });
    await expect(second).rejects.toThrow(/неизвестен/i);
    releaseSend?.();
    const created = await first;

    expect(sends).toBe(1);
    expect(created.Ref_Key).toBe("ref-1");
  });
});
