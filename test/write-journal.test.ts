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

describe("ODataClient write guard vs journal", () => {
  it("does not touch the journal on preview or create when writes are disabled", async () => {
    const { mkdtemp, readdir } = await import("node:fs/promises");
    const { tmpdir } = await import("node:os");
    const { join } = await import("node:path");
    const { ODataClient } = await import("../src/odata/client.js");
    const { withWriteOperation } = await import("../src/odata/write-operation-context.js");
    const dir = await mkdtemp(join(tmpdir(), "journal-ro-"));
    const client = new ODataClient(
      {
        name: "ro",
        baseUrl: "http://localhost/odata/",
        username: "u",
        password: "p",
        writable: false,
      } as never,
      { readOnly: true, writeJournalDir: dir, retries: 0, timeoutMs: 1000 } as never,
    );
    const id = "11111111-1111-4111-8111-111111111111";
    await withWriteOperation(id, "h", async () => {
      await client.prepareCreate("Catalog_X", { a: 1 });
      await expect(client.create("Catalog_X", { a: 1 })).rejects.toThrow();
    });
    expect(await readdir(dir)).toEqual([]);
  });
});

describe("journal payload hash", () => {
  it("ignores time-of-day drift in any date field (Date and Дата)", async () => {
    const dir = await mkdtemp(join(tmpdir(), "journal-date-"));
    const j = new WriteOperationJournal(dir, "db", "http://x/");
    const id = randomUUID();
    await j.prepare(id, "Catalog_Договоры", { Дата: "2026-10-01T10:00:00", Номер: "1" }, "h");
    // тот же день, другое время — не должно считаться расхождением
    await expect(
      j.execute(id, "Catalog_Договоры", { Дата: "2026-10-01T10:05:09", Номер: "1" }, "h", async () => ({
        Ref_Key: "r",
      })),
    ).resolves.toMatchObject({ Ref_Key: "r" });
    // другой день — расхождение
    const id2 = randomUUID();
    await j.prepare(id2, "Catalog_Договоры", { Дата: "2026-10-01T10:00:00" }, "h");
    await expect(
      j.execute(id2, "Catalog_Договоры", { Дата: "2026-10-02T10:00:00" }, "h", async () => ({})),
    ).rejects.toThrow(/не совпадают/);
  });
});

describe("journal pruning", () => {
  it("removes stale tmp/lock files and ancient records, keeps recent and uncertain ones", async () => {
    const { writeFile, readdir, utimes } = await import("node:fs/promises");
    const root = await mkdtemp(join(tmpdir(), "journal-prune-"));
    const j = new WriteOperationJournal(root, "db", "http://x/");
    // Первый prepare создаёт каталог журнала и запускает чистку.
    await j.prepare(randomUUID(), "Catalog_X", { a: 0 }, "h");
    const dir = join(root, (await readdir(root))[0]!);
    const day = 24 * 60 * 60 * 1000;
    const old = (ms: number) => new Date(Date.now() - ms).toISOString();
    const entry = (id: string, state: string, updatedAt: string) =>
      JSON.stringify({
        version: 1,
        operationId: id,
        database: "db",
        entitySet: "Catalog_X",
        state,
        updatedAt,
      });
    const age = async (name: string, ms: number) => {
      const t = new Date(Date.now() - ms);
      await utimes(join(dir, name), t, t);
    };
    const ids = Array.from({ length: 6 }, () => randomUUID());
    await writeFile(join(dir, `${ids[0]}.json`), entry(ids[0]!, "succeeded", old(100 * day))); // старая → удалить
    await writeFile(join(dir, `${ids[1]}.json`), entry(ids[1]!, "outcome_unknown", old(100 * day))); // неизвестная 100д → оставить
    await writeFile(join(dir, `${ids[2]}.json`), entry(ids[2]!, "outcome_unknown", old(400 * day))); // неизвестная >года → удалить
    await writeFile(join(dir, `${ids[3]}.json`), entry(ids[3]!, "executing", old(1000))); // свежая → оставить
    await writeFile(join(dir, `${ids[3]}.lock`), "");
    await writeFile(join(dir, `${ids[4]}.lock`), ""); // сирота без записи
    await age(`${ids[4]}.lock`, 2 * 60 * 60 * 1000);
    await writeFile(join(dir, `${ids[5]}.json.tmp`), "");
    await age(`${ids[5]}.json.tmp`, 2 * day); // зависший tmp
    // Новый экземпляр журнала → чистка запустится снова.
    await new WriteOperationJournal(root, "db", "http://x/").prepare(
      randomUUID(),
      "Catalog_X",
      { a: 1 },
      "h",
    );
    const left = new Set(await readdir(dir));
    expect(left.has(`${ids[0]}.json`)).toBe(false);
    expect(left.has(`${ids[1]}.json`)).toBe(true);
    expect(left.has(`${ids[2]}.json`)).toBe(false);
    expect(left.has(`${ids[3]}.json`)).toBe(true);
    expect(left.has(`${ids[3]}.lock`)).toBe(true);
    expect(left.has(`${ids[4]}.lock`)).toBe(false);
    expect(left.has(`${ids[5]}.json.tmp`)).toBe(false);
  });
});
