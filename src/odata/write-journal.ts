import { createHash, randomUUID } from "node:crypto";
import { mkdir, open, readFile, readdir, rename, rm } from "node:fs/promises";
import { dirname, join } from "node:path";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "./errors.js";

const OPERATION_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const TERMINAL_RETENTION_MS = 90 * 24 * 60 * 60 * 1_000;
const PRUNE_INTERVAL_MS = 60 * 60 * 1_000;

type OperationState = "prepared" | "executing" | "succeeded" | "rejected" | "outcome_unknown";

interface JournalEntry {
  version: 1;
  operationId: string;
  database: string;
  entitySet: string;
  payloadHash: string;
  requestHash: string;
  state: OperationState;
  createdAt: string;
  updatedAt: string;
  result?: Record<string, string>;
  errorKind?: string;
}

/**
 * Durable, local idempotency journal for confirmed OData creates.
 * Only a payload hash and a small result reference are persisted, never the payload itself.
 */
export class WriteOperationJournal {
  private readonly directory: string;
  private lastPrunedAt = 0;

  constructor(rootDirectory: string, database: string, baseUrl: string) {
    const namespace = createHash("sha256").update(`${database}\n${baseUrl}`).digest("hex").slice(0, 24);
    this.directory = join(rootDirectory, namespace);
    this.database = database;
  }

  private readonly database: string;

  async prepare(
    operationId: string,
    entitySet: string,
    payload: Record<string, unknown>,
    requestHash: string,
  ): Promise<void> {
    this.validateOperationId(operationId);
    await mkdir(this.directory, { recursive: true, mode: 0o700 });
    await this.pruneExpired();

    const now = new Date().toISOString();
    const entry: JournalEntry = {
      version: 1,
      operationId,
      database: this.database,
      entitySet,
      payloadHash: this.payloadHash(entitySet, payload),
      requestHash,
      state: "prepared",
      createdAt: now,
      updatedAt: now,
    };

    try {
      await this.writeNew(this.recordPath(operationId), entry);
    } catch (error) {
      if (!isNodeError(error, "EEXIST")) throw error;
      const existing = await this.read(operationId);
      this.assertMatches(existing, entitySet, payload, requestHash);
    }
  }

  async execute<T extends ODataEntity>(
    operationId: string,
    entitySet: string,
    payload: Record<string, unknown>,
    requestHash: string,
    send: () => Promise<T>,
  ): Promise<T> {
    this.validateOperationId(operationId);
    let entry = await this.read(operationId);
    this.assertMatches(entry, entitySet, payload, requestHash);
    const existingResult = this.resultForExisting<T>(entry);
    if (existingResult !== undefined) return existingResult;

    const lockPath = this.lockPath(operationId);
    let lock: Awaited<ReturnType<typeof open>>;
    try {
      lock = await open(lockPath, "wx", 0o600);
    } catch (error) {
      if (!isNodeError(error, "EEXIST")) throw error;
      entry = await this.read(operationId);
      this.assertMatches(entry, entitySet, payload, requestHash);
      const racedResult = this.resultForExisting<T>(entry);
      if (racedResult !== undefined) return racedResult;
      throw this.unknownResult(operationId);
    }

    let requestStarted = false;
    try {
      entry = await this.read(operationId);
      this.assertMatches(entry, entitySet, payload, requestHash);
      const racedResult = this.resultForExisting<T>(entry);
      if (racedResult !== undefined) return racedResult;
      if (entry.state !== "prepared") throw this.unknownResult(operationId);

      entry = { ...entry, state: "executing", updatedAt: new Date().toISOString() };
      await this.replace(entry);
      requestStarted = true;

      let result: T;
      try {
        result = await send();
      } catch (cause) {
        const rejected = isDefinitiveRejection(cause);
        const failed: JournalEntry = {
          ...entry,
          state: rejected ? "rejected" : "outcome_unknown",
          updatedAt: new Date().toISOString(),
          ...(cause instanceof ODataError ? { errorKind: cause.kind } : {}),
        };
        try {
          await this.replace(failed);
        } catch (journalError) {
          throw this.unknownResult(operationId, journalError);
        }
        await rm(lockPath, { force: true }).catch(() => undefined);
        if (rejected) throw cause;
        throw this.unknownResult(operationId, cause);
      }

      const completed: JournalEntry = {
        ...entry,
        state: "succeeded",
        updatedAt: new Date().toISOString(),
        result: resultSummary(result),
      };
      try {
        await this.replace(completed);
      } catch (journalError) {
        // The remote write may have succeeded. Keep the lock and fail closed on future retries.
        throw this.unknownResult(operationId, journalError);
      }
      await rm(lockPath, { force: true }).catch(() => undefined);
      return result;
    } catch (error) {
      if (!requestStarted) await rm(lockPath, { force: true }).catch(() => undefined);
      throw error;
    } finally {
      await lock.close().catch(() => undefined);
    }
  }

  private resultForExisting<T extends ODataEntity>(entry: JournalEntry): T | undefined {
    if (entry.state === "succeeded") {
      return { ...(entry.result ?? {}), _operation_replayed: true } as unknown as T;
    }
    if (entry.state === "rejected") {
      throw new InputError(
        `Попытка записи operationId ${entry.operationId} ранее была отклонена. Исправьте данные, выполните новый dry-run и подтвердите новый operationId.`,
      );
    }
    if (entry.state === "executing" || entry.state === "outcome_unknown") {
      throw this.unknownResult(entry.operationId);
    }
    return undefined;
  }

  private assertMatches(
    entry: JournalEntry,
    entitySet: string,
    payload: Record<string, unknown>,
    requestHash: string,
  ): void {
    if (entry.database !== this.database) {
      throw new InputError("operationId создан для другой базы 1С.");
    }
    if (
      entry.entitySet !== entitySet ||
      entry.payloadHash !== this.payloadHash(entitySet, payload) ||
      entry.requestHash !== requestHash
    ) {
      throw new InputError(
        `Данные, объект или база не совпадают с dry-run для operationId ${entry.operationId}. Выполните новый предпросмотр.`,
      );
    }
  }

  private async read(operationId: string): Promise<JournalEntry> {
    let text: string;
    try {
      text = await readFile(this.recordPath(operationId), "utf8");
    } catch (error) {
      if (isNodeError(error, "ENOENT")) {
        throw new InputError(
          `operationId ${operationId} не найден в локальном журнале. Сначала выполните новый dry-run.`,
        );
      }
      throw error;
    }
    try {
      const entry = JSON.parse(text) as JournalEntry;
      if (entry.version !== 1 || entry.operationId !== operationId || !entry.state)
        throw new Error("invalid entry");
      return entry;
    } catch {
      throw new InputError(
        `Запись operationId ${operationId} в локальном журнале повреждена. Повторная отправка заблокирована; проверьте базу 1С вручную.`,
      );
    }
  }

  private async writeNew(path: string, entry: JournalEntry): Promise<void> {
    const handle = await open(path, "wx", 0o600);
    try {
      await handle.writeFile(JSON.stringify(entry), "utf8");
      await handle.sync();
    } catch (error) {
      await rm(path, { force: true }).catch(() => undefined);
      throw error;
    } finally {
      await handle.close();
    }
  }

  private async replace(entry: JournalEntry): Promise<void> {
    const path = this.recordPath(entry.operationId);
    const temporaryPath = `${path}.${randomUUID()}.tmp`;
    const handle = await open(temporaryPath, "wx", 0o600);
    try {
      await handle.writeFile(JSON.stringify(entry), "utf8");
      await handle.sync();
    } finally {
      await handle.close();
    }
    try {
      await rename(temporaryPath, path);
      const directoryHandle = await open(dirname(path), "r").catch(() => undefined);
      if (directoryHandle) {
        try {
          await directoryHandle.sync().catch(() => undefined);
        } finally {
          await directoryHandle.close();
        }
      }
    } catch (error) {
      await rm(temporaryPath, { force: true }).catch(() => undefined);
      throw error;
    }
  }

  private async pruneExpired(): Promise<void> {
    const now = Date.now();
    if (now - this.lastPrunedAt < PRUNE_INTERVAL_MS) return;
    this.lastPrunedAt = now;
    let files: string[];
    try {
      files = await readdir(this.directory);
    } catch (error) {
      if (isNodeError(error, "ENOENT")) return;
      throw error;
    }
    const cutoff = now - TERMINAL_RETENTION_MS;
    await Promise.all(
      files
        .filter((file) => file.endsWith(".json"))
        .map(async (file) => {
          const path = join(this.directory, file);
          try {
            const entry = JSON.parse(await readFile(path, "utf8")) as JournalEntry;
            if (
              (entry.state === "prepared" || entry.state === "succeeded" || entry.state === "rejected") &&
              Date.parse(entry.updatedAt) < cutoff
            ) {
              await rm(path, { force: true });
            }
          } catch {
            // Keep unreadable and uncertain entries; deleting one could make a retry unsafe.
          }
        }),
    );
  }

  private recordPath(operationId: string): string {
    return join(this.directory, `${operationId}.json`);
  }

  private lockPath(operationId: string): string {
    return join(this.directory, `${operationId}.lock`);
  }

  private payloadHash(entitySet: string, payload: Record<string, unknown>): string {
    return createHash("sha256")
      .update(JSON.stringify([entitySet, canonicalize(payload)]))
      .digest("hex");
  }

  private validateOperationId(operationId: string): void {
    if (!OPERATION_ID_RE.test(operationId))
      throw new InputError("operationId должен быть UUID из предпросмотра.");
  }

  private unknownResult(operationId: string, cause?: unknown): ODataError {
    const reason = cause instanceof ODataError ? ` (${cause.kind})` : "";
    return new ODataError({
      kind: "unknown",
      message:
        `Результат записи operationId ${operationId} неизвестен${reason}. Повторная отправка заблокирована. ` +
        "Проверьте документ в 1С; если он отсутствует, выполните новый dry-run с новым operationId.",
      cause,
    });
  }
}

function isDefinitiveRejection(error: unknown): boolean {
  return error instanceof ODataError && ["auth", "not_found", "bad_request"].includes(error.kind);
}

function isNodeError(error: unknown, code: string): error is NodeJS.ErrnoException {
  return error instanceof Error && "code" in error && error.code === code;
}

function canonicalize(value: unknown, key?: string): unknown {
  if (Array.isArray(value)) return value.map((item) => canonicalize(item));
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([key, nested]) => [key, canonicalize(nested, key)]),
    );
  }
  // Default document times are recalculated between preview and confirmation. The
  // request fingerprint still binds explicit user inputs; compare the business date here.
  if (key === "Date" && typeof value === "string" && /^\d{4}-\d{2}-\d{2}T/.test(value)) {
    return value.slice(0, 10);
  }
  return value;
}

function resultSummary(result: ODataEntity): Record<string, string> {
  const summary: Record<string, string> = {};
  for (const key of ["Ref_Key", "Code", "Number"] as const) {
    const value = result[key];
    if (typeof value === "string") summary[key] = value;
  }
  return summary;
}
