import { createHash, randomUUID } from "node:crypto";
import { mkdir, open, readFile, readdir, rename, rm, stat } from "node:fs/promises";
import { dirname, join } from "node:path";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "./errors.js";

const OPERATION_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const TERMINAL_RETENTION_MS = 90 * 24 * 60 * 60 * 1_000;
const UNCERTAIN_RETENTION_MS = 365 * 24 * 60 * 60 * 1_000;
const STALE_FILE_MS = 24 * 60 * 60 * 1_000;
const PRUNE_INTERVAL_MS = 60 * 60 * 1_000;

export type OperationState = "prepared" | "executing" | "succeeded" | "rejected" | "outcome_unknown";

export interface JournalEntry {
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

  /** Запись журнала по id (undefined — такой операции нет); для инструмента статуса. */
  async lookup(operationId: string): Promise<JournalEntry | undefined> {
    this.validateOperationId(operationId);
    try {
      return await this.read(operationId);
    } catch (error) {
      if (error instanceof InputError && error.message.includes("не найден в локальном журнале"))
        return undefined;
      throw error;
    }
  }

  /**
   * Результат неизвестной операции выяснен сверкой с 1С (объект найден по метке):
   * фиксируем успех, чтобы повторное подтверждение вернуло ссылку, а не блокировалось.
   */
  async reconcile(operationId: string, result: ODataEntity): Promise<void> {
    const entry = await this.read(operationId);
    if (entry.state === "succeeded") return;
    await this.replace({
      ...entry,
      state: "succeeded",
      updatedAt: new Date().toISOString(),
      result: resultSummary(result),
    });
    await rm(this.lockPath(operationId), { force: true }).catch(() => undefined);
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

  /**
   * Раз в час чистит каталог журнала:
   *  - завершённые/отклонённые/непотверждённые записи старше 90 дней;
   *  - записи с неизвестным исходом — только через год (раньше нельзя: удаление разрешило бы опасный повтор);
   *  - «осиротевшие» .lock (записи нет, она завершена, либо подтверждение так и не отправлялось > суток);
   *  - зависшие .tmp старше суток.
   */
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
    const remove = (name: string) => rm(join(this.directory, name), { force: true }).catch(() => undefined);
    const ageOf = async (name: string): Promise<number> => {
      try {
        return now - (await stat(join(this.directory, name))).mtimeMs;
      } catch {
        return 0;
      }
    };
    const states = new Map<string, OperationState | undefined>();
    await Promise.all(
      files
        .filter((file) => file.endsWith(".json"))
        .map(async (file) => {
          const id = file.slice(0, -".json".length);
          try {
            const entry = JSON.parse(await readFile(join(this.directory, file), "utf8")) as JournalEntry;
            const age = now - Date.parse(entry.updatedAt);
            const uncertain = entry.state === "executing" || entry.state === "outcome_unknown";
            if (age > (uncertain ? UNCERTAIN_RETENTION_MS : TERMINAL_RETENTION_MS)) {
              await remove(file);
              await remove(`${id}.lock`);
              return;
            }
            states.set(id, entry.state);
          } catch {
            // Нечитаемые записи не трогаем: удаление могло бы сделать повтор небезопасным.
            states.set(id, "outcome_unknown");
          }
        }),
    );
    await Promise.all(
      files.map(async (file) => {
        if (file.endsWith(".tmp")) {
          if ((await ageOf(file)) > STALE_FILE_MS) await remove(file);
        } else if (file.endsWith(".lock")) {
          const state = states.get(file.slice(0, -".lock".length));
          const age = await ageOf(file);
          const done = state === undefined || state === "succeeded" || state === "rejected";
          // prepared + старый lock: процесс упал до отправки (состояние «executing» пишется до POST).
          if ((done && age > PRUNE_INTERVAL_MS) || (state === "prepared" && age > STALE_FILE_MS)) {
            await remove(file);
          }
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

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((item) => canonicalize(item));
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
        .map(([key, nested]) => [key, canonicalize(nested)]),
    );
  }
  // Время суток по умолчанию («сейчас») пересчитывается между предпросмотром и подтверждением
  // (поля Date / Дата / ДатаНачала … — по-разному в разных документах). Явный ввод вызывающего
  // всё равно связан отпечатком аргументов (requestHash), поэтому здесь сравниваем только день.
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(value)) {
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
