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

/**
 * Метка запуска процесса — в lock-файле операции. Чужая метка — lock упавшего или перезапущенного процесса: один
 * процесс записи на каталог журнала (требование 1.1, docs/WRITE.md). PID для этого не годится: в контейнере после
 * рестарта он снова 1.
 */
const BOOT_ID = randomUUID();
/** Операции, по которым этот процесс прямо сейчас ждёт ответа 1С: каталог журнала + operationId. */
const LIVE = new Set<string>();

/**
 * Создание с назначенным Ref_Key и неизвестным исходом (1.1): повтор подтверждения безопасен — второй объект с той
 * же ссылкой 1С не создаст (уникальный индекс; проверено на копии УТ 11.5 последовательным повтором и пятью
 * одновременными POST — ровно один объект). Правки, проведение, записи регистров и операции без ссылки — нет.
 */
export function isRetryableCreate(entry: JournalEntry): entry is JournalEntry & { refKey: string } {
  return (entry.state === "executing" || entry.state === "outcome_unknown") && !!entry.refKey && !entry.check;
}

/** Запрос по операции ещё идёт в этом процессе: подождать и повторить подтверждение. */
export function inProgressError(operationId: string): ODataError {
  return new ODataError({
    kind: "in_progress",
    message:
      `Запрос записи operationId ${operationId} ещё выполняется в этом процессе. Подождите ответа и повторите ` +
      "подтверждение с тем же operationId и аргументами — второго объекта не будет.",
  });
}

export type OperationState = "prepared" | "executing" | "succeeded" | "rejected" | "outcome_unknown";

/**
 * Как сверить операцию, у которой нет «следа» в 1С (правка строк документа): число строк
 * табличной части до записи и ожидаемое после. Пишется в журнал перед отправкой PATCH.
 */
interface LineCountCheck {
  kind: "lineCount";
  ref: string;
  section: string;
  before: number;
  expected: number;
}

export type OperationCheck =
  | LineCountCheck
  | { kind: "documentState"; ref: string }
  | { kind: "createdRef"; ref: string }
  | { kind: "manual"; target: "counterpartyGroup" }
  | { kind: "registerValues"; recordPath: string; fields: string[]; expectedHash: string };

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
  /** Текст ошибки последней попытки (сообщение 1С), не длиннее ERROR_MESSAGE_MAX — для сверки. */
  errorMessage?: string;
  check?: OperationCheck;
  /** Ref_Key, который MCP задаёт создаваемому объекту: по нему сверка — простым GET. */
  refKey?: string;
}

/**
 * Durable, local idempotency journal for confirmed OData creates and non-idempotent document edits.
 * Only a payload hash, a small result reference and, after a failed attempt, the 1C error text
 * (bounded, see ERROR_MESSAGE_MAX) are persisted — never the payload itself.
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
    refKey?: string,
  ): Promise<string | undefined> {
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
      ...(refKey ? { refKey } : {}),
    };

    try {
      await this.writeNew(this.recordPath(operationId), entry);
      return entry.refKey;
    } catch (error) {
      if (!isNodeError(error, "EEXIST")) throw error;
      const existing = await this.read(operationId);
      this.assertMatches(existing, entitySet, payload, requestHash);
      return existing.refKey;
    }
  }

  /** Запрос по операции идёт в этом процессе прямо сейчас. */
  isLive(operationId: string): boolean {
    return LIVE.has(this.liveKey(operationId));
  }

  private liveKey(operationId: string): string {
    return `${this.directory}\n${operationId}`;
  }

  async execute<T extends ODataEntity>(
    operationId: string,
    entitySet: string,
    payload: Record<string, unknown>,
    requestHash: string,
    send: () => Promise<T>,
    check?: OperationCheck,
    /** Создание с назначенной ссылкой: прочитать объект по ней (undefined — его нет). Включает безопасный повтор. */
    findExisting?: () => Promise<T | undefined>,
  ): Promise<T> {
    this.validateOperationId(operationId);
    let entry = await this.read(operationId);
    if (findExisting && isRetryableCreate(entry))
      return this.retryCreate(entry, entitySet, payload, requestHash, send, findExisting);
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
      if (findExisting && isRetryableCreate(entry))
        return this.retryCreate(entry, entitySet, payload, requestHash, send, findExisting);
      this.assertMatches(entry, entitySet, payload, requestHash);
      if (this.isLive(operationId)) throw inProgressError(operationId);
      const racedResult = this.resultForExisting<T>(entry);
      if (racedResult !== undefined) return racedResult;
      throw this.unknownResult(operationId);
    }
    // «Запрос идёт» — сразу при захвате lock, до записи «executing»: второе подтверждение в этом окне получает
    // in_progress, а не «исход неизвестен» (живая проверка 1.1, два одновременных подтверждения по HTTP).
    const live = this.liveKey(operationId);
    LIVE.add(live);
    await lock.writeFile(BOOT_ID, "utf8").catch(() => undefined);

    let requestStarted = false;
    try {
      entry = await this.read(operationId);
      this.assertMatches(entry, entitySet, payload, requestHash);
      const racedResult =
        findExisting && isRetryableCreate(entry) ? undefined : this.resultForExisting<T>(entry);
      if (racedResult !== undefined) return racedResult;
      if (entry.state !== "prepared") {
        if (findExisting && isRetryableCreate(entry)) {
          // Другое подтверждение успело закончиться неизвестным исходом, пока ждали lock: дальше — путь повтора.
          await lock.close().catch(() => undefined);
          await rm(lockPath, { force: true }).catch(() => undefined);
          LIVE.delete(live);
          return this.retryCreate(entry, entitySet, payload, requestHash, send, findExisting);
        }
        throw this.unknownResult(operationId);
      }

      entry = {
        ...entry,
        state: "executing",
        updatedAt: new Date().toISOString(),
        ...(check ? { check } : {}),
      };
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
          ...(cause instanceof ODataError
            ? { errorKind: cause.kind, errorMessage: boundedMessage(cause.message) }
            : {}),
        };
        try {
          await this.replace(failed);
        } catch (journalError) {
          throw this.unknownResult(operationId, journalError);
        }
        await rm(lockPath, { force: true }).catch(() => undefined);
        if (rejected) throw cause;
        throw this.unknownResult(operationId, cause, !!findExisting && isRetryableCreate(failed));
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
      LIVE.delete(live);
      await lock.close().catch(() => undefined);
    }
  }

  /**
   * Повтор создания с назначенной ссылкой после неизвестного исхода (1.1). Сначала — чтение по ссылке: объект есть →
   * успех без POST. Нет → тот же POST с той же ссылкой: если первый запрос всё-таки дописался или дописывается, 1С
   * откажет по уникальности, и повторное чтение найдёт объект. Исход повтора определяет чтение, а не текст ошибки:
   * текст отказа зависит от СУБД. Отказ 1С на повторе не делает операцию «отклонённой»: первый запрос мог ещё
   * выполняться, — операция остаётся с неизвестным исходом, повтор снова безопасен.
   */
  private async retryCreate<T extends ODataEntity>(
    entry: JournalEntry & { refKey: string },
    entitySet: string,
    payload: Record<string, unknown>,
    requestHash: string,
    send: () => Promise<T>,
    findExisting: () => Promise<T | undefined>,
  ): Promise<T> {
    const operationId = entry.operationId;
    if (entry.database !== this.database) throw new InputError("operationId создан для другой базы 1С.");
    if (entry.entitySet !== entitySet || entry.requestHash !== requestHash)
      throw new InputError(
        `Данные, объект или база не совпадают с dry-run для operationId ${operationId}. Повторите подтверждение с теми же аргументами, что в предпросмотре.`,
      );
    const live = this.liveKey(operationId);
    // Проверка и отметка — без await между ними: два подтверждения в одном процессе не пройдут сюда оба.
    if (LIVE.has(live)) throw inProgressError(operationId);
    LIVE.add(live);
    try {
      const lockPath = this.lockPath(operationId);
      const lock = await this.takeOverLock(lockPath, operationId);
      try {
        const succeed = async (found: T): Promise<T> => {
          await this.replace({
            ...entry,
            state: "succeeded",
            updatedAt: new Date().toISOString(),
            result: resultSummary(found),
          });
          await rm(lockPath, { force: true }).catch(() => undefined);
          return found;
        };
        const before = await findExisting();
        if (before) return { ...(await succeed(before)), _operation_reconciled: true };
        if (entry.payloadHash !== this.payloadHash(entitySet, payload)) {
          await rm(lockPath, { force: true }).catch(() => undefined);
          throw new ODataError({
            kind: "unknown",
            message:
              `Результат записи operationId ${operationId} неизвестен, а данные из 1С для этого объекта изменились с ` +
              "предпросмотра (подстановки — договор, счета, соглашение): повторить ту же запись нельзя. Объекта с " +
              "назначенной ссылкой в 1С пока нет. Сверьте базу и закройте операцию через write.operation.status " +
              "с markNotApplied=true, затем сделайте новый предпросмотр.",
          });
        }
        await this.replace({ ...entry, state: "executing", updatedAt: new Date().toISOString() });
        let result: T;
        try {
          result = await send();
        } catch (cause) {
          const after = await findExisting().catch(() => undefined);
          if (after) return { ...(await succeed(after)), _operation_reconciled: true };
          await this.replace({
            ...entry,
            state: "outcome_unknown",
            updatedAt: new Date().toISOString(),
            ...(cause instanceof ODataError
              ? { errorKind: cause.kind, errorMessage: boundedMessage(cause.message) }
              : {}),
          });
          await rm(lockPath, { force: true }).catch(() => undefined);
          throw this.unknownResult(operationId, cause, true);
        }
        return { ...(await succeed(result)), _operation_retried: true };
      } finally {
        await lock.close().catch(() => undefined);
      }
    } finally {
      LIVE.delete(live);
    }
  }

  /**
   * Lock операции для повтора. Занят процессом, который сейчас ждёт ответа, — сюда не доходит (LIVE выше); остальной
   * lock — от упавшего, перезапущенного или потерявшего журнал процесса: снимаем и берём заново.
   */
  private async takeOverLock(lockPath: string, operationId: string) {
    try {
      const lock = await open(lockPath, "wx", 0o600);
      await lock.writeFile(BOOT_ID, "utf8").catch(() => undefined);
      return lock;
    } catch (error) {
      if (!isNodeError(error, "EEXIST")) throw error;
    }
    await rm(lockPath, { force: true });
    try {
      const lock = await open(lockPath, "wx", 0o600);
      await lock.writeFile(BOOT_ID, "utf8").catch(() => undefined);
      return lock;
    } catch (error) {
      if (isNodeError(error, "EEXIST")) throw inProgressError(operationId);
      throw error;
    }
  }

  /**
   * Сводка для health_check: доступен ли каталог журнала для записи и какие операции ждут
   * сверки (исход неизвестен) — чтобы они не терялись. Каталог создаётся, если его нет.
   */
  async summary(): Promise<{ dir: string; writable: boolean; uncertain: string[]; error?: string }> {
    try {
      await mkdir(this.directory, { recursive: true, mode: 0o700 });
      const probe = join(this.directory, `.probe-${randomUUID()}`);
      await (await open(probe, "wx", 0o600)).close();
      await rm(probe, { force: true });
    } catch (error) {
      return {
        dir: this.directory,
        writable: false,
        uncertain: [],
        error: error instanceof Error ? error.message : String(error),
      };
    }
    const uncertain: string[] = [];
    for (const file of await readdir(this.directory)) {
      if (!file.endsWith(".json")) continue;
      try {
        const entry = JSON.parse(await readFile(join(this.directory, file), "utf8")) as JournalEntry;
        if (entry.state === "executing" || entry.state === "outcome_unknown")
          uncertain.push(entry.operationId);
      } catch {
        uncertain.push(file.slice(0, -".json".length));
      }
    }
    return { dir: this.directory, writable: true, uncertain };
  }

  /** Ref_Key, назначенный операции при предпросмотре (у старых записей и правок строк — нет). */
  async refKeyOf(operationId: string): Promise<string | undefined> {
    return (await this.lookup(operationId))?.refKey;
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

  /**
   * Сверка показала, что запись до 1С не дошла (документ в исходном состоянии):
   * операция закрывается как неприменённая, повтор по этому id запрещён — нужен новый предпросмотр.
   */
  async markNotApplied(operationId: string): Promise<void> {
    const entry = await this.read(operationId);
    if (entry.state === "succeeded" || entry.state === "rejected") return;
    await this.replace({
      ...entry,
      state: "rejected",
      errorKind: "not_applied",
      updatedAt: new Date().toISOString(),
    });
    await rm(this.lockPath(operationId), { force: true }).catch(() => undefined);
  }

  /**
   * Состояние операции для подтверждения с этим operationId — до того, как инструмент
   * заново прочитает 1С: успешную возвращает сразу, неизвестную/отклонённую — ошибкой.
   * undefined — подтверждение ещё не выполнялось (или операции нет в журнале).
   */
  async settled(
    operationId: string,
    requestHash: string,
  ): Promise<{ entitySet: string; result: Record<string, string> } | undefined> {
    const entry = await this.lookup(operationId);
    if (!entry || entry.state === "prepared") return undefined;
    if (entry.database !== this.database || entry.requestHash !== requestHash) {
      throw new InputError(
        `Данные, объект или база не совпадают с dry-run для operationId ${operationId}. Выполните новый предпросмотр.`,
      );
    }
    this.resultForExisting(entry); // неизвестный/отклонённый исход — ошибкой
    return { entitySet: entry.entitySet, result: entry.result ?? {} };
  }

  private resultForExisting<T extends ODataEntity>(entry: JournalEntry): T | undefined {
    if (entry.state === "succeeded") {
      return { ...(entry.result ?? {}), _operation_replayed: true } as unknown as T;
    }
    if (entry.state === "rejected" && entry.errorKind === "not_applied") {
      throw new InputError(
        `Операция operationId ${entry.operationId} не была применена (сверено с документом в 1С). Выполните новый предпросмотр с новым operationId.`,
      );
    }
    if (entry.state === "rejected") {
      throw new InputError(
        `Попытка записи operationId ${entry.operationId} ранее была отклонена. Исправьте данные, выполните новый dry-run и подтвердите новый operationId.`,
      );
    }
    if (entry.state === "executing" || entry.state === "outcome_unknown") {
      throw this.unknownResult(entry.operationId, undefined, isRetryableCreate(entry));
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

  private unknownResult(operationId: string, cause?: unknown, retrySafe = false): ODataError {
    // Текст 1С (например, «Ошибка сервера 1С (500): …») нужен для сверки — не теряем его (#55).
    const reason = cause instanceof ODataError ? ` (${cause.kind}: ${boundedMessage(cause.message)})` : "";
    return new ODataError({
      kind: "unknown",
      message: retrySafe
        ? `Результат записи operationId ${operationId} неизвестен${reason}. Повторите подтверждение с тем же ` +
          "operationId и теми же аргументами: объект с назначенной ссылкой 1С создаст не больше одного раза, а повтор " +
          "сначала проверит, не создан ли он уже. Новый предпросмотр с новым operationId до этого не делайте."
        : `Результат записи operationId ${operationId} неизвестен${reason}. Повторная отправка заблокирована. ` +
          "Проверьте состояние операции через write.operation.status и точный объект или настройку в 1С. Новый предпросмотр допустим только после подтверждения отсутствия результата.",
      cause,
    });
  }
}

const ERROR_MESSAGE_MAX = 500;

function boundedMessage(message: string): string {
  return message.length > ERROR_MESSAGE_MAX ? `${message.slice(0, ERROR_MESSAGE_MAX)}…` : message;
}

function isDefinitiveRejection(error: unknown): boolean {
  // busy — запрос не дождался места под HTTP и в 1С не отправлялся: исход известен, дубля не будет.
  return error instanceof ODataError && ["auth", "not_found", "bad_request", "busy"].includes(error.kind);
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
  for (const key of ["Ref_Key", "Code", "Number", "RecordPath"] as const) {
    const value = result[key];
    if (typeof value === "string") summary[key] = value;
  }
  return summary;
}
