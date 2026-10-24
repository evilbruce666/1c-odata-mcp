import type { Behavior, ConnectionConfig } from "../config/env.js";
import { logger } from "../logger.js";
import { ODataError, fromHttpStatus } from "./errors.js";
import type { ODataCollection, ODataEntity } from "../types/odata.js";
import { currentWriteOperationId, currentWriteRequestHash } from "./write-operation-context.js";
import { WriteOperationJournal, type OperationCheck } from "./write-journal.js";
import { loadMetadata } from "./metadata.js";
import type { MetadataMap } from "../types/odata.js";
import { uuidV1 } from "./uuid-v1.js";
import { RequestLimiter } from "./limiter.js";

/** HTTP-методы только для чтения; остальные (POST/PATCH) считаются записью и гейтуются. */
const READ_METHODS = new Set(["GET", "HEAD"]);

export class ODataClient {
  private readonly authHeader: string;
  private metaPromise: Promise<MetadataMap> | undefined;
  private journal: WriteOperationJournal | undefined;
  /** Необязательный предел одновременных HTTP этого подключения (#57); 0 — без предела. */
  readonly limiter: RequestLimiter;

  constructor(
    private readonly conn: ConnectionConfig,
    private readonly behavior: Behavior,
  ) {
    const token = Buffer.from(`${conn.username}:${conn.password}`).toString("base64");
    this.authHeader = `Basic ${token}`;
    this.limiter = new RequestLimiter(behavior.maxConcurrentRequests ?? 0, behavior.queueTimeoutMs ?? 30_000);
  }

  getMetadata(): Promise<MetadataMap> {
    return this.metaPromise ?? this.refreshMetadata();
  }

  /** Явно перечитывает схему; последующие инструменты используют этот же кеш. */
  refreshMetadata(): Promise<MetadataMap> {
    const pending = loadMetadata(this).catch((error) => {
      // Отказ старого запроса не должен сбрасывать более новое обновление.
      if (this.metaPromise === pending) this.metaPromise = undefined;
      throw error;
    });
    this.metaPromise = pending;
    return pending;
  }

  /** Журнал операций записи создаётся при первом обращении: чтению и read-only он не нужен. */
  private get writeJournal(): WriteOperationJournal {
    this.journal ??= new WriteOperationJournal(
      this.behavior.writeJournalDir,
      this.conn.name,
      this.conn.baseUrl,
    );
    return this.journal;
  }

  /**
   * Гард записи — два явных предохранителя должны быть сняты:
   *  1) глобально READ_ONLY=false (behavior.readOnly === false);
   *  2) на конкретной базе ODATA_DB_<ИМЯ>_WRITABLE=true (conn.writable).
   */
  private assertWritable(method: string): void {
    if (READ_METHODS.has(method)) return;
    if (this.behavior.readOnly) {
      throw new ODataError({
        kind: "bad_request",
        message: `Операция ${method} запрещена: сервер в режиме только-чтение (снимите READ_ONLY=false в .env)`,
      });
    }
    if (!this.conn.writable) {
      throw new ODataError({
        kind: "bad_request",
        message: `Запись в базу "${this.conn.name}" запрещена: задайте ODATA_DB_${this.conn.name.toUpperCase()}_WRITABLE=true в .env`,
      });
    }
  }

  /** Абсолютный URL: базовый URL + относительный путь (path уже с query). */
  private url(path: string): string {
    return new URL(path, this.conn.baseUrl).toString();
  }

  /**
   * Низкоуровневый запрос с таймаутом и retry.
   * Возвращает распарсенный JSON указанного типа.
   */
  async request<T>(
    path: string,
    method = "GET",
    body?: unknown,
    headers: Record<string, string> = {},
  ): Promise<T> {
    this.assertWritable(method);
    const url = this.url(path);
    // Любая запись, включая bodyless Post/Unpost, отправляется один раз; retry только чтению.
    const maxAttempts = READ_METHODS.has(method) ? this.behavior.retries + 1 : 1;

    let lastErr: ODataError | undefined;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      // Место под HTTP — до тайм-аута запроса: ожидание места в тайм-аут 1С не входит.
      const release = await this.limiter.acquire();
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), this.behavior.timeoutMs);
      const started = performance.now();
      try {
        logger.debug({ url, method, attempt }, "odata request");
        const res = await fetch(url, {
          method,
          headers: {
            Authorization: this.authHeader,
            Accept: "application/json",
            ...(body === undefined ? {} : { "Content-Type": "application/json" }),
            ...headers,
          },
          ...(body === undefined ? {} : { body: JSON.stringify(body) }),
          signal: controller.signal,
        });

        const ms = Math.round(performance.now() - started);
        if (!res.ok) {
          const body = await res.text().catch(() => undefined);
          const err = fromHttpStatus(res.status, url, body, { method, path });
          logger.warn({ url, status: res.status, ms, kind: err.kind }, "odata error");
          if (err.retryable && attempt < maxAttempts) {
            lastErr = err;
            release(); // пауза между повторами не держит место
            await backoff(attempt, res.headers.get("Retry-After"));
            continue;
          }
          throw err;
        }

        logger.debug({ url, status: res.status, ms }, "odata ok");
        // Действия (Post/Unpost) могут вернуть пустое тело — это не ошибка.
        const text = await res.text();
        return parseBody<T>(text, res.headers.get("Content-Type"), url, method);
      } catch (e) {
        const err = normalize(e, url);
        if (err.retryable && attempt < maxAttempts) {
          lastErr = err;
          logger.warn({ url, kind: err.kind, attempt }, "retrying");
          release();
          await backoff(attempt);
          continue;
        }
        throw err;
      } finally {
        clearTimeout(timer);
        release();
      }
    }
    throw lastErr ?? new ODataError({ kind: "unknown", message: "Запрос не выполнен", url });
  }

  /** Запрос коллекции (массив в поле value). */
  async getCollection<T extends ODataEntity = ODataEntity>(path: string): Promise<ODataCollection<T>> {
    return this.request<ODataCollection<T>>(path);
  }

  /** Запрос одной сущности по полному пути с ключом. */
  async getEntity<T extends ODataEntity = ODataEntity>(path: string): Promise<T> {
    return this.request<T>(path);
  }

  /** Создаёт объект (POST). Возвращает созданную сущность с Ref_Key. */
  async create<T extends ODataEntity = ODataEntity>(entitySet: string, payload: object): Promise<T> {
    // Гард — до резервирования в журнале: в режиме только-чтение операция не должна оставлять следов.
    this.assertWritable("POST");
    const send = async (body: object) => {
      const created = await this.request<T>(`${entitySet}?$format=json`, "POST", body);
      if (!created || typeof created !== "object" || Array.isArray(created)) {
        throw new ODataError({
          kind: "unknown",
          message: "1С ответила на создание без объекта; результат записи нужно сверить вручную.",
        });
      }
      return created;
    };
    const operationId = currentWriteOperationId();
    if (operationId) {
      const requestHash = currentWriteRequestHash();
      if (!requestHash) throw new Error("Для подтверждённой записи отсутствует отпечаток предпросмотра.");
      // Ref_Key, назначенный при предпросмотре: повтор с ним 1С отвергнет (уникальный индекс),
      // а сверка после таймаута — простой GET. В отпечаток payload не входит.
      const refKey = await this.writeJournal.refKeyOf(operationId);
      const body = refKey ? { Ref_Key: refKey, ...payload } : payload;
      return this.writeJournal.execute(
        operationId,
        entitySet,
        payload as Record<string, unknown>,
        requestHash,
        () => send(body),
      );
    }
    return send(payload);
  }

  /**
   * Предпросмотр операции записи: отпечаток в журнал. Для создания (withRefKey) назначается и
   * Ref_Key будущего объекта (UUIDv1) — подтверждение отправит именно его.
   */
  async prepareCreate(
    entitySet: string,
    payload: Record<string, unknown>,
    { withRefKey = true }: { withRefKey?: boolean } = {},
  ): Promise<void> {
    // Предпросмотр при выключенной записи ничего не пишет на диск (READ_ONLY по умолчанию).
    if (this.behavior.readOnly || !this.conn.writable) return;
    const operationId = currentWriteOperationId();
    const requestHash = currentWriteRequestHash();
    if (!operationId || !requestHash)
      throw new Error("Для предпросмотра записи отсутствуют operationId и отпечаток запроса.");
    await this.writeJournal.prepare(
      operationId,
      entitySet,
      payload,
      requestHash,
      withRefKey ? uuidV1() : undefined,
    );
  }

  /** Запись журнала операции записи (для write.operation.status). */
  operationEntry(operationId: string) {
    return this.writeJournal.lookup(operationId);
  }

  /** Фиксирует найденный сверкой результат неизвестной операции. */
  reconcileOperation(operationId: string, result: ODataEntity): Promise<void> {
    return this.writeJournal.reconcile(operationId, result);
  }

  /**
   * PATCH, который нельзя слепо повторять (добавление/удаление строки документа): при
   * подтверждении с operationId идёт через журнал. key — стабильная часть операции
   * (одинаковая в предпросмотре и подтверждении), check — как сверить результат с 1С.
   */
  async patchOnce<T extends ODataEntity = ODataEntity>(
    path: string,
    payload: object,
    entitySet: string,
    key: Record<string, unknown>,
    check: OperationCheck,
  ): Promise<T> {
    this.assertWritable("PATCH");
    const send = () => this.request<T>(path, "PATCH", payload);
    const operationId = currentWriteOperationId();
    if (!operationId) return send();
    const requestHash = currentWriteRequestHash();
    if (!requestHash) throw new Error("Для подтверждённой записи отсутствует отпечаток предпросмотра.");
    return this.writeJournal.execute(operationId, entitySet, key, requestHash, send, check);
  }

  /** Journaled independent-register write; verification remains inside the durable operation. */
  async writeRecordOnce<T extends ODataEntity>(
    entitySet: string,
    operation: Record<string, unknown>,
    sendAndVerify: () => Promise<T>,
    check?: OperationCheck,
  ): Promise<T> {
    this.assertWritable("POST/PATCH");
    const operationId = currentWriteOperationId();
    const requestHash = currentWriteRequestHash();
    if (!operationId || !requestHash)
      throw new Error("Для записи настройки нужен operationId из предпросмотра.");
    return this.writeJournal.execute(operationId, entitySet, operation, requestHash, sendAndVerify, check);
  }

  /** Итог уже подтверждённой операции (для повтора до повторного чтения 1С); undefined — не подтверждалась. */
  async operationSettled(
    operationId: string,
    requestHash: string,
  ): Promise<{ entitySet: string; result: Record<string, string> } | undefined> {
    if (this.behavior.readOnly || !this.conn.writable) return undefined;
    return this.writeJournal.settled(operationId, requestHash);
  }

  /** Состояние журнала для health_check (только когда запись в базу включена). */
  async journalSummary() {
    if (this.behavior.readOnly || !this.conn.writable) return undefined;
    return this.writeJournal.summary();
  }

  /** Сверка показала, что операция до 1С не дошла. */
  markOperationNotApplied(operationId: string): Promise<void> {
    return this.writeJournal.markNotApplied(operationId);
  }

  /** Изменяет объект (PATCH) по полному пути с ключом. */
  async patch<T extends ODataEntity = ODataEntity>(path: string, payload: object): Promise<T> {
    return this.request<T>(path, "PATCH", payload);
  }

  /** Вызывает bound-действие (напр. .../Post). POST без тела. */
  async action<T = unknown>(path: string): Promise<T> {
    return this.request<T>(path, "POST");
  }

  /** Сырой текст (для $metadata — это XML, не JSON). */
  async getText(path: string): Promise<string> {
    this.assertWritable("GET");
    const url = this.url(path);
    const release = await this.limiter.acquire();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.behavior.timeoutMs);
    try {
      const res = await fetch(url, {
        method: "GET",
        headers: { Authorization: this.authHeader },
        signal: controller.signal,
      });
      if (!res.ok) {
        const body = await res.text().catch(() => undefined);
        throw fromHttpStatus(res.status, url, body, { method: "GET", path });
      }
      const text = await res.text();
      if (isHtml(text, res.headers.get("Content-Type"))) throw htmlPageError(text, url, "GET");
      return text;
    } catch (e) {
      throw normalize(e, url);
    } finally {
      clearTimeout(timer);
      release();
    }
  }
}

/**
 * Тело успешного ответа. Публикация 1С (в том числе 1cfresh) на отказ доступа или исключение
 * сервера может ответить HTML-страницей с кодом 200 — это не OData и не сетевой сбой: повторять
 * такой запрос бесполезно, а при записи исход неизвестен (kind parse — не однозначный отказ).
 */
function parseBody<T>(text: string, contentType: string | null, url: string, method: string): T {
  if (!text) return undefined as T;
  if (isHtml(text, contentType)) throw htmlPageError(text, url, method);
  try {
    return JSON.parse(text) as T;
  } catch (e) {
    throw new ODataError({
      kind: "parse",
      url,
      cause: e,
      message: `1С ответила не в формате OData (JSON): ${pageText(text)}${writeNote(method)}`,
    });
  }
}

function isHtml(text: string, contentType: string | null): boolean {
  return /text\/html/i.test(contentType ?? "") || /^\s*<(!doctype html|html)[\s>]/i.test(text);
}

function htmlPageError(text: string, url: string, method: string): ODataError {
  return new ODataError({
    kind: "parse",
    url,
    message:
      `1С вернула HTML-страницу вместо ответа OData: ${pageText(text)}. Обычно это страница входа, ` +
      "отказа в доступе или ошибки публикации; проверьте доступ к базе и публикацию OData." +
      writeNote(method),
  });
}

/** Короткий текст страницы для сообщения: заголовок или начало текста без разметки, до 200 символов. */
function pageText(text: string): string {
  const title = /<title[^>]*>([^<]*)<\/title>/i.exec(text)?.[1]?.trim();
  const plain = (title || text.replace(/<(script|style)[\s\S]*?<\/\1>/gi, " ").replace(/<[^>]+>/g, " "))
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return plain.length > 200 ? `${plain.slice(0, 200)}…` : plain || "пустая страница";
}

function writeNote(method: string): string {
  return READ_METHODS.has(method)
    ? ""
    : " Запрос на запись отправлен, исход неизвестен — сверьте результат в 1С, прежде чем повторять.";
}

/** Преобразует исключение fetch/AbortController в ODataError. */
function normalize(e: unknown, url: string): ODataError {
  if (e instanceof ODataError) return e;
  if (e instanceof Error && e.name === "AbortError") {
    return new ODataError({ kind: "timeout", message: "Превышен таймаут запроса", url, cause: e });
  }
  return new ODataError({
    kind: "network",
    message: e instanceof Error ? e.message : "Сетевая ошибка",
    url,
    cause: e,
  });
}

/** Экспоненциальная пауза с учётом заголовка Retry-After (если есть). */
async function backoff(attempt: number, retryAfter?: string | null): Promise<void> {
  let ms = Math.min(1_000 * 2 ** (attempt - 1), 10_000);
  if (retryAfter) {
    const sec = Number(retryAfter);
    if (Number.isFinite(sec)) ms = Math.max(ms, sec * 1_000);
  }
  await new Promise((r) => setTimeout(r, ms));
}
