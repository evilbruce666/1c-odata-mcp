import type { AppConfig } from "../config/env.js";
import { logger } from "../logger.js";
import { ODataError, fromHttpStatus } from "./errors.js";
import type { ODataCollection, ODataEntity } from "../types/odata.js";

/** Разрешённые HTTP-методы. На 1-м этапе — только чтение. */
const READ_METHODS = new Set(["GET", "HEAD"]);

export class ODataClient {
  private readonly authHeader: string;

  constructor(private readonly cfg: AppConfig) {
    const token = Buffer.from(`${cfg.ODATA_USERNAME}:${cfg.ODATA_PASSWORD}`).toString("base64");
    this.authHeader = `Basic ${token}`;
  }

  /** Гард: любая запись запрещена при READ_ONLY. */
  private assertReadOnly(method: string): void {
    if (this.cfg.READ_ONLY && !READ_METHODS.has(method)) {
      throw new ODataError({
        kind: "bad_request",
        message: `Операция ${method} запрещена: сервер работает в режиме только-чтение (READ_ONLY=true)`,
      });
    }
  }

  /** Абсолютный URL: базовый URL + относительный путь (path уже с query). */
  private url(path: string): string {
    return new URL(path, this.cfg.ODATA_BASE_URL).toString();
  }

  /**
   * Низкоуровневый запрос с таймаутом и retry.
   * Возвращает распарсенный JSON указанного типа.
   */
  async request<T>(path: string, method = "GET"): Promise<T> {
    this.assertReadOnly(method);
    const url = this.url(path);
    const maxAttempts = this.cfg.ODATA_RETRIES + 1;

    let lastErr: ODataError | undefined;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), this.cfg.ODATA_TIMEOUT_MS);
      const started = performance.now();
      try {
        logger.debug({ url, method, attempt }, "odata request");
        const res = await fetch(url, {
          method,
          headers: {
            Authorization: this.authHeader,
            Accept: "application/json",
          },
          signal: controller.signal,
        });

        const ms = Math.round(performance.now() - started);
        if (!res.ok) {
          const body = await res.text().catch(() => undefined);
          const err = fromHttpStatus(res.status, url, body);
          logger.warn({ url, status: res.status, ms, kind: err.kind }, "odata error");
          if (err.retryable && attempt < maxAttempts) {
            lastErr = err;
            await backoff(attempt, res.headers.get("Retry-After"));
            continue;
          }
          throw err;
        }

        logger.debug({ url, status: res.status, ms }, "odata ok");
        return (await res.json()) as T;
      } catch (e) {
        const err = normalize(e, url);
        if (err.retryable && attempt < maxAttempts) {
          lastErr = err;
          logger.warn({ url, kind: err.kind, attempt }, "retrying");
          await backoff(attempt);
          continue;
        }
        throw err;
      } finally {
        clearTimeout(timer);
      }
    }
    throw lastErr ?? new ODataError({ kind: "unknown", message: "Запрос не выполнен", url });
  }

  /** Запрос коллекции (массив в поле value). */
  async getCollection<T extends ODataEntity = ODataEntity>(
    path: string,
  ): Promise<ODataCollection<T>> {
    return this.request<ODataCollection<T>>(path);
  }

  /** Запрос одной сущности по полному пути с ключом. */
  async getEntity<T extends ODataEntity = ODataEntity>(path: string): Promise<T> {
    return this.request<T>(path);
  }

  /** Сырой текст (для $metadata — это XML, не JSON). */
  async getText(path: string): Promise<string> {
    this.assertReadOnly("GET");
    const url = this.url(path);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.cfg.ODATA_TIMEOUT_MS);
    try {
      const res = await fetch(url, {
        method: "GET",
        headers: { Authorization: this.authHeader },
        signal: controller.signal,
      });
      if (!res.ok) {
        const body = await res.text().catch(() => undefined);
        throw fromHttpStatus(res.status, url, body);
      }
      return await res.text();
    } catch (e) {
      throw normalize(e, url);
    } finally {
      clearTimeout(timer);
    }
  }
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
