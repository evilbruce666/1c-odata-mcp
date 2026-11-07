import { ODataError } from "./errors.js";

/**
 * Необязательный предел одновременных HTTP-запросов одного подключения (#57).
 *
 * Место занимается на время фактического HTTP-обмена (запрос и чтение тела ответа) и всегда
 * освобождается — при успехе, ошибке и тайм-ауте; перед паузой между повторами вызывающий код
 * освобождает его сам. Ожидание места ограничено по времени: не дождался — запрос не отправлялся,
 * и это сообщается явно (kind "busy"), чтобы журнал записи не считал исход неизвестным.
 * limit = 0 — без ограничений, поведение как раньше. Это не очередь задач и не ограничение других
 * процессов или пользователей 1С.
 */
export class RequestLimiter {
  private active = 0;
  private readonly waiting: Array<() => void> = [];

  constructor(
    private readonly limit: number,
    private readonly waitTimeoutMs: number,
  ) {}

  /**
   * Занимает место; возвращает идемпотентную функцию освобождения. signal прерывает ожидание: отменённому
   * вызову место уже не нужно, а причину сообщает вызывающий код.
   */
  async acquire(signal?: AbortSignal): Promise<() => void> {
    if (this.limit <= 0) return () => undefined;
    if (this.active < this.limit) {
      this.active += 1;
      return this.releaser();
    }
    await new Promise<void>((resolve, reject) => {
      const leave = () => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", onAbort);
        const i = this.waiting.indexOf(wake);
        if (i >= 0) this.waiting.splice(i, 1);
      };
      const wake = () => {
        leave();
        resolve();
      };
      const onAbort = () => {
        leave();
        reject(new DOMException("Ожидание места прервано", "AbortError"));
      };
      const timer = setTimeout(() => {
        leave();
        reject(
          new ODataError({
            kind: "busy",
            retryable: false,
            message:
              `Все ${this.limit} мест для запросов к этой базе 1С заняты дольше ${this.waitTimeoutMs} мс ` +
              "(ODATA_MAX_CONCURRENT_REQUESTS). Запрос в 1С не отправлялся — повторите позже.",
          }),
        );
      }, this.waitTimeoutMs);
      signal?.addEventListener("abort", onAbort, { once: true });
      this.waiting.push(wake);
    });
    // Место передано освобождающим (active не уменьшался), FIFO.
    return this.releaser();
  }

  /** Сколько запросов сейчас выполняется (для тестов и диагностики). */
  get inFlight(): number {
    return this.active;
  }

  private releaser(): () => void {
    let released = false;
    return () => {
      if (released) return;
      released = true;
      const next = this.waiting.shift();
      if (next) next();
      else this.active -= 1;
    };
  }
}
