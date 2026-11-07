import { AsyncLocalStorage } from "node:async_hooks";
import { ODataError } from "./errors.js";

/**
 * Контекст одного вызова инструмента чтения: отмена от клиента и общий лимит времени. Без него клиент
 * уходит по таймауту, а сервер ещё минуты шлёт запросы и занимает очередь 1С одной учётной записи.
 * Запись сюда не входит намеренно: обрыв отправленного запроса на запись оставляет неизвестный исход.
 */
interface CallContext {
  signal?: AbortSignal;
  /** Предел текста ответа этого вызова (ODATA_MAX_RESPONSE_CHARS). */
  maxResponseChars?: number | undefined;
  /** Сработавший лимит времени отличается от отмены клиентом только текстом и видом ошибки. */
  deadline?: { signal: AbortSignal; ms: number };
}

const store = new AsyncLocalStorage<CallContext>();
const connection = new AsyncLocalStorage<AbortSignal>();

/**
 * Соединение HTTP-запроса. Без сессии сообщение об отмене приходит отдельным запросом и до вызова не доходит,
 * поэтому признак ухода клиента в этом режиме — разрыв соединения.
 */
export const withConnection = <T>(signal: AbortSignal, fn: () => T): T => connection.run(signal, fn);

/** Выполняет вызов инструмента с отменой от клиента и лимитом времени (0 — без лимита). */
export function withToolCall<T>(
  o: { signal?: AbortSignal | undefined; deadlineMs: number; maxResponseChars?: number | undefined },
  fn: () => Promise<T>,
): Promise<T> {
  const deadline =
    o.deadlineMs > 0 ? { signal: AbortSignal.timeout(o.deadlineMs), ms: o.deadlineMs } : undefined;
  const signals = [o.signal, connection.getStore(), deadline?.signal].filter((s): s is AbortSignal => !!s);
  return store.run(
    {
      ...(signals.length ? { signal: AbortSignal.any(signals) } : {}),
      ...(deadline ? { deadline } : {}),
      maxResponseChars: o.maxResponseChars,
    },
    fn,
  );
}

/** Сигнал текущего вызова; undefined — вне вызова чтения (запись, служебный код). */
export const callSignal = (): AbortSignal | undefined => store.getStore()?.signal;

/** Предел текста ответа текущего вызова; undefined — действует предел по умолчанию. */
export const callResponseLimit = (): number | undefined => store.getStore()?.maxResponseChars;

/** Ошибка прерванного вызова, если он прерван; иначе undefined. */
export const callAbortError = (): ODataError | undefined => abortErrorOf(store.getStore());

function abortErrorOf(call: CallContext | undefined): ODataError | undefined {
  if (!call?.signal?.aborted) return undefined;
  if (call.deadline?.signal.aborted) {
    return new ODataError({
      kind: "deadline",
      message:
        `Вызов не уложился в ODATA_TOOL_DEADLINE_MS (${call.deadline.ms} мс), его запросы к 1С остановлены. ` +
        "Частичных данных в ответе нет: сузьте период или отбор либо повторите позже.",
    });
  }
  return new ODataError({
    kind: "cancelled",
    message: "Вызов отменён клиентом, его запросы к 1С остановлены.",
  });
}

/**
 * Загрузка, которую ждут и другие вызовы (кэш справочников, $metadata): её нельзя прерывать отменой
 * того, кто начал её первым, иначе соседний вызов получит чужую отмену.
 */
export const detachedFromCall = <T>(load: () => T): T => store.exit(load);

/** Ожидание общей загрузки: прерывается отменой текущего вызова, сама загрузка продолжается. */
export function untilCallAborted<T>(pending: Promise<T>): Promise<T> {
  const call = store.getStore();
  const signal = call?.signal;
  if (!signal) return pending;
  return new Promise<T>((resolve, reject) => {
    // Обработчик abort выполняется в контексте того, кто отменил, поэтому вызов берём из замыкания.
    const onAbort = () => reject(abortErrorOf(call));
    if (signal.aborted) onAbort();
    else signal.addEventListener("abort", onAbort, { once: true });
    pending.then(resolve, reject).finally(() => signal.removeEventListener("abort", onAbort));
  });
}
