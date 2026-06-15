/**
 * Классификация ошибок при работе с OData 1С.
 * Нормализуем разнородные сбои (сеть, HTTP, тело ошибки 1С) в один тип
 * с понятным русским сообщением и признаком «можно ли повторить».
 */

export type ODataErrorKind =
  | "auth" // 401/403 — неверные креды или нет прав в «Состав OData»
  | "not_found" // 404 — объект не опубликован или не существует
  | "bad_request" // 400 — кривой $filter/$select
  | "rate_limit" // 429
  | "server" // 5xx
  | "timeout" // прервано по таймауту
  | "network" // DNS/соединение
  | "parse" // не удалось разобрать ответ
  | "unknown";

export class ODataError extends Error {
  readonly kind: ODataErrorKind;
  readonly status?: number;
  readonly retryable: boolean;
  readonly url?: string;

  constructor(opts: {
    kind: ODataErrorKind;
    message: string;
    status?: number;
    retryable?: boolean;
    url?: string;
    cause?: unknown;
  }) {
    super(opts.message, { cause: opts.cause });
    this.name = "ODataError";
    this.kind = opts.kind;
    this.status = opts.status;
    this.retryable = opts.retryable ?? RETRYABLE_DEFAULT[opts.kind];
    this.url = opts.url;
  }
}

const RETRYABLE_DEFAULT: Record<ODataErrorKind, boolean> = {
  auth: false,
  not_found: false,
  bad_request: false,
  rate_limit: true,
  server: true,
  timeout: true,
  network: true,
  parse: false,
  unknown: false,
};

/** Строит ODataError из HTTP-статуса и (опционально) тела ответа. */
export function fromHttpStatus(
  status: number,
  url: string,
  body?: string,
): ODataError {
  const detail = extract1cMessage(body);
  const base = (msg: string, kind: ODataErrorKind) =>
    new ODataError({
      kind,
      status,
      url,
      message: detail ? `${msg}: ${detail}` : msg,
    });

  if (status === 401 || status === 403) {
    return base(
      "Ошибка авторизации (401/403). Проверьте логин/пароль и что объект включён в «Состав OData»",
      "auth",
    );
  }
  if (status === 404) {
    return base(
      "Объект не найден (404). Вероятно, не добавлен в «Настройка стандартного интерфейса OData → Состав»",
      "not_found",
    );
  }
  if (status === 400) return base("Некорректный запрос к OData (400)", "bad_request");
  if (status === 429) return base("Слишком много запросов (429)", "rate_limit");
  if (status >= 500) return base(`Ошибка сервера 1С (${status})`, "server");
  return base(`Неожиданный ответ OData (${status})`, "unknown");
}

/** Пытается достать человекочитаемое сообщение из тела ошибки OData (JSON или XML). */
function extract1cMessage(body?: string): string | undefined {
  if (!body) return undefined;
  try {
    const json = JSON.parse(body) as {
      error?: { message?: { value?: string } | string };
    };
    const m = json.error?.message;
    if (typeof m === "string") return m;
    if (m?.value) return m.value;
  } catch {
    // не JSON — пробуем грубо выдрать <message> из XML
    const match = body.match(/<message[^>]*>([^<]+)<\/message>/i);
    if (match?.[1]) return match[1].trim();
  }
  return undefined;
}
