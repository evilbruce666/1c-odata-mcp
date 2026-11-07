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
  | "busy" // не дождались места под запрос (ODATA_MAX_CONCURRENT_REQUESTS) — запрос не отправлялся
  | "cancelled" // клиент отменил вызов инструмента — повторять нельзя, ответ уже никому не нужен
  | "deadline" // вызов не уложился в ODATA_TOOL_DEADLINE_MS — повторять внутри вызова незачем
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
  busy: false,
  cancelled: false,
  deadline: false,
  unknown: false,
};

/** Строит ODataError из HTTP-статуса и (опционально) тела ответа. */
export function fromHttpStatus(
  status: number,
  url: string,
  body?: string,
  request?: { method: string; path: string },
): ODataError {
  const detail = extract1cMessage(body);
  const base = (msg: string, kind: ODataErrorKind) =>
    new ODataError({
      kind,
      status,
      url,
      message: detail ? `${msg}: ${detail}` : msg,
    });

  const authError = (message: string) =>
    new ODataError({
      kind: "auth",
      status,
      url,
      // Ответ сервера может содержать данные/ключи — при отказе не цитируем его.
      message: `${message}${request ? safeRequestContext(request) : ""}`,
    });
  if (status === 401 || status === 403) {
    return authError(`Ошибка авторизации (${status}). Проверьте учётную запись и доступ в 1С.`);
  }
  if (status === 404) {
    // «Экземпляр сущности не найден» — набор опубликован, нет записи с таким ключом;
    // подсказка про «Состав» здесь только сбила бы с толку.
    if (detail && /Экземпляр сущности не найден/i.test(detail)) {
      return base("Объект не найден (404): записи с таким ключом нет", "not_found");
    }
    return base(
      "Объект не найден (404). Вероятно, не добавлен в «Настройка стандартного интерфейса OData → Состав»",
      "not_found",
    );
  }
  if (status === 400) return base("Некорректный запрос к OData (400)", "bad_request");
  if (status === 429) return base("Слишком много запросов (429)", "rate_limit");
  // 1С отвечает 500 «Нарушение прав доступа!», когда у пользователя OData нет прав на объект или
  // действие. Права проверяются до записи — это однозначный отказ, а не сбой с неизвестным исходом.
  if (status >= 500 && detail && /Нарушение прав доступа/i.test(detail)) {
    return authError(`Нет прав в 1С (${status}): Нарушение прав доступа!`);
  }
  // Ошибки самого запроса, которые 1С отдаёт как 500: повтор того же запроса даст тот же отказ,
  // поэтому не повторяем (лишние запросы и паузы, #53). Вид остаётся server — журнал записи
  // классифицирует такие отказы как раньше.
  if (status >= 500 && detail && DETERMINISTIC_500.some((re) => re.test(detail)))
    return new ODataError({
      kind: "server",
      status,
      url,
      retryable: false,
      message: `Ошибка сервера 1С (${status}): ${detail}. Запрос в таком виде 1С не выполняет — повтор не поможет`,
    });
  if (status >= 500) return base(`Ошибка сервера 1С (${status})`, "server");
  return base(`Неожиданный ответ OData (${status})`, "unknown");
}

/** Тексты 500-ответов 1С об ошибке в самом запросе (условие, сравнение типов) — не временные. */
const DETERMINISTIC_500: readonly RegExp[] = [
  /Операция не разрешена в предложении/i,
  /Нельзя сравнивать поля неограниченной длины/i,
];

/** Пытается достать человекочитаемое сообщение из тела ошибки OData (JSON или XML). */
function extract1cMessage(body?: string): string | undefined {
  if (!body) return undefined;
  try {
    // OData v3 (1С) кладёт ошибку в ключ "odata.error", v2/v4 — в "error".
    const json = JSON.parse(body) as Record<string, { message?: { value?: string } | string }>;
    const m = (json["odata.error"] ?? json["error"])?.message;
    if (typeof m === "string") return m;
    if (m?.value) return m.value;
  } catch {
    // не JSON — пробуем грубо выдрать <message> из XML
    const match = body.match(/<message[^>]*>([^<]+)<\/message>/i);
    if (match?.[1]) return match[1].trim();
  }
  return undefined;
}

/** Только метод, известное действие и ограниченное имя: без URL, ключей и query. */
function safeRequestContext({ method, path }: { method: string; path: string }): string {
  const operation = ["GET", "HEAD", "POST", "PATCH", "PUT", "DELETE"].includes(method) ? method : "HTTP";
  let decoded: string;
  try {
    decoded = decodeURIComponent(path.split("?")[0]!);
  } catch {
    return ` Отказ получен при операции ${operation}.`;
  }
  const entitySet = decoded.split(/[(/]/)[0]!;
  const valid = entitySet === "$metadata" || /^[A-Za-zА-Яа-яЁё_][A-Za-z0-9А-Яа-яЁё_]{0,199}$/.test(entitySet);
  const action = method === "POST" ? /\)\/(Post|Unpost)$/.exec(decoded)?.[1] : undefined;
  return ` Отказ получен при операции ${operation}${action ? ` ${action}` : ""}${valid ? ` над объектом ${entitySet}` : ""}.`;
}
