import { z } from "zod";
import { AccountingSchema, type AccountingConfig } from "./accounting.js";
import { homedir } from "node:os";
import { resolve, join } from "node:path";

/**
 * Конфигурация сервера. Поддерживает несколько баз 1С одновременно.
 *
 * Способы задать базы в .env:
 *  1) Одна база (обратная совместимость):
 *       ODATA_BASE_URL / ODATA_USERNAME / ODATA_PASSWORD  → база "default"
 *  2) Несколько баз — индексируются по <ИМЯ> в названии переменной:
 *       ODATA_DB_BUH_BASE_URL / ODATA_DB_BUH_USERNAME / ODATA_DB_BUH_PASSWORD
 *       ODATA_DB_BUH_LABEL   (необязательно — человекочитаемое название)
 *       ODATA_DB_TORG_BASE_URL / ...
 *       ODATA_DEFAULT_DB=buh (необязательно; иначе первая по алфавиту)
 */

export interface ConnectionConfig {
  name: string;
  label?: string;
  baseUrl: string;
  username: string;
  password: string;
  /** Разрешена ли запись в эту базу (POST/PATCH). По умолчанию false. */
  writable: boolean;
  accounting?: AccountingConfig;
  /** Явное значение только для этой базы; иначе используется общая настройка. */
  regulatedCurrencyCode?: string;
}

export interface Behavior {
  timeoutMs: number;
  retries: number;
  pageSize: number;
  maxRows: number;
  /** Потолок строк для агрегаторов (они только суммируют, строки наружу не отдают). */
  analyticsMaxRows: number;
  readOnly: boolean;
  /** Durable local state for idempotent confirmed document creates. */
  writeJournalDir: string;
  /**
   * Дописывать «[op:<operationId>]» в «Комментарий» создаваемых документов. По умолчанию выключено:
   * сверка идёт по назначенному Ref_Key; метка — только для тех, кому нужна видимая в 1С привязка.
   */
  writeOperationMarker: boolean;
  /**
   * Читаемый вид (`view`) в предпросмотре записи: названия вместо GUID — до 4 запросов к справочникам.
   * По умолчанию включён: запись подтверждает человек. false — для интеграций с массовой записью без
   * человека в цикле; машинная часть предпросмотра не меняется.
   */
  previewView?: boolean;
  /** Предел одновременных HTTP-запросов одного подключения; 0 — без предела (по умолчанию). */
  maxConcurrentRequests: number;
  /** Сколько ждать свободного места под запрос, мс. */
  queueTimeoutMs: number;
  /** Общий лимит времени на вызов инструмента чтения, мс; 0 — без лимита (по умолчанию). */
  toolDeadlineMs: number;
  /** Предел текста ответа инструмента чтения, символов: только уменьшение нормы для клиентов с малым контекстом. */
  maxResponseChars: number;
  /** Сервер доступен по сети: ответы не раскрывают адрес базы. */
  networkTransport?: boolean;
  /** Какие инструменты видит клиент: all — все, read — только чтение. Запись защищена не этим, а READ_ONLY. */
  toolsProfile: "all" | "read";
  /** Точный список инструментов вместо профиля. */
  toolsAllow?: string[] | undefined;
  /**
   * Код валюты регламентированного учёта (по справочнику «Валюты»): по нему договор относится к виду
   * расчётов «в валюте регламентированного учёта» при подборе счетов расчётов. По умолчанию 643 (рубль).
   */
  regulatedCurrencyCode?: string;
  regulatedCurrencySource?: "global" | "default";
}

/** Настройки HTTP-транспорта. Без них сервер работает через stdio. */
export interface HttpTransport {
  host: string;
  port: number;
  token: string;
  /** Допустимые имена в заголовках Host и Origin; без списка — только loopback. */
  allowedHosts?: string[] | undefined;
}

export type LogLevel = "trace" | "debug" | "info" | "warn" | "error" | "fatal";

export interface RuntimeConfig {
  connections: ConnectionConfig[];
  defaultName: string;
  behavior: Behavior;
  logLevel: LogLevel;
  http?: HttpTransport | undefined;
}

const normalizeUrl = (u: string): string => (u.endsWith("/") ? u : `${u}/`);

// Catalog Code is a string, not necessarily an ISO numeric code. Blank means unset.
const CurrencyCodeSchema = z.preprocess(
  (value) => (typeof value === "string" && value.trim() === "" ? undefined : value),
  z
    .string()
    .trim()
    .min(1)
    .max(100)
    .refine(
      (value) => [...value].every((char) => char.charCodeAt(0) >= 32 && char.charCodeAt(0) !== 127),
      "не должен содержать управляющие символы",
    )
    .optional(),
);

const BehaviorSchema = z.object({
  ODATA_TIMEOUT_MS: z.coerce.number().int().positive().default(30_000),
  ODATA_RETRIES: z.coerce.number().int().min(0).max(10).default(3),
  ODATA_PAGE_SIZE: z.coerce.number().int().positive().max(5_000).default(100),
  ODATA_MAX_ROWS: z.coerce.number().int().positive().max(100_000).default(1_000),
  ODATA_ANALYTICS_MAX_ROWS: z.coerce.number().int().positive().max(1_000_000).default(200_000),
  ODATA_WRITE_JOURNAL_DIR: z.string().min(1).optional(),
  ODATA_WRITE_OPERATION_MARKER: z.enum(["true", "false"]).default("false"),
  ODATA_PREVIEW_VIEW: z.enum(["true", "false"]).default("true"),
  // Предел одновременных HTTP-запросов к одной базе (0 — без предела) и сколько ждать места.
  ODATA_MAX_CONCURRENT_REQUESTS: z.coerce.number().int().min(0).max(64).default(0),
  ODATA_QUEUE_TIMEOUT_MS: z.coerce.number().int().positive().default(30_000),
  ODATA_TOOL_DEADLINE_MS: z.coerce.number().int().min(0).default(0),
  ODATA_MAX_RESPONSE_CHARS: z.coerce.number().int().min(10_000).max(45_000).default(45_000),
  TOOLS_PROFILE: z.preprocess((v) => (v === "" ? undefined : v), z.enum(["all", "read"]).optional()),
  TOOLS_ALLOW: z.string().optional(),
  MCP_TRANSPORT: z.preprocess((v) => (v === "" ? undefined : v), z.enum(["stdio", "http"]).default("stdio")),
  MCP_HTTP_HOST: z.preprocess((v) => (v === "" ? undefined : v), z.string().default("127.0.0.1")),
  MCP_HTTP_PORT: z.preprocess(
    (v) => (v === "" ? undefined : v),
    z.coerce.number().int().min(0).max(65_535).default(8787),
  ),
  MCP_HTTP_TOKEN: z.string().optional(),
  MCP_HTTP_ALLOWED_HOSTS: z.string().optional(),
  // Пустое значение (обычное в шаблонах .env) = «не задано», а не ошибка старта.
  ODATA_REGULATED_CURRENCY_CODE: CurrencyCodeSchema,
  LOG_LEVEL: z.enum(["trace", "debug", "info", "warn", "error", "fatal"]).default("info"),
  READ_ONLY: z.enum(["true", "false"]).default("true"),
});

const nonblank = (value: string | undefined): string | undefined => value?.trim() || undefined;
const JsonAccountingSchema = z.preprocess((value, ctx) => {
  if (typeof value !== "string") return value;
  if (!value.trim()) return undefined;
  try {
    return JSON.parse(value);
  } catch {
    ctx.addIssue({ code: "custom", message: "должна быть корректным JSON настройки плана счетов" });
    return z.NEVER;
  }
}, AccountingSchema.optional());

const ConnSchema = z.object({
  accounting: JsonAccountingSchema,
  regulatedCurrencyCode: CurrencyCodeSchema,
  name: z.string().min(1),
  label: z.string().min(1).optional(),
  baseUrl: z.string().url("должен быть валидным URL").transform(normalizeUrl),
  username: z.string().min(1, "обязателен"),
  password: z.string().min(1, "обязателен"),
  writable: z.preprocess((v) => v === true || v === "true", z.boolean()).default(false),
});

type Env = Record<string, string | undefined>;

/** Собирает «сырые» описания баз из переменных окружения. */
function collectRawConnections(env: Env): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = [];

  // Индексированные базы: ODATA_DB_<ИМЯ>_BASE_URL
  for (const key of Object.keys(env)) {
    const m = /^ODATA_DB_(.+)_BASE_URL$/.exec(key);
    if (!m) continue;
    const tag = m[1] as string;
    out.push({
      name: tag.toLowerCase(),
      label: env[`ODATA_DB_${tag}_LABEL`],
      baseUrl: env[key],
      username: env[`ODATA_DB_${tag}_USERNAME`],
      password: env[`ODATA_DB_${tag}_PASSWORD`],
      writable: env[`ODATA_DB_${tag}_WRITABLE`],
      regulatedCurrencyCode: env[`ODATA_DB_${tag}_REGULATED_CURRENCY_CODE`],
      accounting: nonblank(env[`ODATA_DB_${tag}_ACCOUNTING_CONFIG`]) ?? nonblank(env.ODATA_ACCOUNTING_CONFIG),
    });
  }

  // Обратная совместимость: одиночная база.
  if (out.length === 0 && env.ODATA_BASE_URL) {
    out.push({
      name: (env.ODATA_DB_NAME ?? "default").toLowerCase(),
      label: env.ODATA_DB_LABEL,
      baseUrl: env.ODATA_BASE_URL,
      username: env.ODATA_USERNAME,
      password: env.ODATA_PASSWORD,
      writable: env.ODATA_WRITABLE,
      accounting: nonblank(env.ODATA_ACCOUNTING_CONFIG),
    });
  }

  return out;
}

let cached: RuntimeConfig | undefined;

export function loadConfig(): RuntimeConfig {
  if (cached) return cached;
  cached = parseConfig(process.env as Env);
  return cached;
}

const LOOPBACK_HOSTS = new Set(["127.0.0.1", "localhost", "::1"]);
const LOOPBACK_URL_HOSTS = new Set(["127.0.0.1", "localhost", "[::1]"]);
const MIN_TOKEN_LENGTH = 32;

/**
 * Требования к сетевому режиму. Отказ при старте, а не предупреждение: сервер с паролем 1С без токена доступен
 * всем в сети, а запись в этом режиме не защищена. Журнал без тома теряется вместе с контейнером и повтор создаёт
 * дубль, а держатель токена неотличим от человека, подтверждающего запись.
 */
function httpIssues(b: z.infer<typeof BehaviorSchema>): string[] {
  if (b.MCP_TRANSPORT !== "http") return [];
  const issues: string[] = [];
  if ((b.MCP_HTTP_TOKEN?.length ?? 0) < MIN_TOKEN_LENGTH) {
    issues.push(`MCP_HTTP_TOKEN обязателен при MCP_TRANSPORT=http, не короче ${MIN_TOKEN_LENGTH} символов`);
  }
  if (b.READ_ONLY !== "true") {
    issues.push("MCP_TRANSPORT=http работает только для чтения: задайте READ_ONLY=true");
  }
  if (!LOOPBACK_HOSTS.has(b.MCP_HTTP_HOST) && !toolList(b.MCP_HTTP_ALLOWED_HOSTS)) {
    issues.push(
      `MCP_HTTP_ALLOWED_HOSTS обязателен, если MCP_HTTP_HOST=${b.MCP_HTTP_HOST}: перечислите имена, по которым клиент обращается к серверу`,
    );
  }
  return issues;
}

/** Список имён через запятую; пустое значение (обычное в шаблонах .env) — «не задано». */
const toolList = (value: string | undefined): string[] | undefined => {
  const names = (value ?? "")
    .split(",")
    .map((n) => n.trim())
    .filter(Boolean);
  return names.length ? names : undefined;
};

/**
 * Базы с нешифрованным адресом не на этой машине: логин и пароль 1С идут по такому каналу открытым текстом.
 * Только имена: адрес в журнал не нужен. Локальная сеть может быть защищена, поэтому это предупреждение, а не отказ.
 */
export function unencryptedDatabases(connections: readonly ConnectionConfig[]): string[] {
  return connections
    .filter((c) => {
      const { protocol, hostname } = new URL(c.baseUrl);
      return protocol === "http:" && !LOOPBACK_URL_HOSTS.has(hostname);
    })
    .map((c) => c.name);
}

/** Чистый разбор конфигурации из объекта env (без кеша) — удобно для тестов. */
export function parseConfig(env: Env): RuntimeConfig {
  const issues: string[] = [];

  const behaviorParsed = BehaviorSchema.safeParse(env);
  if (!behaviorParsed.success) {
    for (const i of behaviorParsed.error.issues) issues.push(`${i.path.join(".")}: ${i.message}`);
  }

  if (
    behaviorParsed.success &&
    behaviorParsed.data.TOOLS_PROFILE &&
    toolList(behaviorParsed.data.TOOLS_ALLOW)
  ) {
    issues.push("TOOLS_PROFILE и TOOLS_ALLOW вместе не задаются: оставьте одну из настроек");
  }

  if (behaviorParsed.success) issues.push(...httpIssues(behaviorParsed.data));

  const raw = collectRawConnections(env);
  if (raw.length === 0) {
    issues.push(
      "не задана ни одна база: укажите ODATA_BASE_URL (одна база) или ODATA_DB_<ИМЯ>_BASE_URL (несколько)",
    );
  }

  const connections: ConnectionConfig[] = [];
  const seen = new Set<string>();
  for (const r of raw) {
    const parsed = ConnSchema.safeParse(r);
    if (!parsed.success) {
      const who = typeof r.name === "string" ? r.name : "?";
      for (const i of parsed.error.issues) issues.push(`база "${who}": ${i.path.join(".")} ${i.message}`);
      continue;
    }
    if (seen.has(parsed.data.name)) {
      issues.push(`дублирующееся имя базы: "${parsed.data.name}"`);
      continue;
    }
    seen.add(parsed.data.name);
    connections.push(parsed.data);
  }

  if (issues.length > 0) {
    throw new Error(`Некорректная конфигурация (.env):\n${issues.map((s) => `  • ${s}`).join("\n")}`);
  }

  const b = behaviorParsed.success ? behaviorParsed.data : BehaviorSchema.parse({});
  const wantDefault = env.ODATA_DEFAULT_DB?.toLowerCase();
  const sortedNames = connections.map((c) => c.name).sort();
  const defaultName = wantDefault && seen.has(wantDefault) ? wantDefault : (sortedNames[0] as string);

  return {
    connections,
    defaultName,
    logLevel: b.LOG_LEVEL,
    http:
      b.MCP_TRANSPORT === "http"
        ? {
            host: b.MCP_HTTP_HOST,
            port: b.MCP_HTTP_PORT,
            token: b.MCP_HTTP_TOKEN!,
            allowedHosts: toolList(b.MCP_HTTP_ALLOWED_HOSTS),
          }
        : undefined,
    behavior: {
      timeoutMs: b.ODATA_TIMEOUT_MS,
      retries: b.ODATA_RETRIES,
      pageSize: b.ODATA_PAGE_SIZE,
      maxRows: b.ODATA_MAX_ROWS,
      analyticsMaxRows: b.ODATA_ANALYTICS_MAX_ROWS,
      readOnly: b.READ_ONLY === "true",
      writeOperationMarker: b.ODATA_WRITE_OPERATION_MARKER === "true",
      previewView: b.ODATA_PREVIEW_VIEW === "true",
      maxConcurrentRequests: b.ODATA_MAX_CONCURRENT_REQUESTS,
      queueTimeoutMs: b.ODATA_QUEUE_TIMEOUT_MS,
      toolDeadlineMs: b.ODATA_TOOL_DEADLINE_MS,
      maxResponseChars: b.ODATA_MAX_RESPONSE_CHARS,
      networkTransport: b.MCP_TRANSPORT === "http",
      toolsProfile: b.TOOLS_PROFILE ?? "all",
      toolsAllow: toolList(b.TOOLS_ALLOW),
      regulatedCurrencyCode: b.ODATA_REGULATED_CURRENCY_CODE ?? "643",
      regulatedCurrencySource: b.ODATA_REGULATED_CURRENCY_CODE ? "global" : "default",
      writeJournalDir: resolve(
        b.ODATA_WRITE_JOURNAL_DIR ?? join(homedir(), ".1c-odata-mcp", "write-journal"),
      ),
    },
  };
}
