import { z } from "zod";

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
}

export interface Behavior {
  timeoutMs: number;
  retries: number;
  pageSize: number;
  maxRows: number;
  readOnly: boolean;
}

export type LogLevel = "trace" | "debug" | "info" | "warn" | "error" | "fatal";

export interface RuntimeConfig {
  connections: ConnectionConfig[];
  defaultName: string;
  behavior: Behavior;
  logLevel: LogLevel;
}

const normalizeUrl = (u: string): string => (u.endsWith("/") ? u : `${u}/`);

const BehaviorSchema = z.object({
  ODATA_TIMEOUT_MS: z.coerce.number().int().positive().default(30_000),
  ODATA_RETRIES: z.coerce.number().int().min(0).max(10).default(3),
  ODATA_PAGE_SIZE: z.coerce.number().int().positive().max(5_000).default(100),
  ODATA_MAX_ROWS: z.coerce.number().int().positive().max(100_000).default(1_000),
  LOG_LEVEL: z.enum(["trace", "debug", "info", "warn", "error", "fatal"]).default("info"),
  READ_ONLY: z.enum(["true", "false"]).default("true"),
});

const ConnSchema = z.object({
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
    });
  }

  return out;
}

let cached: RuntimeConfig | undefined;

export function loadConfig(): RuntimeConfig {
  if (cached) return cached;

  const env = process.env as Env;
  const issues: string[] = [];

  const behaviorParsed = BehaviorSchema.safeParse(env);
  if (!behaviorParsed.success) {
    for (const i of behaviorParsed.error.issues) issues.push(`${i.path.join(".")}: ${i.message}`);
  }

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
  const defaultName =
    wantDefault && seen.has(wantDefault) ? wantDefault : (sortedNames[0] as string);

  cached = {
    connections,
    defaultName,
    logLevel: b.LOG_LEVEL,
    behavior: {
      timeoutMs: b.ODATA_TIMEOUT_MS,
      retries: b.ODATA_RETRIES,
      pageSize: b.ODATA_PAGE_SIZE,
      maxRows: b.ODATA_MAX_ROWS,
      readOnly: b.READ_ONLY === "true",
    },
  };
  return cached;
}
