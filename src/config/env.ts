import { z } from "zod";

/**
 * Схема переменных окружения. Валидируется один раз при старте —
 * при некорректной конфигурации сервер падает сразу с понятным сообщением,
 * а не на первом запросе.
 */
const EnvSchema = z.object({
  ODATA_BASE_URL: z
    .string()
    .url("ODATA_BASE_URL должен быть валидным URL")
    .transform((u) => (u.endsWith("/") ? u : `${u}/`)),
  ODATA_USERNAME: z.string().min(1, "ODATA_USERNAME обязателен"),
  ODATA_PASSWORD: z.string().min(1, "ODATA_PASSWORD обязателен"),

  ODATA_TIMEOUT_MS: z.coerce.number().int().positive().default(30_000),
  ODATA_RETRIES: z.coerce.number().int().min(0).max(10).default(3),
  ODATA_PAGE_SIZE: z.coerce.number().int().positive().max(5_000).default(100),
  ODATA_MAX_ROWS: z.coerce.number().int().positive().max(100_000).default(1_000),

  LOG_LEVEL: z
    .enum(["trace", "debug", "info", "warn", "error", "fatal"])
    .default("info"),

  // Жёсткий режим только-чтение. На 1-м этапе должен быть true.
  READ_ONLY: z
    .enum(["true", "false"])
    .default("true")
    .transform((v) => v === "true"),
});

export type AppConfig = z.infer<typeof EnvSchema>;

let cached: AppConfig | undefined;

/**
 * Возвращает провалидированную конфигурацию (синглтон).
 * Бросает ZodError с человекочитаемым описанием при ошибке.
 */
export function loadConfig(): AppConfig {
  if (cached) return cached;

  const parsed = EnvSchema.safeParse(process.env);
  if (!parsed.success) {
    const issues = parsed.error.issues
      .map((i) => `  • ${i.path.join(".")}: ${i.message}`)
      .join("\n");
    throw new Error(`Некорректная конфигурация (.env):\n${issues}`);
  }

  cached = parsed.data;
  return cached;
}
