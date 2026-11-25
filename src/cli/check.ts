/**
 * `1c-odata-mcp --check` — диагностика подключения из терминала до клиента: тот же разбор настроек, что у сервера, и по
 * каждой базе один GET $metadata. Печатает версию OData, число объектов, что это за конфигурация и каких объектов для
 * основных вопросов не хватает в «Составе OData». Ничего не пишет ни в 1С, ни в файлы; пароль не печатает.
 * Настройки у новичка лежат не в терминале, а в блоке env конфига приложения — его читает --config.
 */
import { readFileSync } from "node:fs";
import { parseConfig, unencryptedDatabases } from "../config/env.js";
import { accountingChart, accountingRegister } from "../config/accounting.js";
import { CATALOGS, REGISTERS } from "../config/mapping.js";
import { ServerContext, type Connection } from "../context.js";
import { ODataError } from "../odata/errors.js";

type Env = Record<string, string | undefined>;
type Print = (line: string) => void;

/** Регистры и справочники ключей, без которых в Управлении торговлей не работают остатки, продажи и расчёты. */
export const TRADE_REQUIRED = [
  "AccumulationRegister_ТоварыНаСкладах",
  "AccumulationRegister_ВыручкаИСебестоимостьПродаж",
  "AccumulationRegister_РасчетыСКлиентамиПоСрокам",
  "AccumulationRegister_РасчетыСПоставщикамиПоСрокам",
  "Catalog_КлючиАналитикиУчетаНоменклатуры",
  "Catalog_КлючиАналитикиУчетаПоПартнерам",
];
export const TRADE_CASH = [
  "AccumulationRegister_ДенежныеСредстваБезналичные",
  "AccumulationRegister_ДенежныеСредстваНаличные",
  "AccumulationRegister_ДенежныеСредстваВКассахККМ",
  "AccumulationRegister_ДенежныеСредстваВПути",
];
const COMMON_REQUIRED = [CATALOGS.counterparties[0], CATALOGS.nomenclature[0], CATALOGS.organizations[0]];

export interface Diagnosis {
  kind: "accounting" | "trade" | "both" | "other" | "empty";
  /** Объекты для основных вопросов, которых в составе нет. На код возврата не влияют. */
  missing: string[];
}

export function diagnose(conn: Connection, published: ReadonlySet<string>): Diagnosis {
  if (!published.size) return { kind: "empty", missing: [] };
  const accounting = accountingRegister(conn.cfg).some((r) => published.has(r));
  const trade = published.has(REGISTERS.stock[0]) || published.has(TRADE_REQUIRED[1]!);
  const kind = accounting && trade ? "both" : accounting ? "accounting" : trade ? "trade" : "other";
  const want = [
    ...COMMON_REQUIRED,
    ...(accounting ? accountingChart(conn.cfg) : []),
    ...(trade ? TRADE_REQUIRED : []),
  ];
  const missing = want.filter((n) => !published.has(n));
  if (trade && !TRADE_CASH.some((n) => published.has(n))) missing.push(TRADE_CASH[0]!);
  return { kind, missing };
}

const KIND_TEXT: Record<Diagnosis["kind"], string> = {
  accounting: "есть регистр бухгалтерии: работают инструменты Бухгалтерии",
  trade: "оперативные регистры без бухгалтерских: работают инструменты Управления торговлей",
  both: "есть и бухгалтерский регистр, и оперативные (КА, ERP или доработанная база)",
  other: "регистров для аналитики не видно: доступны документы, справочники и схема",
  empty:
    "«Состав OData» пуст: добавьте объекты в 1С («Все функции» → «Настройка стандартного интерфейса OData»)",
};

/** Что делать при отказе — одной строкой, по виду ошибки. */
function advice(e: unknown): string {
  if (!(e instanceof ODataError)) return "";
  if (e.kind === "auth")
    return "Проверьте логин и пароль пользователя OData; в облаке часто нужна учётная запись площадки, а не пользователь 1С.";
  if (e.kind === "not_found")
    return "Проверьте адрес: он должен заканчиваться на /odata/standard.odata/ и открываться в браузере.";
  if (e.kind === "network")
    return "Проверьте, что сервер 1С доступен с этого компьютера (адрес, VPN, сертификат).";
  if (e.kind === "timeout") return "1С не ответила вовремя: проверьте сеть или увеличьте ODATA_TIMEOUT_MS.";
  if (e.kind === "parse")
    return "Ответ не похож на OData 1С: проверьте, что адрес ведёт на публикацию OData.";
  return "";
}

/** Описание базы одним запросом; true — база отвечает и состав не пуст. */
async function checkBase(ctx: ServerContext, name: string, print: Print): Promise<boolean> {
  const conn = ctx.db(name);
  print(`База «${name}»${conn.cfg.label ? ` (${conn.cfg.label})` : ""}: ${conn.cfg.baseUrl}`);
  try {
    const meta = await conn.client.refreshMetadata();
    const published = new Set(meta.entities.keys());
    const d = diagnose(conn, published);
    print(`  подключение: есть, OData ${meta.odataVersion}, объектов в составе: ${published.size}`);
    print(`  конфигурация: ${KIND_TEXT[d.kind]}`);
    if (d.missing.length)
      print(
        `  нет в «Составе OData» (нужны для части вопросов): ${d.missing.map((m) => m.replace("_", " ")).join("; ")}`,
      );
    return d.kind !== "empty";
  } catch (e) {
    print(`  не получилось: ${e instanceof Error ? e.message : String(e)}`);
    const hint = advice(e);
    if (hint) print(`  ${hint}`);
    return false;
  }
}

/**
 * env из конфига приложения (claude_desktop_config.json, mcp.json Cursor/VS Code): записи mcpServers или servers,
 * которые запускают этот коннектор. Текст ошибки разбора файла не печатается: в нём фрагмент с паролем.
 */
export function readClientConfig(path: string): Array<{ name: string; env: Env }> {
  let text: string;
  try {
    text = readFileSync(path, "utf8");
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    throw new Error(`не удалось открыть файл ${path}${code ? ` (${code})` : ""}`, { cause: e });
  }
  let json: unknown;
  try {
    json = JSON.parse(text.replace(/^\uFEFF/, ""));
  } catch {
    throw new Error(
      `файл ${path} не разбирается как JSON (частая причина — лишняя запятая или комментарий). Содержимое не показано: в нём пароль.`,
    );
  }
  const root = (json ?? {}) as Record<string, unknown>;
  const servers = (root["mcpServers"] ?? root["servers"]) as Record<string, unknown> | undefined;
  if (!servers || typeof servers !== "object")
    throw new Error(`в файле ${path} нет раздела mcpServers (или servers, как в VS Code)`);
  const found: Array<{ name: string; env: Env }> = [];
  for (const [name, raw] of Object.entries(servers)) {
    const s = (raw ?? {}) as { command?: unknown; args?: unknown; env?: unknown };
    const call = [s.command, ...(Array.isArray(s.args) ? s.args : [])].map(String).join(" ");
    if (!call.includes("1c-odata-mcp") && name !== "1c-odata" && name !== "1c-odata-mcp") continue;
    const env: Env = {};
    for (const [k, v] of Object.entries((s.env ?? {}) as Record<string, unknown>)) {
      if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") env[k] = String(v);
    }
    found.push({ name, env });
  }
  if (!found.length)
    throw new Error(
      `в файле ${path} нет записи, запускающей 1c-odata-mcp. Записи в нём: ${Object.keys(servers).join(", ") || "нет"}`,
    );
  return found;
}

/** Проверить настройки из одного источника; true — все базы отвечают. */
async function checkEnv(env: Env, print: Print): Promise<boolean> {
  const substitution = Object.entries(env).find(([k, v]) => k.startsWith("ODATA_") && v?.includes("${"));
  if (substitution) {
    print(
      `Значение ${substitution[0]} задано подстановкой «\${…}»: её разворачивает приложение, а не --check. ` +
        "Задайте переменные в терминале или проверьте из приложения: «проверь соединение с 1С».",
    );
    return false;
  }
  const cfg = parseConfig(env);
  const unencrypted = unencryptedDatabases(cfg.connections);
  if (unencrypted.length)
    print(`Внимание: адрес без HTTPS (${unencrypted.join(", ")}) — логин и пароль идут открытым текстом.`);
  const ctx = new ServerContext(cfg);
  let ok = true;
  for (const c of cfg.connections) ok = (await checkBase(ctx, c.name, print)) && ok;
  print(
    cfg.behavior.readOnly
      ? "Режим: только чтение."
      : `Режим: запись включена для баз: ${
          cfg.connections
            .filter((c) => c.writable)
            .map((c) => c.name)
            .join(", ") || "ни одной (нужен ODATA_WRITABLE)"
        }.`,
  );
  return ok;
}

/** 0 — все базы отвечают и состав не пуст; 1 — хотя бы одна нет или настройки не читаются. */
export async function runCheck(args: readonly string[], env: Env, print: Print): Promise<number> {
  const at = args.indexOf("--config");
  try {
    if (at >= 0) {
      const path = args[at + 1];
      if (!path || path.startsWith("--")) throw new Error("после --config укажите путь к конфигу приложения");
      let ok = true;
      for (const entry of readClientConfig(path)) {
        print(`Настройки из ${path}, запись «${entry.name}».`);
        ok = (await checkEnv(entry.env, print)) && ok;
      }
      return ok ? 0 : 1;
    }
    return (await checkEnv(env, print)) ? 0 : 1;
  } catch (e) {
    print(e instanceof Error ? e.message : String(e));
    return 1;
  }
}
