#!/usr/bin/env node
/**
 * Точка входа. Флаги разбираются до импорта сервера: журнал открывает файл при импорте, а --help и --version не должны
 * ни читать настройки, ни что-то создавать. Без этих флагов — прежний сервер (serve.ts), прочие аргументы не мешают.
 */
import { readFileSync } from "node:fs";

const HELP = `1c-odata-mcp — MCP-сервер для 1С:Предприятие через OData.

Без аргументов запускается сервер (так его запускает MCP-клиент). Настройки — переменные окружения, см. README.

  --check                    проверить подключение к 1С по переменным окружения
  --check --config <файл>    то же по блоку env из конфига приложения (claude_desktop_config.json, mcp.json)
  --version                  версия
  --help                     эта справка
`;

const args = process.argv.slice(2);
if (args.includes("--version")) {
  const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
    version: string;
  };
  process.stdout.write(`${pkg.version}\n`);
} else if (args.includes("--help")) {
  process.stdout.write(HELP);
} else if (args.includes("--check")) {
  // Журнал диагностике не нужен: она сама печатает всё, что человеку нужно знать. Проверяются настройки как есть
  // (копия до подмены): с неверным LOG_LEVEL сервер не запустится, и проверка должна это сказать.
  const env = { ...process.env };
  process.env.LOG_LEVEL = "fatal";
  process.env.MCP_LOG_STDERR = "1";
  const { runCheck } = await import("./cli/check.js");
  process.exitCode = await runCheck(args, env, (line) => process.stdout.write(`${line}\n`));
} else {
  await import("./serve.js");
}
