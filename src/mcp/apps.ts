import { readFileSync } from "node:fs";
import type { McpServer } from "@modelcontextprotocol/server";

/**
 * MCP Apps: инструмент со ссылкой `_meta.ui.resourceUri` на `ui://`-ресурс — HTML, который клиент
 * рисует прямо в ответе (в изолированном iframe) и которому передаёт результат инструмента.
 * Клиент без поддержки Apps просто показывает текстовый ответ — он не меняется.
 *
 * HTML лежит в apps/ в корне пакета: самодостаточный, без внешних загрузок (данные 1С не уходят
 * никуда), все строки из 1С выводятся только через textContent — названия контрагентов и
 * документов вводят люди.
 *
 * Известное ограничение: Claude (вкладка Code в приложении) не рисует карточку, если в имени
 * инструмента есть точка, — https://github.com/anthropics/claude-code/issues/104045. Наши имена
 * с точками, поэтому там пока виден только текстовый ответ. Другие клиенты MCP Apps (проверено в
 * MCPJam Inspector) точку в имени принимают.
 */

export const APP_MIME_TYPE = "text/html;profile=mcp-app";
export const APPS_EXTENSION_ID = "io.modelcontextprotocol/ui";

export interface AppBinding {
  tools: string[];
  uri: string;
  file: string;
  title: string;
}

export const APPS: AppBinding[] = [
  {
    tools: ["read.analytics.get_debtors", "read.analytics.get_creditors"],
    uri: "ui://1c-odata/settlements.html",
    file: "settlements.html",
    title: "Долги по срокам",
  },
];

/** `_meta` для инструмента с приложением; вложенный ключ по спецификации и плоский — для ранних клиентов. */
export function appMetaFor(tool: string): Record<string, unknown> | undefined {
  const app = APPS.find((a) => a.tools.includes(tool));
  return app ? { ui: { resourceUri: app.uri }, "ui/resourceUri": app.uri } : undefined;
}

/** HTML приложения: из dist/mcp/ и из src/mcp/ папка apps/ — на два уровня выше. */
export function loadAppHtml(file: string): string {
  return readFileSync(new URL(`../../apps/${file}`, import.meta.url), "utf8");
}

export function registerApps(server: McpServer): void {
  for (const app of APPS) {
    server.registerResource(
      app.file.replace(/\.html$/, ""),
      app.uri,
      { title: app.title, mimeType: APP_MIME_TYPE },
      async (uri) => ({
        contents: [{ uri: uri.href, mimeType: APP_MIME_TYPE, text: loadAppHtml(app.file) }],
      }),
    );
  }
}
