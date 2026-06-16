/**
 * Проверка инструментов записи БЕЗ фактической записи:
 *   - dry-run отдаёт предпросмотр (ничего не пишет);
 *   - confirm=true на незаписываемой базе блокируется предохранителем.
 * Запуск: node --env-file=.env dist/scripts/probe-write.js
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

type ToolText = { content: Array<{ type: string; text?: string }>; isError?: boolean };
const line = (s = ""): void => void process.stdout.write(`${s}\n`);
const textOf = (r: ToolText): string =>
  `${r.isError ? "[isError] " : ""}${r.content.map((c) => (c.type === "text" ? (c.text ?? "") : "")).join("\n")}`;

async function main(): Promise<void> {
  const env: Record<string, string> = {};
  for (const [k, v] of Object.entries(process.env)) if (v !== undefined) env[k] = v;

  const transport = new StdioClientTransport({ command: "node", args: ["dist/index.js"], env });
  const client = new Client({ name: "probe-write", version: "0.1.0" });
  await client.connect(transport);

  const { tools } = await client.listTools();
  line(`✓ tools/list: ${tools.length} инструментов`);
  const writeTools = tools.filter((t) => t.name.startsWith("create_")).map((t) => t.name);
  line(`  инструменты записи: ${writeTools.join(", ")}`);

  const args = (extra: object) => ({
    name: "ТЕСТ MCP — проверка (можно удалить)",
    inn: "7700000000",
    ...extra,
  });

  line("\n— create_counterparty (confirm=false, DRY-RUN, записи быть НЕ должно) —");
  line(
    textOf(
      (await client.callTool({
        name: "create_counterparty",
        arguments: args({ confirm: false }),
      })) as ToolText,
    ),
  );

  line("\n— create_counterparty (confirm=true на НЕзаписываемой базе → ожидаем блок предохранителя) —");
  line(
    textOf(
      (await client.callTool({
        name: "create_counterparty",
        arguments: args({ confirm: true }),
      })) as ToolText,
    ),
  );

  await client.close();
  line("\n✓ проверка завершена (фактических записей не выполнялось)");
}

main().catch((e) => {
  process.stderr.write(`probe-write error: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
