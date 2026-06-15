/**
 * РАЗОВАЯ живая проверка цикла записи: создать тестового контрагента → сразу
 * пометить на удаление. Пишет в боевую базу, поэтому имя — явно тестовое.
 * Запуск только через временный env с READ_ONLY=false (см. вызывающий скрипт).
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

type ToolText = { content: Array<{ type: string; text?: string }>; isError?: boolean };
const line = (s = ""): void => void process.stdout.write(`${s}\n`);
const textOf = (r: ToolText): string =>
  r.content.map((c) => (c.type === "text" ? c.text ?? "" : "")).join("\n");

async function main(): Promise<void> {
  const db = process.env.PROBE_DB ?? "ip";
  const env: Record<string, string> = {};
  for (const [k, v] of Object.entries(process.env)) if (v !== undefined) env[k] = v;

  const transport = new StdioClientTransport({ command: "node", args: ["dist/index.js"], env });
  const client = new Client({ name: "probe-write-live", version: "0.1.0" });
  await client.connect(transport);

  line(`— create_counterparty (база ${db}, confirm=true) —`);
  const created = (await client.callTool({
    name: "create_counterparty",
    arguments: { database: db, name: "ТЕСТ MCP — удалить", inn: "7700000000", confirm: true },
  })) as ToolText;
  const createdText = textOf(created);
  line(created.isError ? `[isError] ${createdText}` : createdText);

  const ref = (() => {
    try {
      return (JSON.parse(createdText) as { ref?: string }).ref;
    } catch {
      return undefined;
    }
  })();

  if (!ref) {
    line("\nRef не получен — пометку на удаление пропускаю.");
    await client.close();
    return;
  }

  line(`\n— mark_for_deletion (ref ${ref}, confirm=true) —`);
  const marked = (await client.callTool({
    name: "mark_for_deletion",
    arguments: { database: db, entitySet: "Catalog_Контрагенты", ref, mark: true, confirm: true },
  })) as ToolText;
  line(marked.isError ? `[isError] ${textOf(marked)}` : textOf(marked));

  await client.close();
  line("\n✓ цикл создать→пометить-на-удаление выполнен");
}

main().catch((e) => {
  process.stderr.write(`probe-write-live error: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
