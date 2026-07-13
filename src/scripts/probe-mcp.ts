/**
 * End-to-end проверка MCP-сервера: поднимает dist/index.js по stdio и
 * общается с ним настоящим MCP-клиентом (как Claude Desktop).
 * Запуск: node --env-file=.env dist/scripts/probe-mcp.js
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

function line(s = ""): void {
  process.stdout.write(`${s}\n`);
}

type ToolText = { content: Array<{ type: string; text?: string }> };

function textOf(res: ToolText): string {
  return res.content.map((c) => (c.type === "text" ? (c.text ?? "") : "")).join("\n");
}

async function main(): Promise<void> {
  // Чистим env от undefined для типобезопасной передачи дочернему процессу.
  const env: Record<string, string> = {};
  for (const [k, v] of Object.entries(process.env)) if (v !== undefined) env[k] = v;

  const transport = new StdioClientTransport({
    command: "node",
    args: ["dist/index.js"],
    env,
  });
  const client = new Client({ name: "probe-mcp", version: "0.1.0" });
  await client.connect(transport);
  line("✓ connect + initialize");

  const { tools } = await client.listTools();
  line(`✓ tools/list: ${tools.length} инструментов`);
  line(`  ${tools.map((t) => t.name).join(", ")}`);

  const db = process.env.PROBE_DB; // имя базы для проверки выбора (необязательно)
  const dbArg = db ? { database: db } : {};

  line("\n— list_databases —");
  line(textOf((await client.callTool({ name: "read.system.list_databases", arguments: {} })) as ToolText));

  line(`\n— health_check ${db ? `(database=${db})` : ""} —`);
  line(textOf((await client.callTool({ name: "read.system.health_check", arguments: dbArg })) as ToolText));

  line("\n— list_organizations —");
  line(
    textOf((await client.callTool({ name: "read.system.list_organizations", arguments: dbArg })) as ToolText),
  );

  line("\n— get_debtors (top 3) —");
  line(
    textOf(
      (await client.callTool({
        name: "read.analytics.get_debtors",
        arguments: { ...dbArg, limit: 3 },
      })) as ToolText,
    ),
  );

  await client.close();
  line("\n✓ end-to-end ok");
}

main().catch((e) => {
  process.stderr.write(`MCP probe error: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
