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
  return res.content.map((c) => (c.type === "text" ? c.text ?? "" : "")).join("\n");
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

  line("\n— health_check —");
  line(textOf((await client.callTool({ name: "health_check", arguments: {} })) as ToolText));

  line("\n— find_counterparty (ВК) —");
  line(textOf((await client.callTool({ name: "find_counterparty", arguments: { query: "ВК", limit: 2 } })) as ToolText));

  line("\n— get_debtors (top 3) —");
  line(textOf((await client.callTool({ name: "get_debtors", arguments: { limit: 3 } })) as ToolText));

  await client.close();
  line("\n✓ end-to-end ok");
}

main().catch((e) => {
  process.stderr.write(`MCP probe error: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
