/**
 * Пересобирает список инструментов в manifest.json из живого сервера — то, что видит
 * MCP-клиент (имя, заголовок, описание, inputSchema, outputSchema, аннотации), — и
 * синхронизирует версию с package.json. Реальная база не нужна: сервер поднимается с
 * фиктивным подключением в режиме только-чтения, tools/list к 1С не обращается.
 * Запуск: npm run manifest (после него — npm test: тест согласованности сверит тексты).
 */
import { readFileSync, writeFileSync } from "node:fs";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

async function main(): Promise<void> {
  const env: Record<string, string> = {};
  for (const [k, v] of Object.entries(process.env)) {
    if (v !== undefined && !k.startsWith("ODATA_")) env[k] = v;
  }
  Object.assign(env, {
    ODATA_BASE_URL: "https://example.com/odata/standard.odata/",
    ODATA_USERNAME: "user",
    ODATA_PASSWORD: "pass",
    READ_ONLY: "true",
  });

  const client = new Client({ name: "gen-manifest", version: "0.1.0" });
  await client.connect(new StdioClientTransport({ command: "node", args: ["dist/index.js"], env }));
  const { tools } = await client.listTools();
  await client.close();

  const pkg = JSON.parse(readFileSync("package.json", "utf8")) as { version: string };
  const manifest = JSON.parse(readFileSync("manifest.json", "utf8")) as Record<string, unknown>;
  manifest["version"] = pkg.version;
  manifest["tools"] = tools.map((t) => ({
    name: t.name,
    ...(t.title ? { title: t.title } : {}),
    description: t.description,
    inputSchema: t.inputSchema,
    ...(t.outputSchema ? { outputSchema: t.outputSchema } : {}),
    ...(t.annotations ? { annotations: t.annotations } : {}),
  }));
  writeFileSync("manifest.json", `${JSON.stringify(manifest, null, 2)}\n`);
  process.stdout.write(`manifest.json: ${tools.length} инструментов, версия ${pkg.version}\n`);
}

main().catch((e: unknown) => {
  process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
