/**
 * Живой прогон записи по HTTP (#67) против НАСТОЯЩЕЙ, лучше тестовой базы 1С: экземпляр коннектора в режиме
 * MCP_TRANSPORT=http с записью только установки цен, сценарий интеграции-бота:
 *   предпросмотр → рестарт процесса → подтверждение с обрывом соединения клиентом → повтор подтверждения →
 *   write.operation.status → ровно один непроведённый документ в 1С.
 * Пометку удаления созданного документа делает человек (или test:live): этому экземпляру она намеренно не дана.
 *
 *   npm run test:live:http -- --db <база> --item <Ref_Key позиции> --yes [--price-kind <Ref_Key>] [--price 123.45]
 *
 * Вид цен по умолчанию — из последней проведённой установки цен базы. Журнал — во временном каталоге (том).
 * Код возврата 1 — если хоть одна проверка не прошла.
 */
import { spawn, type ChildProcess } from "node:child_process";
import { randomBytes } from "node:crypto";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

type Json = Record<string, unknown>;
type CallResult = { isError?: boolean; structuredContent?: Json; content?: Array<{ text?: string }> };

const arg = (name: string): string | undefined => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const line = (s = "") => process.stdout.write(`${s}\n`);
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function main(): Promise<number> {
  const db = arg("db");
  const item = arg("item");
  if (!db || !item || !process.argv.includes("--yes")) {
    line("Пишет в базу 1С по HTTP. Запуск: npm run test:live:http -- --db <база> --item <Ref_Key> --yes");
    return 2;
  }
  const { Client, StreamableHTTPClientTransport } = await import("@modelcontextprotocol/client");
  const { loadConfig } = await import("../config/env.js");
  const { ServerContext } = await import("../context.js");
  const conn = new ServerContext(loadConfig()).db(db);
  let priceKind = arg("price-kind");
  if (!priceKind) {
    // Вид цен: в УТ — в таблице «ВидыЦен», в Бухгалтерии — в шапке; поле выбирается по схеме базы.
    const set = "Document_УстановкаЦенНоменклатуры";
    const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((f) => f.name) ?? []);
    const field = props.has("ВидыЦен") ? "ВидыЦен" : "ТипЦен_Key";
    const last = (
      await conn.client.getCollection(
        `${set}?$format=json&$top=1&$orderby=Date desc&$filter=Posted eq true&$select=${field}`,
      )
    ).value[0] as Json | undefined;
    priceKind = String(
      (last?.["ВидыЦен"] as Json[] | undefined)?.[0]?.["ВидЦены_Key"] ?? last?.["ТипЦен_Key"] ?? "",
    );
  }
  if (!priceKind) {
    line("Вид цен не найден: укажите --price-kind.");
    return 2;
  }

  const token = randomBytes(24).toString("hex");
  const port = 18000 + Math.floor(Math.random() * 1000);
  const journal = mkdtempSync(join(tmpdir(), "1c-odata-live-http-"));
  const entry = fileURLToPath(new URL("../index.js", import.meta.url));
  const env = {
    ...process.env,
    MCP_TRANSPORT: "http",
    MCP_HTTP_HOST: "127.0.0.1",
    MCP_HTTP_PORT: String(port),
    MCP_HTTP_TOKEN: token,
    READ_ONLY: "false",
    TOOLS_PROFILE: "",
    TOOLS_ALLOW: "write.pricing.create_price_setting,write.operation.status,read.system.health_check",
    ODATA_WRITE_JOURNAL_DIR: journal,
  };
  let proc: ChildProcess | undefined;
  const start = async () => {
    proc = spawn(process.execPath, [entry], { env, stdio: ["ignore", "ignore", "inherit"] });
    for (let i = 0; i < 100; i++) {
      try {
        if ((await fetch(`http://127.0.0.1:${port}/healthz`)).ok) return;
      } catch {
        // ещё не слушает
      }
      await sleep(200);
    }
    throw new Error("экземпляр не поднялся за 20 с");
  };
  const stop = async () => {
    const p = proc;
    if (!p || p.exitCode !== null) return;
    p.kill("SIGTERM");
    await new Promise((r) => p.once("exit", r));
  };
  process.on("exit", () => proc?.kill("SIGTERM"));
  const url = new URL(`http://127.0.0.1:${port}/mcp`);
  const connect = async () => {
    const c = new Client({ name: "live-http-write", version: "1" });
    await c.connect(
      new StreamableHTTPClientTransport(url, {
        requestInit: { headers: { Authorization: `Bearer ${token}` } },
      }),
    );
    return c;
  };
  const call = async (c: InstanceType<typeof Client>, name: string, a: Json) =>
    (await c.callTool({ name, arguments: { database: db, ...a } })) as CallResult;
  const text = (r: CallResult) => (r.content ?? []).map((x) => x.text ?? "").join("");

  const rows: Array<[string, boolean, string]> = [];
  const check = (name: string, ok: boolean, detail = "") => rows.push([name, ok, detail]);
  const comment = `ZZ-TEST live http ${new Date().toISOString().slice(0, 16)}`;
  const args = { priceKind, lines: [{ itemRef: item, price: Number(arg("price") ?? 123.45) }], comment };

  await start();
  let c = await connect();
  const h = await call(c, "read.system.health_check", {});
  check("health_check: запись включена", h.structuredContent?.["writeEnabled"] === true);
  const p = await call(c, "write.pricing.create_price_setting", args);
  const operationId = p.structuredContent?.["operationId"];
  check(
    "предпросмотр",
    !p.isError && typeof operationId === "string",
    p.isError ? text(p).slice(0, 160) : "",
  );
  await c.close();
  await stop();
  await start();
  c = await connect();
  // Подтверждение; соединение клиент рвёт вскоре после отправки.
  const abort = new AbortController();
  const raw = fetch(url, {
    method: "POST",
    signal: abort.signal,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "tools/call",
      params: {
        name: "write.pricing.create_price_setting",
        arguments: { database: db, ...args, confirm: true, operationId },
      },
    }),
  });
  setTimeout(() => abort.abort(), 150);
  await raw.catch(() => undefined);
  await sleep(8_000);
  const again = await call(c, "write.pricing.create_price_setting", { ...args, confirm: true, operationId });
  check(
    "рестарт + обрыв: повтор подтверждения — из журнала",
    again.structuredContent?.["replayed"] === true,
    again.isError ? text(again).slice(0, 200) : "",
  );
  const st = await call(c, "write.operation.status", { operationId });
  const ref = st.structuredContent?.["ref"];
  check("write.operation.status — succeeded", st.structuredContent?.["status"] === "succeeded");
  await c.close();
  await stop();

  const day = new Date().toISOString().slice(0, 10);
  const docs = (
    await conn.client.getCollection(
      `Document_УстановкаЦенНоменклатуры?$format=json&$filter=Date ge datetime'${day}T00:00:00'&$select=Ref_Key,Posted,Комментарий`,
    )
  ).value.filter((d) => String(d["Комментарий"] ?? "").startsWith(comment));
  check(
    "в 1С ровно один непроведённый документ",
    docs.length === 1 && docs[0]!["Posted"] === false && String(docs[0]!["Ref_Key"]) === String(ref),
    `документов: ${docs.length}`,
  );

  line();
  for (const [name, ok, detail] of rows) line(`${ok ? "✓" : "✗"} ${name.padEnd(52)} ${detail}`);
  if (typeof ref === "string")
    line(`\nСозданный документ (пометьте на удаление): Document_УстановкаЦенНоменклатуры ${ref}`);
  return rows.every(([, ok]) => ok) ? 0 : 1;
}

main()
  .then((code) => process.exit(code))
  .catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? (e.stack ?? e.message) : String(e)}\n`);
    process.exit(1);
  });
