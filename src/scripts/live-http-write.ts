/**
 * Живой прогон записи по HTTP (#67) против НАСТОЯЩЕЙ, лучше тестовой базы 1С: экземпляр коннектора в режиме
 * MCP_TRANSPORT=http с записью только установки цен, сценарий интеграции-бота:
 *   предпросмотр → рестарт процесса → подтверждение с обрывом соединения клиентом → повтор подтверждения →
 *   write.operation.status → ровно один непроведённый документ в 1С.
 * 1.1, безопасный повтор создания — через локального посредника перед 1С, который на один POST документа цен
 * «теряет» ответ (запрос доходит до 1С, соединение рвётся) или «роняет» запрос (до 1С не доходит): исход неизвестен →
 * рестарт → повтор доводит до одного документа со ссылкой предпросмотра; два одновременных подтверждения — один.
 * Пометку удаления созданного документа делает человек (или test:live): этому экземпляру она намеренно не дана.
 *
 *   npm run test:live:http -- --db <база> --item <Ref_Key позиции> --yes [--price-kind <Ref_Key>] [--price 123.45]
 *
 * Вид цен по умолчанию — из последней проведённой установки цен базы. Журнал — во временном каталоге (том).
 * Код возврата 1 — если хоть одна проверка не прошла.
 */
import { spawn, type ChildProcess } from "node:child_process";
import { createServer as createHttpServer, type IncomingMessage } from "node:http";
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

  // Посредник перед 1С: всё пересылает как есть, а один POST создания документа цен — по режиму.
  const tag = Object.keys(process.env)
    .map((k) => /^ODATA_DB_(.+)_BASE_URL$/.exec(k)?.[1])
    .find((t) => t?.toLowerCase() === db);
  const baseVar = tag ? `ODATA_DB_${tag}_BASE_URL` : "ODATA_BASE_URL";
  const upstream = new URL(String(process.env[baseVar]));
  let mode: "lost" | "drop" | undefined;
  const body = (req: IncomingMessage) =>
    new Promise<Buffer>((resolve) => {
      const parts: Buffer[] = [];
      req.on("data", (d: Buffer) => parts.push(d));
      req.on("end", () => resolve(Buffer.concat(parts)));
    });
  const proxy = createHttpServer((req, res) => {
    void (async () => {
      const data = await body(req);
      const path = decodeURIComponent(req.url ?? "");
      const target = req.method === "POST" && /\/Document_УстановкаЦенНоменклатуры\?/.test(path);
      const m = target ? mode : undefined;
      if (m) mode = undefined;
      if (m === "drop") return req.socket.destroy();
      const headers = Object.fromEntries(
        Object.entries(req.headers).filter(([k]) => !["host", "connection", "content-length"].includes(k)),
      ) as Record<string, string>;
      const r = await fetch(new URL(req.url ?? "/", upstream.origin), {
        method: req.method ?? "GET",
        headers,
        ...(data.length ? { body: new Uint8Array(data) } : {}),
      });
      const out = Buffer.from(await r.arrayBuffer());
      if (m === "lost") return req.socket.destroy(); // 1С записала, ответ потерян
      res.writeHead(r.status, { "content-type": r.headers.get("content-type") ?? "application/json" });
      res.end(out);
    })().catch(() => req.socket.destroy());
  });
  await new Promise<void>((r) => proxy.listen(0, "127.0.0.1", () => r()));
  const proxyPort = (proxy.address() as { port: number }).port;
  const proxied = `http://127.0.0.1:${proxyPort}${upstream.pathname}${upstream.search}`;

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
    [baseVar]: proxied,
  };
  let proc: ChildProcess | undefined;
  const start = async (extra: Record<string, string> = {}) => {
    proc = spawn(process.execPath, [entry], {
      env: { ...env, ...extra },
      stdio: ["ignore", "ignore", "inherit"],
    });
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
  const comment = `ZZ-TEST live http ${new Date().toISOString().slice(0, 19)} ${randomBytes(3).toString("hex")}`;
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
  const docsOf = async (tag: string) =>
    (
      await conn.client.getCollection(
        `Document_УстановкаЦенНоменклатуры?$format=json&$filter=Date ge datetime'${day}T00:00:00'&$select=Ref_Key,Posted,Комментарий`,
      )
    ).value.filter((d) => String(d["Комментарий"] ?? "").startsWith(tag));
  const docs = await docsOf(comment);
  check(
    "в 1С ровно один непроведённый документ",
    docs.length === 1 && docs[0]!["Posted"] === false && String(docs[0]!["Ref_Key"]) === String(ref),
    `документов: ${docs.length}`,
  );
  const created = typeof ref === "string" ? [ref] : [];

  // 1.1: неизвестный исход (ответ потерян / запрос не дошёл) → рестарт → повтор подтверждения — один документ.
  for (const fail of ["lost", "drop"] as const) {
    const tag = `${comment} ${fail}`;
    const a = { ...args, comment: tag };
    await start();
    c = await connect();
    const pv = await call(c, "write.pricing.create_price_setting", a);
    const op = pv.structuredContent?.["operationId"];
    const assigned = pv.structuredContent?.["ref"];
    check(`${fail}: предпросмотр отдаёт ref`, typeof assigned === "string");
    mode = fail;
    const first = await call(c, "write.pricing.create_price_setting", {
      ...a,
      confirm: true,
      operationId: op,
    });
    check(
      `${fail}: первая попытка — исход неизвестен`,
      first.isError === true && /неизвестен/.test(text(first)),
    );
    await c.close();
    await stop();
    await start();
    c = await connect();
    // Статус до повтора — только у «не дошло» (показать retrySafe): при «потерян» он сам сверил бы операцию, а
    // проверяем именно повтор подтверждения без ручной сверки.
    let stInfo = "не спрашивали";
    if (fail === "drop") {
      const st = await call(c, "write.operation.status", { operationId: op });
      stInfo = `${String(st.structuredContent?.["status"])}${st.structuredContent?.["retrySafe"] ? ", retrySafe" : ""}`;
    }
    const retry = await call(c, "write.pricing.create_price_setting", {
      ...a,
      confirm: true,
      operationId: op,
    });
    const how = retry.structuredContent?.["reconciled"]
      ? "найден по ссылке"
      : retry.structuredContent?.["retried"]
        ? "POST повторён"
        : "";
    const found = await docsOf(tag);
    check(
      `${fail}: рестарт и повтор — один документ со ссылкой предпросмотра`,
      !retry.isError &&
        how === (fail === "lost" ? "найден по ссылке" : "POST повторён") &&
        found.length === 1 &&
        String(found[0]!["Ref_Key"]) === String(assigned),
      `status до повтора: ${stInfo}; повтор: ${retry.isError ? text(retry).slice(0, 120) : how}; документов: ${found.length}`,
    );
    const again = await call(c, "write.pricing.create_price_setting", {
      ...a,
      confirm: true,
      operationId: op,
    });
    check(
      `${fail}: ещё один повтор — из журнала, документ тот же`,
      again.structuredContent?.["replayed"] === true && (await docsOf(tag)).length === 1,
    );
    if (typeof assigned === "string") created.push(assigned);
    await c.close();
    await stop();
  }

  // 1.1: два одновременных подтверждения одного operationId — один документ.
  {
    const tag = `${comment} par`;
    const a = { ...args, comment: tag };
    await start();
    c = await connect();
    const pv = await call(c, "write.pricing.create_price_setting", a);
    const op = pv.structuredContent?.["operationId"];
    const c2 = await connect();
    const [r1, r2] = await Promise.all([
      call(c, "write.pricing.create_price_setting", { ...a, confirm: true, operationId: op }),
      call(c2, "write.pricing.create_price_setting", { ...a, confirm: true, operationId: op }),
    ]);
    const kinds = [r1, r2].map((r) => (r.isError ? text(r).slice(0, 13) : "ok")).join(", ");
    const found = await docsOf(tag);
    check(
      "два одновременных подтверждения — один документ",
      [r1, r2].some((r) => !r.isError) && found.length === 1,
      `ответы: ${kinds}; документов: ${found.length}`,
    );
    for (const d of found) created.push(String(d["Ref_Key"]));
    await c.close();
    await c2.close();
    await stop();
  }

  proxy.close();
  line();
  for (const [name, ok, detail] of rows) line(`${ok ? "✓" : "✗"} ${name.padEnd(52)} ${detail}`);
  if (created.length)
    line(
      `\nСозданные документы (пометьте на удаление): Document_УстановкаЦенНоменклатуры ${created.join(", ")}`,
    );
  return rows.every(([, ok]) => ok) ? 0 : 1;
}

main()
  .then((code) => process.exit(code))
  .catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? (e.stack ?? e.message) : String(e)}\n`);
    process.exit(1);
  });
