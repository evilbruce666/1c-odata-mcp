/**
 * Живая проверка на ЧТЕНИЕ против НАСТОЯЩЕЙ базы 1С — перед релизом, когда поддельной 1С мало.
 * Ничего не пишет: глобальный fetch пропускает только GET/HEAD, любая попытка записи — провал
 * прогона. Журнал операций — во временном каталоге.
 *
 *   npm run test:live:read -- [--db <имя базы>] [--http]   (без --db — все настроенные базы)
 *
 * С --http те же проверки идут настоящим клиентом через HTTP-транспорт на случайном порту локальной машины;
 * outputSchema тогда проверяет сам клиент.
 *
 * Проверяет: health_check; все инструменты чтения без обязательных аргументов; проводки последних
 * проведённых документов — что субконто действительно приходят (dimensionsStatus), а не только что
 * ответ есть. Именно такой прогон нашёл, что на БП 3.0 записи регистра субконто не публикуют (#49).
 * Код возврата 1 — если хоть одна проверка не прошла.
 */
import { randomBytes } from "node:crypto";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

process.env["ODATA_WRITE_JOURNAL_DIR"] = mkdtempSync(join(tmpdir(), "1c-odata-live-read-"));

type Json = Record<string, unknown>;
type Tool = {
  inputSchema?: { safeParse: (v: unknown) => { success: boolean; data?: Json; error?: unknown } };
  outputSchema?: { safeParse: (v: unknown) => { success: boolean; error?: { issues?: unknown } } };
  handler: (
    args: Json,
    extra: Json,
  ) => Promise<{ isError?: boolean; structuredContent?: Json; content?: Array<{ text?: string }> }>;
};

/** Инструменты, которым без аргументов положено ответить отказом проверки ввода. */
const NEEDS_ARGS = new Set(["read.analytics.get_deal_history"]);
/** Документы, проводки которых смотрим: банк и реализация есть почти в любой базе БП. */
const POSTING_DOCS = [
  "Document_СписаниеСРасчетногоСчета",
  "Document_ПоступлениеНаРасчетныйСчет",
  "Document_РеализацияТоваровУслуг",
];

const arg = (name: string) => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const line = (s = "") => process.stdout.write(`${s}\n`);

const OVER_HTTP = process.argv.includes("--http");
if (OVER_HTTP) {
  Object.assign(process.env, {
    MCP_TRANSPORT: "http",
    MCP_HTTP_PORT: "0",
    MCP_HTTP_HOST: "127.0.0.1",
    MCP_HTTP_TOKEN: randomBytes(32).toString("hex"),
    READ_ONLY: "true",
  });
}

async function main(): Promise<number> {
  const blockedWrites: string[] = [];
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    const method = (init?.method ?? "GET").toUpperCase();
    if (method !== "GET" && method !== "HEAD") {
      blockedWrites.push(method);
      throw new Error(`live-read: ${method} заблокирован — прогон только на чтение`);
    }
    return realFetch(input, init);
  };

  const { loadConfig } = await import("../config/env.js");
  const { ServerContext } = await import("../context.js");
  const { createServer } = await import("../mcp/server.js");
  const cfg = loadConfig();
  const ctx = new ServerContext(cfg);
  const tools = (createServer(ctx) as unknown as { _registeredTools: Record<string, Tool> })._registeredTools;
  const dbs = arg("db") ? [arg("db")!] : cfg.connections.map((c) => c.name);

  const rows: Array<[string, boolean, string]> = [];
  const check = (name: string, ok: boolean, detail = "") => rows.push([name, ok, detail]);
  const http = OVER_HTTP ? await connectOverHttp(ctx, cfg.http!, realFetch) : undefined;
  const run = async (name: string, a: Json) => {
    if (http) {
      try {
        const r = (await http.client.callTool({ name, arguments: a })) as {
          isError?: boolean;
          structuredContent?: Json;
          content?: Array<{ text?: string }>;
          _meta?: Record<string, unknown>;
        };
        return {
          err: !!r.isError,
          sc: r.structuredContent ?? {},
          text: r.content?.[0]?.text ?? "",
          kind: String(r._meta?.["odata-1c/errorKind"] ?? ""),
        };
      } catch (e) {
        return { err: true, sc: {} as Json, text: e instanceof Error ? e.message : String(e), kind: "" };
      }
    }
    const tool = tools[name]!;
    const parsed = tool.inputSchema ? tool.inputSchema.safeParse(a) : { success: true, data: a };
    if (!parsed.success)
      return { err: true, sc: {} as Json, text: "входные данные не прошли схему", kind: "" };
    const r = await tool.handler(parsed.data!, {});
    // Настоящий MCP-клиент отвергает ответ, не прошедший outputSchema; прямой вызов обработчика — нет.
    if (!r.isError && tool.outputSchema) {
      const out = tool.outputSchema.safeParse(r.structuredContent);
      if (!out.success)
        return {
          err: true,
          sc: {},
          text: `ответ не проходит outputSchema: ${JSON.stringify(out.error?.issues).slice(0, 300)}`,
          kind: "",
        };
    }
    const meta = (r as { _meta?: Record<string, unknown> })._meta;
    return {
      err: !!r.isError,
      sc: r.structuredContent ?? {},
      text: r.content?.[0]?.text ?? "",
      kind: String(meta?.["odata-1c/errorKind"] ?? ""),
    };
  };
  /**
   * Отказ «нет в этой конфигурации» (регистр склада в Бухгалтерии, Хозрасчётный в УТ) и «нужны аргументы» — не
   * провал прогона: инструмент ответил как задумано. Печатается пропуском, чтобы его было видно.
   */
  const skipped = (r: { err: boolean; kind: string }) =>
    r.err && (r.kind === "not_published" || r.kind === "input");
  const noArgReads = Object.entries(tools)
    .filter(([name, t]) => {
      if (!name.startsWith("read.")) return false;
      return t.inputSchema ? t.inputSchema.safeParse({}).success : true;
    })
    .map(([name]) => name);

  for (const db of dbs) {
    const h = await run("read.system.health_check", { database: db });
    check(
      `${db}: health_check`,
      !h.err,
      h.err ? h.text.slice(0, 120) : `объектов ${String(h.sc["entityCount"])}`,
    );

    for (const name of noArgReads) {
      const r = await run(name, { database: db });
      if (!NEEDS_ARGS.has(name) && skipped(r)) {
        line(
          `  ${db}: ${name} — ${r.kind === "input" ? "нужны аргументы" : "нет в этой конфигурации"}, пропуск`,
        );
        continue;
      }
      const ok = NEEDS_ARGS.has(name) ? r.err : !r.err;
      if (!ok) check(`${db}: ${name}`, false, r.text.slice(0, 120));
    }
    check(`${db}: инструменты чтения без аргументов (${noArgReads.length})`, true);

    for (const entitySet of POSTING_DOCS) {
      const s = await run("read.document.search_documents", {
        database: db,
        entitySet,
        postedOnly: true,
        limit: 1,
      });
      const doc = (s.sc["rows"] as Json[] | undefined)?.[0];
      if (s.err || !doc) {
        line(`  ${db}: ${entitySet} — нет проведённых документов, пропуск`);
        continue;
      }
      const p = await run("read.accounting.get_document_postings", {
        database: db,
        documentEntity: entitySet,
        documentRef: doc["ref"],
      });
      if (skipped(p)) {
        line(`  ${db}: проводки ${entitySet} — нет в этой конфигурации, пропуск`);
        continue;
      }
      const statuses = new Map<string, number>();
      for (const posting of (p.sc["postings"] as Json[] | undefined) ?? [])
        for (const side of [posting["debit"], posting["credit"]] as Array<Json | null>)
          if (side)
            statuses.set(
              String(side["dimensionsStatus"]),
              (statuses.get(String(side["dimensionsStatus"])) ?? 0) + 1,
            );
      const unavailable = (statuses.get("not_in_schema") ?? 0) + (statuses.get("not_returned") ?? 0);
      check(
        `${db}: субконто проводок ${entitySet.replace("Document_", "")}`,
        !p.err && unavailable === 0,
        p.err ? p.text.slice(0, 120) : [...statuses].map(([k, v]) => `${k}=${v}`).join(", "),
      );
    }
  }
  check("ни одной попытки записи", blockedWrites.length === 0, blockedWrites.join(", "));
  if (http) {
    check("по HTTP видно только чтение", http.writeTools === 0, `инструментов записи: ${http.writeTools}`);
    await http.close();
  }

  for (const [name, ok, detail] of rows) line(`${ok ? "✓" : "✗"} ${name}${detail ? ` — ${detail}` : ""}`);
  const failed = rows.filter(([, ok]) => !ok).length;
  line(failed ? `\nНе прошло: ${failed}` : "\nВсё прошло.");
  return failed ? 1 : 0;
}

/** Настоящий клиент через HTTP-транспорт: то, что увидит интеграция на сервере. */
async function connectOverHttp(
  ctx: import("../context.js").ServerContext,
  cfg: import("../config/env.js").HttpTransport,
  fetchImpl: typeof fetch,
) {
  const { startHttpServer } = await import("../http.js");
  const { Client, StreamableHTTPClientTransport } = await import("@modelcontextprotocol/client");
  const server = await startHttpServer(ctx, cfg);
  const client = new Client({ name: "live-read", version: "1" });
  await client.connect(
    new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${server.port}/mcp`), {
      fetch: fetchImpl,
      requestInit: { headers: { Authorization: `Bearer ${cfg.token}` } },
    }),
  );
  const { tools } = await client.listTools();
  return {
    client,
    writeTools: tools.filter((t) => !t.name.startsWith("read.")).length,
    close: async () => {
      await client.close();
      await server.close();
    },
  };
}

main().then(
  (code) => process.exit(code),
  (e: unknown) => {
    line(`Ошибка прогона: ${e instanceof Error ? e.message : String(e)}`);
    process.exit(1);
  },
);
