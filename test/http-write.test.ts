import { mkdtempSync, mkdirSync, chmodSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, StreamableHTTPClientTransport } from "@modelcontextprotocol/client";
import { parseConfig, type HttpTransport } from "../src/config/env.js";
import { assertWriteJournalWritable, startHttpServer } from "../src/http.js";
import { fake1C } from "./helpers/fake-1c.js";

/**
 * Запись по HTTP (#67): сценарии, которые в сети случаются чаще, чем в stdio. Метрика: ни одного лишнего объекта в
 * 1С и ни одного «успеха» без перечитки — при рестарте между предпросмотром и подтверждением, параллельных
 * подтверждениях, обрыве клиента и лимите времени во время записи, устаревшем предпросмотре.
 */
const realFetch = globalThis.fetch;
const TOKEN = "w".repeat(32);
const http: HttpTransport = { host: "127.0.0.1", port: 0, token: TOKEN };
const base = { ODATA_BASE_URL: "http://1c.test/base/", ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };
const CREATE = "write.counterparty.create_counterparty";
const ALLOW = `${CREATE},write.operation.status,read.system.health_check`;

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const writeEnv = (journal: string, extra: Record<string, string> = {}) => ({
  READ_ONLY: "false",
  ODATA_WRITABLE: "true",
  MCP_TRANSPORT: "http",
  MCP_HTTP_TOKEN: TOKEN,
  ODATA_WRITE_JOURNAL_DIR: journal,
  TOOLS_ALLOW: ALLOW,
  ...extra,
});

type Result = { isError?: boolean; structuredContent: Record<string, unknown>; content: { text: string }[] };

async function serve(journal: string, extra: Record<string, string> = {}) {
  const f = fake1C(writeEnv(journal, extra));
  const server = await startHttpServer(f.ctx, http);
  const url = new URL(`http://127.0.0.1:${server.port}/mcp`);
  const client = new Client({ name: "bot", version: "1" });
  await client.connect(
    new StreamableHTTPClientTransport(url, {
      fetch: realFetch,
      requestInit: { headers: { Authorization: `Bearer ${TOKEN}` } },
    }),
  );
  const call = async (name: string, args: Record<string, unknown>) =>
    (await client.callTool({ name, arguments: args })) as Result;
  const stop = async () => {
    await client.close();
    await server.close();
  };
  return { f, url, call, stop };
}

const posts = (...fakes: Array<ReturnType<typeof fake1C>>) =>
  fakes.flatMap((f) => f.requests).filter((q) => q.method === "POST").length;

/** POST в 1С ждёт, пока тест его не отпустит; остальное — как обычно. */
function holdPosts() {
  const inner = globalThis.fetch;
  let release!: () => void;
  const gate = new Promise<void>((r) => (release = r));
  const stats = { started: 0 };
  vi.stubGlobal("fetch", async (url: string | URL, init?: RequestInit) => {
    if ((init?.method ?? "GET").toUpperCase() === "POST") {
      stats.started += 1;
      // Отмена по тайм-ауту 1С (ODATA_TIMEOUT_MS) прерывает ожидание, как у настоящего fetch.
      await new Promise<void>((resolve, reject) => {
        void gate.then(resolve);
        init?.signal?.addEventListener("abort", () => reject(new DOMException("aborted", "AbortError")), {
          once: true,
        });
      });
    }
    return inner(url, init);
  });
  return { stats, release };
}

/** Сырой вызов инструмента по HTTP — чтобы оборвать соединение посреди записи. */
function rawCall(url: URL, name: string, args: Record<string, unknown>, signal?: AbortSignal) {
  return realFetch(url, {
    method: "POST",
    ...(signal ? { signal } : {}),
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
    },
    body: JSON.stringify({ jsonrpc: "2.0", id: 9, method: "tools/call", params: { name, arguments: args } }),
  });
}

const journalEntries = (dir: string) =>
  readdirSync(dir, { recursive: true })
    .map(String)
    .filter((f) => f.endsWith(".json"))
    .map((f) => JSON.parse(readFileSync(join(dir, f), "utf8")) as Record<string, unknown>);

describe("запись по HTTP: условия старта", () => {
  const good = { ...base, ...writeEnv("/var/lib/1c-odata/journal") };

  it("все условия сразу — сервер настраивается на запись", () => {
    const c = parseConfig(good);
    expect(c.behavior.readOnly).toBe(false);
    expect(c.behavior.previewMaxAgeMs).toBe(15 * 60_000);
    expect(c.behavior.toolsAllow).toContain(CREATE);
  });

  it.each([
    ["журнал не задан", { ODATA_WRITE_JOURNAL_DIR: undefined }, /ODATA_WRITE_JOURNAL_DIR/],
    ["профиль all", { TOOLS_PROFILE: "all", TOOLS_ALLOW: undefined }, /TOOLS_PROFILE=all/],
    ["инструменты записи не перечислены", { TOOLS_ALLOW: "read.system.health_check" }, /поимённо/],
    ["нет write.operation.status", { TOOLS_ALLOW: CREATE }, /write\.operation\.status/],
    ["только write.operation.status", { TOOLS_ALLOW: "write.operation.status" }, /поимённо/],
    ["журнал из пробелов", { ODATA_WRITE_JOURNAL_DIR: "   " }, /ODATA_WRITE_JOURNAL_DIR/],
    ["база без WRITABLE", { ODATA_WRITABLE: "false" }, /ни одна база не открыта на запись/],
    ["запись в TOOLS_ALLOW при READ_ONLY=true", { READ_ONLY: "true" }, /только для чтения/],
  ])("%s — отказ при старте", (_n, extra, message) => {
    expect(() => parseConfig({ ...good, ...extra } as Record<string, string>)).toThrow(message);
  });

  it("срок предпросмотра: явный 0 — без срока; в stdio по умолчанию без срока", () => {
    expect(parseConfig({ ...good, ODATA_WRITE_PREVIEW_MAX_AGE: "0" }).behavior.previewMaxAgeMs).toBe(0);
    expect(parseConfig({ ...base, READ_ONLY: "false" }).behavior.previewMaxAgeMs).toBe(0);
  });

  it("каталог журнала без права записи — сервер не стартует и называет причину", async () => {
    const dir = mkdtempSync(join(tmpdir(), "1c-odata-ro-"));
    const locked = join(dir, "locked");
    mkdirSync(locked);
    chmodSync(locked, 0o500);
    try {
      await expect(assertWriteJournalWritable(join(locked, "journal"))).rejects.toThrow(
        /ODATA_WRITE_JOURNAL_DIR недоступен на запись/,
      );
      const f = fake1C(writeEnv(join(locked, "journal")));
      await expect(startHttpServer(f.ctx, http)).rejects.toThrow(/недоступен на запись/);
    } finally {
      chmodSync(locked, 0o700);
    }
  });

  it("инструмент записи вне TOOLS_ALLOW вызвать нельзя", async () => {
    const s = await serve(mkdtempSync(join(tmpdir(), "1c-odata-hw-")));
    await expect(s.call("write.sales.create_invoice", {})).rejects.toThrow(/not found/);
    await s.stop();
  });

  it("health_check: запись включена, срок предпросмотра, путь журнала по сети не показывается", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal);
    const r = await s.call("read.system.health_check", {});
    expect(r.structuredContent).toMatchObject({ writeEnabled: true, previewMaxAgeMinutes: 15 });
    expect(JSON.stringify(r.structuredContent)).not.toContain(journal);
    await s.stop();
  });
});

describe("запись по HTTP: гарантии журнала", () => {
  it("рестарт процесса между предпросмотром и подтверждением: подтверждение находит операцию, один POST", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const args = { name: "ZZ-TEST HTTP рестарт" };
    const a = await serve(journal);
    const p = await a.call(CREATE, args);
    expect(p.isError, p.content[0]?.text).toBeFalsy();
    await a.stop();
    // Новый процесс: новый контекст, тот же том журнала.
    const b = await serve(journal);
    const c = await b.call(CREATE, {
      ...args,
      confirm: true,
      operationId: p.structuredContent["operationId"],
    });
    expect(c.isError, c.content[0]?.text).toBeFalsy();
    expect(c.structuredContent).toMatchObject({ created: true });
    const again = await b.call(CREATE, {
      ...args,
      confirm: true,
      operationId: p.structuredContent["operationId"],
    });
    expect(again.structuredContent).toMatchObject({ replayed: true, ref: c.structuredContent["ref"] });
    expect(posts(a.f, b.f)).toBe(1);
    await b.stop();
  });

  it("два одновременных подтверждения одного operationId: в 1С ровно один POST, второе — из журнала или «неизвестен»", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal);
    const args = { name: "ZZ-TEST HTTP параллельно" };
    const p = await s.call(CREATE, args);
    const confirm = { ...args, confirm: true, operationId: p.structuredContent["operationId"] };
    const hold = holdPosts();
    const first = s.call(CREATE, confirm);
    await vi.waitFor(() => expect(hold.stats.started).toBe(1));
    const second = s.call(CREATE, confirm);
    const r2 = await second;
    hold.release();
    const r1 = await first;
    expect(r1.isError, r1.content[0]?.text).toBeFalsy();
    // Второе пришло, пока первое писало: дубля нет, исход второго — «неизвестен», сверка — по статусу.
    expect(r2.isError).toBe(true);
    expect(r2.content[0]!.text).toMatch(/неизвестен|write\.operation\.status/);
    expect(hold.stats.started).toBe(1);
    const st = await s.call("write.operation.status", { operationId: p.structuredContent["operationId"] });
    expect(st.structuredContent["status"]).toBe("succeeded");
    await s.stop();
  });

  it("клиент разорвал соединение во время записи: POST дописывается, журнал «успешно», повтор — из журнала", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal);
    const args = { name: "ZZ-TEST HTTP обрыв" };
    const p = await s.call(CREATE, args);
    const operationId = p.structuredContent["operationId"];
    const hold = holdPosts();
    const controller = new AbortController();
    const call = rawCall(s.url, CREATE, { ...args, confirm: true, operationId }, controller.signal);
    await vi.waitFor(() => expect(hold.stats.started).toBe(1));
    controller.abort();
    await expect(call).rejects.toThrow();
    hold.release();
    // Отправленная запись не превращается в «не записано»: журнал доводит её до перечитки.
    await vi.waitFor(() =>
      expect(journalEntries(journal).find((e) => e["operationId"] === operationId)?.["state"]).toBe(
        "succeeded",
      ),
    );
    const again = await s.call(CREATE, { ...args, confirm: true, operationId });
    expect(again.structuredContent).toMatchObject({ replayed: true });
    expect(posts(s.f)).toBe(1);
    await s.stop();
  });

  it("ODATA_TOOL_DEADLINE_MS во время POST: лимит времени чтения запись не обрывает — один POST, успех", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal, { ODATA_TOOL_DEADLINE_MS: "50" });
    const args = { name: "ZZ-TEST HTTP лимит" };
    const p = await s.call(CREATE, args);
    const hold = holdPosts();
    const confirm = s.call(CREATE, {
      ...args,
      confirm: true,
      operationId: p.structuredContent["operationId"],
    });
    await vi.waitFor(() => expect(hold.stats.started).toBe(1));
    await new Promise((r) => setTimeout(r, 200));
    hold.release();
    const r = await confirm;
    expect(r.isError, r.content[0]?.text).toBeFalsy();
    expect(r.structuredContent).toMatchObject({ created: true });
    expect(posts(s.f)).toBe(1);
    await s.stop();
  });

  it("1С не ответила за ODATA_TIMEOUT_MS: исход неизвестен, повтор заблокирован до write.operation.status", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal, { ODATA_TIMEOUT_MS: "100" });
    const args = { name: "ZZ-TEST HTTP тайм-аут" };
    const p = await s.call(CREATE, args);
    const operationId = p.structuredContent["operationId"];
    const hold = holdPosts();
    const r = await s.call(CREATE, { ...args, confirm: true, operationId });
    expect(r.isError).toBe(true);
    expect(r.content[0]!.text).toContain("неизвестен");
    hold.release();
    const blocked = await s.call(CREATE, { ...args, confirm: true, operationId });
    expect(blocked.isError).toBe(true);
    expect(blocked.content[0]!.text).toMatch(/неизвестен|write\.operation\.status/);
    expect(hold.stats.started).toBe(1);
    await s.stop();
  });

  it("устаревший предпросмотр — отказ «сделайте новый», без POST; выполненная операция — по-прежнему из журнала", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal, { ODATA_WRITE_PREVIEW_MAX_AGE: "15" });
    const args = { name: "ZZ-TEST HTTP старый предпросмотр" };
    const p = await s.call(CREATE, args);
    const operationId = String(p.structuredContent["operationId"]);
    // Предпросмотр «сделан» 20 минут назад: сдвигаем дату в журнале, а не часы процесса.
    const file = readdirSync(journal, { recursive: true })
      .map(String)
      .find((f) => f.endsWith(`${operationId}.json`))!;
    const entry = JSON.parse(readFileSync(join(journal, file), "utf8")) as Record<string, unknown>;
    writeFileSync(
      join(journal, file),
      JSON.stringify({ ...entry, createdAt: new Date(Date.now() - 20 * 60_000).toISOString() }),
    );
    const r = await s.call(CREATE, { ...args, confirm: true, operationId });
    expect(r.isError).toBe(true);
    expect(r.content[0]!.text).toContain("Сделайте новый предпросмотр");
    expect(posts(s.f)).toBe(0);
    // Свежий предпросмотр подтверждается; его повтор через 20 минут — всё равно из журнала, без отказа.
    const fresh = await s.call(CREATE, args);
    const id2 = String(fresh.structuredContent["operationId"]);
    const done = await s.call(CREATE, { ...args, confirm: true, operationId: id2 });
    expect(done.isError, done.content[0]?.text).toBeFalsy();
    const file2 = readdirSync(journal, { recursive: true })
      .map(String)
      .find((f) => f.endsWith(`${id2}.json`))!;
    const e2 = JSON.parse(readFileSync(join(journal, file2), "utf8")) as Record<string, unknown>;
    writeFileSync(
      join(journal, file2),
      JSON.stringify({ ...e2, createdAt: new Date(Date.now() - 20 * 60_000).toISOString() }),
    );
    const replay = await s.call(CREATE, { ...args, confirm: true, operationId: id2 });
    expect(replay.structuredContent).toMatchObject({ replayed: true });
    expect(posts(s.f)).toBe(1);
    await s.stop();
  });

  it("ревью #67: журнал перестал быть доступен на запись — путь к нему не уходит клиенту ни в health_check, ни в ошибке", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal);
    const args = { name: "ZZ-TEST HTTP журнал сломался" };
    const p = await s.call(CREATE, args);
    chmodSync(journal, 0o500);
    const inner = readdirSync(journal).map((d) => join(journal, d));
    for (const d of inner) chmodSync(d, 0o500);
    try {
      const h = await s.call("read.system.health_check", {});
      expect(JSON.stringify(h)).not.toContain(journal);
      const c = await s.call(CREATE, {
        ...args,
        confirm: true,
        operationId: p.structuredContent["operationId"],
      });
      expect(c.isError).toBe(true);
      expect(JSON.stringify(c)).not.toContain(journal);
      // Ни пути родителя, ни текста ошибки файловой системы (ревью 1.0: замена подстроки пропускала родителя).
      expect(JSON.stringify([h, c])).not.toMatch(/EACCES|EPERM|\/tmp|\/var\/folders|permission denied/i);
      expect(posts(s.f)).toBe(0);
    } finally {
      for (const d of inner) chmodSync(d, 0o700);
      chmodSync(journal, 0o700);
    }
    await s.stop();
  });

  it("ревью #67: нечитаемая или будущая дата предпросмотра — как устаревшая, без POST", async () => {
    const journal = mkdtempSync(join(tmpdir(), "1c-odata-hw-"));
    const s = await serve(journal);
    for (const createdAt of ["not-a-date", new Date(Date.now() + 86_400_000).toISOString()]) {
      const args = { name: `ZZ-TEST HTTP дата ${createdAt.slice(0, 4)}` };
      const p = await s.call(CREATE, args);
      const id = String(p.structuredContent["operationId"]);
      const file = readdirSync(journal, { recursive: true })
        .map(String)
        .find((f) => f.endsWith(`${id}.json`))!;
      const entry = JSON.parse(readFileSync(join(journal, file), "utf8")) as Record<string, unknown>;
      writeFileSync(join(journal, file), JSON.stringify({ ...entry, createdAt }));
      const r = await s.call(CREATE, { ...args, confirm: true, operationId: id });
      expect(r.isError, createdAt).toBe(true);
      expect(r.content[0]!.text).toContain("Сделайте новый предпросмотр");
    }
    expect(posts(s.f)).toBe(0);
    await s.stop();
  });
});
