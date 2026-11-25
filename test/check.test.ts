import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { runCheck, TRADE_CASH, TRADE_REQUIRED } from "../src/cli/check.js";
import { BP30_METADATA_XML, UT11_METADATA_XML } from "./helpers/metadata-snapshot.js";

// npx -y 1c-odata-mcp --check: диагностика подключения из терминала до клиента (аудит 25.11).
const PASSWORD = "pass-secret-987";
const ENV = {
  ODATA_BASE_URL: "https://host.invalid/db/odata/standard.odata/",
  ODATA_USERNAME: "user",
  ODATA_PASSWORD: PASSWORD,
};
const calls: Array<{ method: string; url: string }> = [];
function stub(body: string, status = 200, type = "application/xml") {
  calls.length = 0;
  vi.stubGlobal("fetch", async (url: string, init?: { method?: string }) => {
    calls.push({ method: init?.method ?? "GET", url: String(url) });
    return new Response(body, { status, headers: { "Content-Type": type } });
  });
}
async function run(args: string[], env: Record<string, string> = {}) {
  const out: string[] = [];
  const code = await runCheck(args, env, (l) => out.push(l));
  return { code, text: out.join("\n") };
}
afterEach(() => vi.unstubAllGlobals());

describe("--check по переменным окружения", () => {
  it("Бухгалтерия: код 0, конфигурация, только GET $metadata, без пароля", async () => {
    stub(BP30_METADATA_XML);
    const r = await run(["--check"], ENV);
    expect(r.code).toBe(0);
    expect(r.text).toContain("работают инструменты Бухгалтерии");
    expect(r.text).toContain("Режим: только чтение");
    expect(r.text).not.toContain(PASSWORD);
    expect(calls.map((c) => c.method)).toEqual(["GET"]);
    expect(calls[0]!.url).toContain("$metadata");
  });

  it("УТ: код 0, оперативные регистры", async () => {
    stub(UT11_METADATA_XML);
    const r = await run(["--check"], ENV);
    expect(r.code).toBe(0);
    expect(r.text).toContain("Управления торговлей");
    expect(r.text).not.toContain("нет в «Составе OData»");
  });

  it("недостающий объект называется, код остаётся 0", async () => {
    stub(
      UT11_METADATA_XML.replaceAll(
        "AccumulationRegister_РасчетыСКлиентамиПоСрокам",
        "AccumulationRegister_Другое",
      ),
    );
    const r = await run(["--check"], ENV);
    expect(r.code).toBe(0);
    expect(r.text).toContain("РасчетыСКлиентамиПоСрокам");
  });

  it("пустой состав — код 1 с подсказкой", async () => {
    stub(
      '<?xml version="1.0"?><edmx:Edmx xmlns:edmx="http://schemas.microsoft.com/ado/2007/06/edmx" Version="1.0"><edmx:DataServices><Schema xmlns="http://schemas.microsoft.com/ado/2008/09/edm" Namespace="StandardODATA"><EntityContainer Name="x"></EntityContainer></Schema></edmx:DataServices></edmx:Edmx>',
    );
    const r = await run(["--check"], ENV);
    expect(r.code).toBe(1);
    expect(r.text).toContain("«Состав OData» пуст");
  });

  it.each([
    [401, "Проверьте логин и пароль"],
    [404, "odata/standard.odata"],
  ])("ответ %i — код 1 и что проверить", async (status, hint) => {
    stub("<x/>", status);
    const r = await run(["--check"], ENV);
    expect(r.code).toBe(1);
    expect(r.text).toContain(hint);
    expect(r.text).not.toContain(PASSWORD);
  });

  it("сеть недоступна и страница входа вместо OData — код 1", async () => {
    vi.stubGlobal("fetch", async () => {
      throw new TypeError("fetch failed");
    });
    const r = await run(["--check"], ENV);
    expect(r.code).toBe(1);
    expect(r.text).toContain("сервер 1С доступен");
    stub("<html><body>login</body></html>", 200, "text/html");
    expect((await run(["--check"], ENV)).code).toBe(1);
  });

  it("настройки неполные — сообщение разбора конфигурации, код 1, без запросов", async () => {
    stub(BP30_METADATA_XML);
    const r = await run(["--check"], { ODATA_BASE_URL: ENV.ODATA_BASE_URL });
    expect(r.code).toBe(1);
    expect(r.text).toContain("не задан ODATA_PASSWORD");
    expect(calls).toEqual([]);
  });

  it("несколько баз: одна не отвечает — код 1, но проверены все", async () => {
    vi.stubGlobal("fetch", async (url: string) =>
      String(url).includes("bad")
        ? new Response("", { status: 401 })
        : new Response(BP30_METADATA_XML, { status: 200, headers: { "Content-Type": "application/xml" } }),
    );
    const env: Record<string, string> = {};
    for (const [name, host] of [
      ["A", "ok"],
      ["B", "bad"],
    ] as const) {
      env[`ODATA_DB_${name}_BASE_URL`] = `https://${host}.invalid/x/odata/standard.odata/`;
      env[`ODATA_DB_${name}_USERNAME`] = "u";
      env[`ODATA_DB_${name}_PASSWORD`] = PASSWORD;
    }
    const r = await run(["--check"], env);
    expect(r.code).toBe(1);
    expect(r.text).toContain("База «a»");
    expect(r.text).toContain("База «b»");
  });
});

describe("--check --config: блок env из конфига приложения", () => {
  const dir = mkdtempSync(join(tmpdir(), "check-cfg-"));
  const file = (name: string, content: string) => {
    const p = join(dir, name);
    writeFileSync(p, content);
    return p;
  };
  afterEach(() => undefined);
  const entry = (env: object) => ({ command: "npx", args: ["-y", "1c-odata-mcp"], env });

  it.each(["mcpServers", "servers"])("раздел %s", async (section) => {
    stub(BP30_METADATA_XML);
    const p = file(
      `${section}.json`,
      JSON.stringify({ [section]: { "1c": entry(ENV), other: { command: "x" } } }),
    );
    const r = await run(["--check", "--config", p], {});
    expect(r.code).toBe(0);
    expect(r.text).toContain("запись «1c»");
    expect(r.text).not.toContain(PASSWORD);
    expect(calls).toHaveLength(1);
  });

  // Путь из документации — «~/Library/Application Support/…» в кавычках из-за пробела: оболочка «~» не раскрывает.
  it("«~/…» в пути раскрывается в домашний каталог", async () => {
    stub(BP30_METADATA_XML);
    file("tilde.json", JSON.stringify({ mcpServers: { "1c-odata": entry(ENV) } }));
    vi.stubEnv("HOME", dir);
    vi.stubEnv("USERPROFILE", dir);
    const r = await run(["--check", "--config", "~/tilde.json"], {});
    vi.unstubAllEnvs();
    expect(r.code).toBe(0);
    expect(calls).toHaveLength(1);
  });

  it("переменные терминала не подмешиваются к конфигу", async () => {
    stub(BP30_METADATA_XML);
    const p = file(
      "only.json",
      JSON.stringify({ mcpServers: { "1c-odata": entry({ ODATA_BASE_URL: ENV.ODATA_BASE_URL }) } }),
    );
    const r = await run(["--check", "--config", p], ENV);
    expect(r.code).toBe(1);
    expect(r.text).toContain("не задан ODATA_PASSWORD");
  });

  // Ошибка разбора Node цитирует файл: при незакавыченном значении — вместе с паролем.
  it.each([
    ["лишняя запятая", `{"mcpServers":{"1c-odata":{"env":{"ODATA_PASSWORD":"${PASSWORD}",}}}}`],
    ["значение без кавычек", `{"a":1,"ODATA_PASSWORD":${PASSWORD}}`],
  ])("сломанный JSON (%s) — без фрагмента файла и пароля", async (_l, content) => {
    const p = file("broken.json", content);
    const r = await run(["--check", "--config", p], {});
    expect(r.code).toBe(1);
    expect(r.text).toContain("не разбирается как JSON");
    expect(r.text).not.toMatch(/secre|PASSWORD/);
  });

  it.each([
    ["нет файла", join(dir, "absent.json"), "не удалось открыть"],
    ["нет путь", "", "укажите путь"],
  ])("%s", async (_l, p, expected) => {
    const r = await run(p ? ["--check", "--config", p] : ["--check", "--config"], {});
    expect(r.code).toBe(1);
    expect(r.text).toContain(expected);
  });

  it("нет записи коннектора — названы имеющиеся записи, без их env", async () => {
    const p = file(
      "none.json",
      JSON.stringify({ mcpServers: { other: { command: "x", env: { S: PASSWORD } } } }),
    );
    const r = await run(["--check", "--config", p], {});
    expect(r.code).toBe(1);
    expect(r.text).toContain("other");
    expect(r.text).not.toContain(PASSWORD);
  });

  it("подстановка ${…} — объяснение, без запросов", async () => {
    stub(BP30_METADATA_XML);
    const p = file(
      "subst.json",
      JSON.stringify({ servers: { "1c-odata": entry({ ...ENV, ODATA_PASSWORD: "${input:pwd}" }) } }),
    );
    const r = await run(["--check", "--config", p], {});
    expect(r.code).toBe(1);
    expect(r.text).toContain("подстановкой");
    expect(calls).toEqual([]);
  });

  it("уборка", () => rmSync(dir, { recursive: true, force: true }));
});

describe("объекты, которые называет --check", () => {
  it("совпадают с минимальными составами из CONNECTING.md", () => {
    const doc = readFileSync("docs/CONNECTING.md", "utf8");
    for (const name of [...TRADE_REQUIRED, TRADE_CASH[0]!])
      expect(doc).toContain(name.slice(name.indexOf("_") + 1));
  });
});

const entry = join(process.cwd(), "dist", "index.js");

// Запускает собранный сервер. publish.yml гоняет тесты без сборки: без dist — пропуск (как в startup.test.ts).
describe.skipIf(!existsSync(entry))("точка входа", () => {
  const clean = { PATH: process.env.PATH ?? "" };
  const pkg = JSON.parse(readFileSync("package.json", "utf8")) as { version: string };

  it("--version и --help — без настроек, код 0", () => {
    expect(execFileSync("node", [entry, "--version"], { env: clean, encoding: "utf8" }).trim()).toBe(
      pkg.version,
    );
    expect(execFileSync("node", [entry, "--help"], { env: clean, encoding: "utf8" })).toContain("--check");
  });

  it("--check без настроек: понятное сообщение, код 1, без файла лога и мусора в stderr", () => {
    const r = spawnSync("node", [entry, "--check"], { env: clean, encoding: "utf8" });
    expect(r.status).toBe(1);
    expect(r.stdout).toContain("не задана ни одна база");
    expect(r.stderr).toBe("");
  });
});
