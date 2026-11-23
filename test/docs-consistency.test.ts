import { bundleManifest } from "../src/mcp/bundle-manifest.js";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { parseConfig } from "../src/config/env.js";
import { ServerContext } from "../src/context.js";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { dirname, join, normalize } from "node:path";
import { describe, expect, it } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { PROMPTS, renderPrompt } from "../src/mcp/prompts.js";
import { renderToolsReference } from "../src/mcp/tools-reference.js";

/**
 * Тексты, которые при релизе расходились с кодом: число инструментов в README / package.json /
 * manifest.json, состав и схемы инструментов в манифесте, версия в трёх местах, раздел CHANGELOG.
 * Если тест упал — `npm run manifest` и поправить числа в текстах.
 */
const read = (path: string) => readFileSync(path, "utf8");
const pkg = JSON.parse(read("package.json")) as { version: string; description: string };
const manifest = JSON.parse(read("manifest.json")) as {
  version: string;
  long_description: string;
  tools: Array<{ name: string; inputSchema: { properties?: Record<string, unknown>; required?: string[] } }>;
  prompts?: Array<{ name: string; arguments?: string[]; text: string }>;
};
const readme = read("README.md");
const changelog = read("CHANGELOG.md");

type Registered = Record<string, { inputSchema?: { shape?: Record<string, unknown> } }>;
const registered = (
  createServer({ db: () => undefined, databases: () => [], behavior: {} } as never) as unknown as {
    _registeredTools: Registered;
  }
)._registeredTools;
const names = Object.keys(registered).sort();
/** «63 инструмента», «60 инструментов», «61 инструмент». */
const toolsWord = (n: number) => {
  const d = n % 10;
  const dd = n % 100;
  if (d === 1 && dd !== 11) return "инструмент";
  if (d >= 2 && d <= 4 && (dd < 12 || dd > 14)) return "инструмента";
  return "инструментов";
};
const reads = names.filter((n) => n.startsWith("read.")).length;
const writes = names.filter((n) => n.startsWith("write.")).length;

describe("документация совпадает с кодом", () => {
  it("каждый инструмент — read.* или write.*", () => {
    expect(reads + writes).toBe(names.length);
  });

  it("manifest.json: тот же состав инструментов, что регистрирует сервер", () => {
    expect(manifest.tools.map((t) => t.name).sort()).toEqual(names);
  });

  it("manifest.json: инструменты целиком как в tools/list — описания и схемы, а не только имена", async () => {
    // Сверка только имён пропустила переименование поля ответа: в бандле осталась схема со старым именем.
    const ctx = new ServerContext(
      parseConfig({
        ODATA_BASE_URL: "https://example.com/odata/standard.odata/",
        ODATA_USERNAME: "u",
        ODATA_PASSWORD: "p",
        READ_ONLY: "true",
      }),
    );
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const { tools } = await client.listTools();
    const live = tools.map((t) => ({
      name: t.name,
      ...(t.title ? { title: t.title } : {}),
      description: t.description,
      inputSchema: t.inputSchema,
      ...(t.outputSchema ? { outputSchema: t.outputSchema } : {}),
      ...(t.annotations ? { annotations: t.annotations } : {}),
    }));
    expect(JSON.parse(JSON.stringify(live)), "manifest.json устарел: npm run manifest").toEqual(
      manifest.tools,
    );
  });

  it("manifest.json: operationId в схемах там же, где в коде (иначе манифест не пересобран)", () => {
    const withOperationId = (n: string) => "operationId" in (registered[n]!.inputSchema?.shape ?? {});
    const inManifest = (n: string) =>
      "operationId" in (manifest.tools.find((t) => t.name === n)!.inputSchema.properties ?? {});
    expect(names.filter(inManifest)).toEqual(names.filter(withOperationId));
  });

  it("bank settlement schema exposes optional explicit CP without an account override", () => {
    const name = "write.document.apply_bank_settlement_settings";
    const schema = manifest.tools.find((t) => t.name === name)!.inputSchema;
    expect(schema.properties?.counterpartyRef).toMatchObject({ type: "string", format: "uuid" });
    expect(schema.required).not.toContain("counterpartyRef");
    expect(Object.keys(schema.properties ?? {}).sort()).toEqual([
      "confirm",
      "contractRef",
      "counterpartyRef",
      "database",
      "entitySet",
      "expectedDataVersion",
      "operationId",
      "ref",
      "side",
    ]);
    expect(registered[name]!.inputSchema?.shape).toHaveProperty("counterpartyRef");
  });

  it("версия одна в package.json и manifest.json, в CHANGELOG есть её раздел", () => {
    expect(manifest.version).toBe(pkg.version);
    expect(changelog).toContain(`## [${pkg.version}]`);
    expect(changelog).toContain(`[${pkg.version}]: `);
  });

  it("CHANGELOG: один раздел «[Не выпущено]» на русском", () => {
    expect(changelog, "раздел «Unreleased» — переименуйте в «[Не выпущено]»").not.toMatch(
      /^## \[?Unreleased\]?/im,
    );
    expect(changelog.match(/^## \[Не выпущено\]/gm) ?? []).toHaveLength(1);
    const start = changelog.indexOf("## [Не выпущено]");
    const end = changelog.indexOf("\n## [", start + 1);
    const section = changelog.slice(start, end < 0 ? undefined : end);
    // Каждая строка с текстом — с русскими словами: код и имена в `кавычках` не считаются.
    const english = section
      .split("\n")
      .filter((l) => /[A-Za-z]{3}/.test(l.replace(/`[^`]*`|\[[^\]]*\]\([^)]*\)|https?:\S+/g, "")))
      .filter((l) => !/[А-Яа-яЁё]/.test(l));
    expect(english, "строки CHANGELOG без русского текста").toEqual([]);
  });

  it("CHANGELOG после 1.0.0 — компактный формат (CONTRIBUTING, «История изменений»)", () => {
    // Разделы 1.0.0 и старше выпущены и на них ссылаются GitHub Releases — их не переписываем.
    const recent = changelog.slice(changelog.indexOf("## [Не выпущено]"), changelog.indexOf("\n## [1.0.0]"));
    // Без заголовка «[Не выпущено]» или раздела 1.0.0 срез пуст, и проверять было бы нечего.
    expect(recent.length, "в CHANGELOG нет «## [Не выпущено]» выше «## [1.0.0]»").toBeGreaterThan(0);
    const ORDER = [
      "Что изменится при обновлении",
      "Добавлено",
      "Исправлено",
      "Удалено",
      "Безопасность",
      "Прочее",
    ];
    const problems: string[] = [];
    for (const section of recent.split(/\n(?=## \[)/)) {
      const version = /^## \[([^\]]+)\]/.exec(section)?.[1] ?? "?";
      const bullets = section.split("\n").filter((l) => l.startsWith("- "));
      if (bullets.length === 0) continue;
      const at = (m: string) => problems.push(`[${version}] ${m}`);
      if (section.length > 5000) at(`раздел ${section.length} знаков, предел 5000`);
      const body = section.split("\n").slice(1).join("\n").trimStart();
      if (!body.startsWith("**Главное.**")) at("нет абзаца «**Главное.**» в начале раздела");
      const headings = [...section.matchAll(/^### (.+)$/gm)].map((m) => m[1]!.trim());
      for (const h of headings) if (!ORDER.includes(h)) at(`заголовок «${h}» не из списка`);
      const known = headings.filter((h) => ORDER.includes(h));
      if (known.join() !== [...known].sort((a, b) => ORDER.indexOf(a) - ORDER.indexOf(b)).join())
        at(`порядок заголовков: ${known.join(" → ")}`);
      for (const part of section.split(/^### /m).slice(1)) {
        const list = part.slice(part.indexOf("\n") + 1).trim();
        if (/^- .*\n\s*\n- /m.test(list)) at(`пустая строка внутри списка «${part.split("\n")[0]}»`);
      }
      const leads = new Map<string, number>();
      for (const b of bullets) {
        if (b.length > 350) at(`пункт ${b.length} знаков, предел 350: ${b.slice(0, 60)}…`);
        if (/ревью/i.test(b)) at(`ревью — внутренняя кухня: ${b.slice(0, 60)}…`);
        const lead = /^- \*\*(.+?)\*\*/.exec(b)?.[1]?.toLowerCase();
        if (lead) leads.set(lead, (leads.get(lead) ?? 0) + 1);
      }
      for (const [lead, n] of leads) if (n > 1) at(`один факт в ${n} пунктах: «${lead}»`);
    }
    expect(problems).toEqual([]);
  });

  it("число инструментов в README, package.json и manifest.json", () => {
    expect(readme).toContain(
      `${names.length} ${toolsWord(names.length)} (${reads} чтение/аналитика + ${writes} записей)`,
    );
    // Родительный падеж: «у всех 81 инструмента», «у всех 80 инструментов».
    const genitive = names.length % 10 === 1 && names.length % 100 !== 11 ? "инструмента" : "инструментов";
    expect(readme).toContain(`У всех ${names.length} ${genitive}`);
    expect(pkg.description).toContain(`${names.length} tools`);
    expect(manifest.long_description).toContain(`${names.length} ${toolsWord(names.length)}`);
    // Объём описаний для профиля инструментов — там же, где число: CONNECTING и .env.example (аудит: отставали).
    expect(read("docs/CONNECTING.md")).toContain(`всех ${names.length} ${genitive}`);
    expect(read("docs/INTEGRATIONS.md")).toContain(`У всех ${names.length} ${genitive}`);
    expect(read(".env.example")).toContain(`всех ${names.length} —`);
  });

  it("README заканчивается «Лицензией»: новые разделы — в тело README или в docs/", () => {
    const tail = readme.slice(readme.lastIndexOf("\n## "));
    expect(tail.startsWith("\n## Лицензия"), "последний раздел README — не «Лицензия»").toBe(true);
    expect(tail.match(/^#{2,} .*/gm), "подразделы после «Лицензии»").toEqual(["## Лицензия"]);
  });

  it("каждый инструмент упомянут в обзоре docs/TOOLS.md полным именем", () => {
    const tools = read("docs/TOOLS.md");
    expect(names.filter((n) => !tools.includes(`\`${n}\``))).toEqual([]);
  });

  it("справочник параметров docs/TOOLS-REFERENCE.md собран из manifest.json (npm run manifest)", () => {
    expect(read("docs/TOOLS-REFERENCE.md")).toBe(renderToolsReference(manifest.tools as never));
  });

  it("сценарии (prompts): manifest.json совпадает с кодом, каждый описан в README", () => {
    const fromCode = PROMPTS.map((p) => ({
      name: p.name,
      arguments: p.arguments.map((a) => a.name),
      text: p.text,
    }));
    expect(manifest.prompts?.map(({ name, arguments: a, text }) => ({ name, arguments: a, text }))).toEqual(
      fromCode,
    );
    expect(PROMPTS.filter((p) => !readme.includes(`\`${p.name}\``)).map((p) => p.name)).toEqual([]);
  });

  it("сценарии: в тексте только объявленные аргументы и существующие инструменты", () => {
    for (const p of PROMPTS) {
      const used = [...p.text.matchAll(/\$\{arguments\.([A-Za-z_]+)\}/g)].map((m) => m[1]);
      expect(used.filter((u) => !p.arguments.some((a) => a.name === u))).toEqual([]);
      const tools = [...p.text.matchAll(/\b((?:read|write)\.[a-z_]+\.[a-z_]+)/g)].map((m) => m[1]);
      expect(tools.filter((t) => !names.includes(t!))).toEqual([]);
      expect(renderPrompt(p, {})).not.toContain("${arguments.");
    }
  });
});

describe("манифест расширения .mcpb", () => {
  it("в tools только name и description: остальное Claude Desktop отвергает", () => {
    const bundled = bundleManifest(manifest);
    expect(bundled.tools).toHaveLength(manifest.tools.length);
    for (const tool of bundled.tools ?? []) expect(Object.keys(tool).sort()).toEqual(["description", "name"]);
    // Исходный манифест со схемами нужен тестам этого файла.
    expect(manifest.tools[0]).toHaveProperty("inputSchema");
  });
});

/** Якорь заголовка, как его строит GitHub: строчные, без знаков препинания, пробел — дефис. */
const slug = (heading: string) =>
  heading
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s_-]/gu, "")
    .replace(/\s/g, "-");
const anchorsOf = (md: string) =>
  new Set(
    md
      .replace(/```[\s\S]*?```/g, "")
      .split("\n")
      .filter((l) => /^#{1,6} /.test(l))
      .map((l) => slug(l.replace(/^#+ /, ""))),
  );

describe("ссылки между документами", () => {
  const files = [
    "README.md",
    "README.en.md",
    "CONTRIBUTING.md",
    "SECURITY.md",
    ...readdirSync("docs")
      .filter((f) => f.endsWith(".md"))
      .map((f) => `docs/${f}`),
  ];
  it("каждая ссылка с якорем ведёт на существующий заголовок", () => {
    const broken: string[] = [];
    for (const file of files) {
      const text = read(file);
      for (const [, target, anchor] of text.matchAll(/\]\(([^)#\s]*)#([^)\s]+)\)/g)) {
        if (/^https?:/.test(target!)) continue;
        const path = target ? normalize(join(dirname(file), target)) : file;
        if (!path.endsWith(".md")) continue;
        if (!anchorsOf(read(path)).has(decodeURIComponent(anchor!)))
          broken.push(`${file} → ${path}#${anchor}`);
      }
    }
    expect(broken).toEqual([]);
  });

  it("каждая относительная ссылка ведёт на существующий файл", () => {
    const missing: string[] = [];
    for (const file of files) {
      const text = read(file).replace(/```[\s\S]*?```/g, "");
      for (const [, target] of text.matchAll(/\]\(([^)#\s]+)(?:#[^)\s]*)?\)/g)) {
        if (/^(https?:|mailto:)/.test(target!)) continue;
        if (!existsSync(normalize(join(dirname(file), decodeURIComponent(target!)))))
          missing.push(`${file} → ${target}`);
      }
    }
    expect(missing).toEqual([]);
  });

  it("каждая переменная окружения, которую читает код, описана в .env.example (справочник настроек)", () => {
    const env = read(".env.example");
    const sources = (dir: string): string[] =>
      readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
        e.isDirectory() ? sources(join(dir, e.name)) : e.name.endsWith(".ts") ? [join(dir, e.name)] : [],
      );
    const used = new Set<string>();
    for (const file of sources("src").filter((f) => !f.includes("/scripts/")))
      // env.X, имена в кавычках и ключи схемы конфигурации («  ODATA_TIMEOUT_MS: z…»): без ключей схемы
      // удалённый из .env.example ODATA_TIMEOUT_MS не ловился (ревью документации, раунд 2).
      for (const [, name] of read(file).matchAll(
        /(?:env\.|["'`]|^\s+)((?:ODATA|MCP|TOOLS|LOG)_[A-Z0-9_]+|READ_ONLY)\b/gm,
      ))
        if (name !== "MCP_PATH") used.add(name!);
    expect([...used].filter((n) => !env.includes(n)).sort()).toEqual([]);
  });

  it("примеры конфига npx одинаковы и только на чтение: ни одна копия не отстаёт", () => {
    // Конфиг повторяется в README, README.en и INSTALL-WITH-AI (основная копия). Копия без TOOLS_PROFILE=read
    // или с переменной записи незаметно расходится с остальными.
    const envs: Array<{ file: string; env: Record<string, string> }> = [];
    for (const file of files) {
      for (const [, body] of read(file).matchAll(/```json\n([\s\S]*?)```/g)) {
        if (!body!.includes('"1c-odata-mcp"')) continue;
        const cfg = JSON.parse(body!) as Record<string, Record<string, { env: Record<string, string> }>>;
        const server = (cfg.mcpServers ?? cfg.servers)!["1c-odata"]!;
        envs.push({ file, env: server.env });
      }
      for (const [, cmd] of read(file).matchAll(/^(?:claude|codex) mcp add 1c-odata (.*-- npx.*)$/gm)) {
        const env = Object.fromEntries([...cmd!.matchAll(/--env (\w+)=(\S+)/g)].map(([, k, v]) => [k!, v!]));
        envs.push({ file, env });
      }
    }
    for (const file of files)
      for (const [, body] of read(file).matchAll(/```toml\n([\s\S]*?)```/g))
        if (body!.includes('"1c-odata-mcp"')) expect(body, file).toContain('TOOLS_PROFILE = "read"');
    expect(envs.length).toBeGreaterThanOrEqual(5);
    for (const { file, env } of envs) {
      expect(Object.keys(env).sort(), file).toEqual([
        "ODATA_BASE_URL",
        "ODATA_PASSWORD",
        "ODATA_USERNAME",
        "TOOLS_PROFILE",
      ]);
      expect(env.TOOLS_PROFILE, file).toBe("read");
    }
    // Значение принимает разбор настроек и прячет запись.
    const cfg = parseConfig({
      ODATA_BASE_URL: "https://example.com/odata/standard.odata/",
      ODATA_USERNAME: "u",
      ODATA_PASSWORD: "p",
      TOOLS_PROFILE: "read",
    });
    expect(cfg.behavior.toolsProfile).toBe("read");
  });

  it("README короткий: подробности — в docs/, README — вход и карта", () => {
    expect(read("README.md").split("\n").length).toBeLessThanOrEqual(250);
  });
});
