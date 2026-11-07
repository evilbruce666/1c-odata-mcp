import { bundleManifest } from "../src/mcp/bundle-manifest.js";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createServer } from "../src/mcp/server.js";
import { PROMPTS, renderPrompt } from "../src/mcp/prompts.js";

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

  it("число инструментов в README, package.json и manifest.json", () => {
    expect(readme).toContain(
      `${names.length} ${toolsWord(names.length)} (${reads} чтение/аналитика + ${writes} записей)`,
    );
    expect(readme).toContain(`У всех ${names.length} инструментов`);
    expect(pkg.description).toContain(`${names.length} tools`);
    expect(manifest.long_description).toContain(`${names.length} ${toolsWord(names.length)}`);
  });

  it("README заканчивается «Лицензией»: новые разделы — в тело README или в docs/", () => {
    const tail = readme.slice(readme.lastIndexOf("\n## "));
    expect(tail.startsWith("\n## Лицензия"), "последний раздел README — не «Лицензия»").toBe(true);
    expect(tail.match(/^#{2,} .*/gm), "подразделы после «Лицензии»").toEqual(["## Лицензия"]);
  });

  it("каждый инструмент упомянут в README полным именем", () => {
    expect(names.filter((n) => !readme.includes(`\`${n}\``))).toEqual([]);
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
