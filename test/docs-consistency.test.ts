import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createServer } from "../src/mcp/server.js";

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
  tools: Array<{ name: string; inputSchema: { properties?: Record<string, unknown> } }>;
};
const readme = read("README.md");
const changelog = read("CHANGELOG.md");

type Registered = Record<string, { inputSchema?: { shape?: Record<string, unknown> } }>;
const registered = (
  createServer({ db: () => undefined } as never) as unknown as { _registeredTools: Registered }
)._registeredTools;
const names = Object.keys(registered).sort();
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

  it("версия одна в package.json и manifest.json, в CHANGELOG есть её раздел", () => {
    expect(manifest.version).toBe(pkg.version);
    expect(changelog).toContain(`## [${pkg.version}]`);
    expect(changelog).toContain(`[${pkg.version}]: `);
  });

  it("число инструментов в README, package.json и manifest.json", () => {
    expect(readme).toContain(`${names.length} инструментов (${reads} чтение/аналитика + ${writes} записей)`);
    expect(readme).toContain(`У всех ${names.length} инструментов`);
    expect(pkg.description).toContain(`${names.length} tools`);
    expect(manifest.long_description).toContain(`${names.length} инструментов`);
  });

  it("каждый инструмент упомянут в README полным именем", () => {
    expect(names.filter((n) => !readme.includes(`\`${n}\``))).toEqual([]);
  });
});
