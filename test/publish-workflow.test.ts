import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

// Публикация в реестр MCP из publish.yml: у задачи есть OIDC-токен, а скачанный бинарник исполняется — проверяем,
// что он закреплён и сверяется, а сбой реестра не мешает выпуску в npm.
const yml = readFileSync(".github/workflows/publish.yml", "utf8");
const job = yml.slice(yml.indexOf("\n  registry:\n"));

describe("publish.yml: реестр MCP", () => {
  it("отдельная задача после npm, без права записи в репозиторий", () => {
    expect(yml).toContain("\n  registry:\n");
    expect(job).toContain("needs: publish");
    const permissions = /permissions:\n((?: {6}.*\n)+)/.exec(job)![1]!;
    expect(permissions.match(/^ {6}[\w-]+: \w+/gm)!.map((l) => l.trim())).toEqual([
      "contents: read",
      "id-token: write",
    ]);
    expect(job).not.toContain("secrets.");
  });

  it("mcp-publisher закреплён версией и sha256, сверка — до распаковки и запуска", () => {
    expect(job).toMatch(/MCP_PUBLISHER_VERSION: v\d+\.\d+\.\d+\n/);
    expect(job).toMatch(/MCP_PUBLISHER_SHA256: [0-9a-f]{64}\n/);
    expect(job).not.toContain("releases/latest");
    expect(job).toContain("set -euo pipefail");
    const check = job.indexOf("sha256sum --check");
    expect(check).toBeGreaterThan(0);
    expect(job.indexOf("tar xzf")).toBeGreaterThan(check);
    expect(job.indexOf("./mcp-publisher")).toBeGreaterThan(check);
  });

  it("сначала ждёт пакет в npm", () => {
    expect(job.indexOf("registry.npmjs.org/1c-odata-mcp/-/")).toBeGreaterThan(0);
    expect(job.indexOf("registry.npmjs.org")).toBeLessThan(job.indexOf("mcp-publisher publish"));
  });
});
