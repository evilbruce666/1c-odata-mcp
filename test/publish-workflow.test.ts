import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

// Публикация в реестр MCP: у задачи есть OIDC-токен, а скачанный бинарник исполняется — проверяем, что он закреплён
// и сверяется, сбой реестра не мешает выпуску в npm, а токен задачи не годится для npm (доверие там — по файлу
// publish.yml).
const npm = readFileSync(".github/workflows/publish.yml", "utf8");
const job = readFileSync(".github/workflows/publish-registry.yml", "utf8");

describe("publish-registry.yml: реестр MCP", () => {
  it("в publish.yml реестра нет: OIDC-токен этого файла принимает npm", () => {
    expect(npm).not.toContain("mcp-publisher");
    expect(npm.match(/^ {2}[\w-]+:\n {4}runs-on/gm)).toHaveLength(1);
  });

  it("запускается только после успешной публикации по тегу своим publish.yml", () => {
    const name = /^name: (.+)$/m.exec(npm)![1]!;
    expect(job).toContain(`workflows: ["${name}"]`);
    expect(job).not.toMatch(/^ {2}(push|pull_request|pull_request_target):/m);
    for (const guard of [
      "workflow_run.conclusion == 'success'",
      "workflow_run.event == 'push'",
      "workflow_run.path == '.github/workflows/publish.yml'",
      "workflow_run.head_repository.full_name == github.repository",
      "startsWith(github.event.workflow_run.head_branch, 'v')",
    ])
      expect(job).toContain(guard);
    expect(job).toContain("ref: ${{ github.event.workflow_run.head_sha }}");
    expect(job).toContain("persist-credentials: false");
    expect(job).not.toMatch(/\bnpm (ci|install|test|run)\b|\bnpx\b/);
  });

  it("без права записи в репозиторий и без секретов", () => {
    expect(job).toContain("\npermissions: {}\n");
    const permissions = /\n {4}permissions:\n((?: {6}.*\n)+)/.exec(job)![1]!;
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

  it("сначала ждёт пакет в npm — версию из package.json выпущенного коммита", () => {
    expect(job).toContain('version="$(jq -r .version package.json)"');
    expect(job).not.toContain("GITHUB_REF_NAME");
    expect(job.indexOf("registry.npmjs.org/1c-odata-mcp/-/")).toBeGreaterThan(0);
    expect(job.indexOf("registry.npmjs.org")).toBeLessThan(job.indexOf("mcp-publisher publish"));
  });
});
