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

  it("запускается только пушем тега, без кода по ссылке из события (Scorecard: Dangerous-Workflow)", () => {
    expect(job).toMatch(/^on:\n {2}push:\n {4}tags: \["v\*"\]\n/m);
    expect(job).not.toMatch(/workflow_run|pull_request|branches:/);
    expect(job).not.toMatch(/\bref: \$\{\{/);
    expect(job).toContain("persist-credentials: false");
    // Тег должен равняться версии из package.json — и до публикации.
    expect(job).toContain('[ "$GITHUB_REF_NAME" = "v$version" ]');
    expect(job.indexOf('[ "$GITHUB_REF_NAME"')).toBeLessThan(job.indexOf("mcp-publisher publish"));
    // Значение из события — только через переменную окружения, не подстановкой в скрипт.
    expect(job).not.toMatch(/run: [|>][^]*\$\{\{ *github\.(event|head_ref)/);
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

  it("сначала ждёт пакет в npm: не опубликован — реестр не трогается", () => {
    expect(job).toContain('version="$(jq -r .version package.json)"');
    expect(job.indexOf("registry.npmjs.org/1c-odata-mcp/-/")).toBeGreaterThan(0);
    expect(job.indexOf("registry.npmjs.org")).toBeLessThan(job.indexOf("mcp-publisher login"));
    // Ожидание кончается ошибкой, а не публикацией: «exit 0» — только когда пакет найден.
    const wait = job.slice(job.indexOf("Дождаться пакета в npm"), job.indexOf("Скачать и проверить"));
    expect(wait).toMatch(/registry\.npmjs\.org[^\n]*&& exit 0/);
    expect(wait.match(/exit \d/g)!.at(-1)).toBe("exit 1");
  });

  it("publish.yml: npm из комплекта Node, без установки поверх (Scorecard: Pinned-Dependencies)", () => {
    expect(npm).not.toMatch(/npm install -g/);
    expect(npm).toMatch(/node-version: 24\n/);
  });
});

// Подпись релиза: у задачи OIDC-токен и право записи в релиз — поэтому код репозитория не исполняется, тег
// проверяется до всего остального, а подписывается только архив, совпавший с npm.
const sign = readFileSync(".github/workflows/sign-release.yml", "utf8");

describe("sign-release.yml: подписи файлов релиза", () => {
  it("запускается тегом или вручную; тег из события — только через окружение и проверяется первым", () => {
    expect(sign).toMatch(/^on:\n {2}push:\n {4}tags: \["v\*"\]\n {2}workflow_dispatch:\n/m);
    expect(sign).not.toMatch(/workflow_run|pull_request|branches:/);
    expect(sign).toMatch(
      /\n {6}TAG: \$\{\{ github\.event_name == 'workflow_dispatch' && inputs\.tag \|\| github\.ref_name \}\}\n/,
    );
    expect(sign).not.toMatch(/run: [|>][^]*\$\{\{/);
    const check = sign.indexOf('[[ "$TAG" =~ ^v[0-9]+\\.[0-9]+\\.[0-9]+$ ]]');
    expect(check).toBeGreaterThan(0);
    for (const later of ["curl", "gh release", "cosign-installer"])
      expect(sign.indexOf(later)).toBeGreaterThan(check);
  });

  it("код репозитория не скачивается и не исполняется, секретов нет", () => {
    expect(sign).not.toMatch(/actions\/checkout|\bnpm (ci|install|test|run|publish)\b|\bnpx\b/);
    expect(sign).not.toContain("secrets.");
  });

  it("права — только запись в релиз и OIDC", () => {
    expect(sign).toContain("\npermissions: {}\n");
    const permissions = /\n {4}permissions:\n((?: {6}.*\n)+)/.exec(sign)![1]!;
    expect(permissions.match(/^ {6}[\w-]+: \w+/gm)!.map((l) => l.trim())).toEqual([
      "contents: write",
      "id-token: write",
    ]);
  });

  it("cosign закреплён версией; архив сверен с npm до подписи; каждая подпись проверяется", () => {
    expect(sign).toMatch(/cosign-release: v\d+\.\d+\.\d+\n/);
    const integrity = sign.indexOf('[ "$actual" = "$expected" ]');
    expect(integrity).toBeGreaterThan(sign.indexOf(".dist.integrity"));
    expect(sign.indexOf("cosign sign-blob")).toBeGreaterThan(integrity);
    expect(sign.indexOf("cosign verify-blob")).toBeGreaterThan(sign.indexOf("cosign sign-blob"));
    expect(sign).toContain("--certificate-oidc-issuer https://token.actions.githubusercontent.com");
    // Scorecard (Signed-Releases) узнаёт подпись по расширению файла в релизе.
    expect(sign).toMatch(/--bundle "\$f\.sigstore\.json"/);
    expect(sign).toMatch(/gh release upload[^\n]*\*\.sigstore\.json/);
  });

  it("publish.yml и реестр подписи не трогают: сбой подписи не мешает выпуску", () => {
    expect(npm).not.toContain("cosign");
    expect(job).not.toContain("cosign");
  });
});
