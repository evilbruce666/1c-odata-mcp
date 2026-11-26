import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";

// Цепочка поставки: ядро встраивают в чужие продукты, поэтому сборка и выпуск не должны зависеть от того, куда
// завтра передвинут тег стороннего действия, а токены задач — давать больше, чем нужно.
const dir = ".github/workflows";
const workflows = readdirSync(dir)
  .filter((f) => f.endsWith(".yml"))
  .map((f) => ({ name: f, text: readFileSync(`${dir}/${f}`, "utf8") }));

describe("workflows: цепочка поставки", () => {
  it.each(workflows.map((w) => [w.name, w.text]))(
    "%s: каждое действие закреплено хешем коммита с версией в комментарии",
    (_, text) => {
      const uses = [...text.matchAll(/^\s*-?\s*uses:\s*(.+)$/gm)].map((m) => m[1]!.trim());
      expect(uses.length).toBeGreaterThan(0);
      for (const u of uses) expect(u).toMatch(/^[\w.-]+\/[\w./-]+@[0-9a-f]{40} # v\d+(\.\d+)*$/);
    },
  );

  it.each(workflows.map((w) => [w.name, w.text]))(
    "%s: права токена заданы явно на уровне файла",
    (_, text) => {
      expect(text).toMatch(/^permissions:/m);
      expect(text).not.toMatch(/permissions:\s*write-all/);
    },
  );

  it("одна версия каждого действия во всех файлах", () => {
    const pins = new Map<string, Set<string>>();
    for (const w of workflows)
      for (const m of w.text.matchAll(/uses:\s*([\w.-]+\/[\w.-]+)[\w./-]*@([0-9a-f]{40})/g))
        (pins.get(m[1]!) ?? pins.set(m[1]!, new Set()).get(m[1]!)!).add(m[2]!);
    for (const [action, shas] of pins) expect([...shas], action).toHaveLength(1);
  });

  it("статический анализ (CodeQL) и Scorecard подключены", () => {
    const names = workflows.map((w) => w.name);
    expect(names).toContain("codeql.yml");
    expect(names).toContain("scorecard.yml");
    const codeql = workflows.find((w) => w.name === "codeql.yml")!.text;
    expect(codeql).toContain("languages: javascript-typescript");
    expect(codeql).toMatch(/pull_request:/);
  });

  it("релиз прикладывает SBOM только рабочих зависимостей, собранный после npm ci", () => {
    const publish = workflows.find((w) => w.name === "publish.yml")!.text;
    const sbom = publish.indexOf("npm sbom --sbom-format cyclonedx --omit dev > sbom.cdx.json");
    expect(sbom).toBeGreaterThan(publish.indexOf("npm ci"));
    expect(publish.indexOf("gh release create")).toBeGreaterThan(sbom);
    expect(publish).toMatch(/gh release create[^\n]*sbom\.cdx\.json/);
  });
});
