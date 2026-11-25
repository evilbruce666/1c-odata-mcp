import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { PROMPTS, promptTools } from "../src/mcp/prompts.js";

/**
 * Короткий набор для моделей с маленьким контекстом (docs/CONNECTING.md): имена существуют, только чтение, все
 * сценарии остаются, а заявленный объём описаний не отстаёт от настоящего. Набор берётся из самой документации.
 */
const doc = readFileSync("docs/CONNECTING.md", "utf8");
const line = /^TOOLS_ALLOW=((?:read\.[\w.]+,){9,}read\.[\w.]+)$/m.exec(doc)?.[1];
const manifest = JSON.parse(readFileSync("manifest.json", "utf8")) as {
  tools: Array<{ name: string; description: string; inputSchema: unknown }>;
};
const size = (names: (n: string) => boolean) =>
  JSON.stringify(
    manifest.tools
      .filter((t) => names(t.name))
      .map((t) => ({ name: t.name, description: t.description, inputSchema: t.inputSchema })),
  ).length;

describe("набор для маленького контекста", () => {
  const short = new Set(line?.split(",") ?? []);
  const known = new Set(manifest.tools.map((t) => t.name));

  it("есть в CONNECTING.md, все имена существуют и только читают", () => {
    expect(short.size).toBeGreaterThanOrEqual(10);
    expect([...short].filter((n) => !known.has(n) || !n.startsWith("read."))).toEqual([]);
  });

  it("все сценарии остаются доступны", () => {
    for (const p of PROMPTS)
      expect(
        promptTools(p).filter((t) => !short.has(t)),
        p.name,
      ).toEqual([]);
  });

  it("объёмы, названные в документации, — в пределах 10 % от настоящих", () => {
    const claim = (re: RegExp) => Number(re.exec(doc)![1]) * 1_000;
    const near = (said: number, real: number) => expect(Math.abs(said - real) / real).toBeLessThan(0.1);
    near(
      claim(/всех \d+ инструмент\S*\s+—\s+около\s+(\d+)\s+тыс/),
      size(() => true),
    );
    near(
      claim(/только чтения\s+—\s+около\s+(\d+)\s+тыс/),
      size((n) => n.startsWith("read.")),
    );
    near(
      claim(/этого набора\s+—\s+около\s+(\d+)\s+тыс/),
      size((n) => short.has(n)),
    );
  });
});
