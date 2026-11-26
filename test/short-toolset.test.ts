import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { PROMPTS, promptTools } from "../src/mcp/prompts.js";

/**
 * Короткие наборы для моделей с маленьким контекстом (docs/CONNECTING.md), свой на каждую конфигурацию: имена
 * существуют, только чтение, каждый инструмент проверен вживую именно в ней, а заявленный объём описаний не отстаёт от
 * настоящего. Наборы берутся из самой документации. Прежний единый набор годился только наполовину: в УТ отказывали
 * 8 из 18, долгов УТ в нём не было, а в Бухгалтерии отказывал get_warehouse_stock.
 */
const doc = readFileSync("docs/CONNECTING.md", "utf8");
const set = (title: string) =>
  new RegExp(`\\*\\*Набор для ${title}[\\s\\S]*?\`\`\`text\\nTOOLS_ALLOW=([\\w.,]+)\\n`).exec(doc)?.[1];
const manifest = JSON.parse(readFileSync("manifest.json", "utf8")) as {
  tools: Array<{ name: string; description: string; inputSchema: unknown }>;
};
const size = (names: (n: string) => boolean) =>
  JSON.stringify(
    manifest.tools
      .filter((t) => names(t.name))
      .map((t) => ({ name: t.name, description: t.description, inputSchema: t.inputSchema })),
  ).length;
/** Инструмент → [клетка Бухгалтерии, клетка УТ] из таблиц SUPPORT.md. */
const support = new Map(
  [...readFileSync("docs/SUPPORT.md", "utf8").matchAll(/^\| `([\w.]+)` \| ([^|]+) \| ([^|]+) \|$/gm)].map(
    (m) => [m[1]!, [m[2]!, m[3]!] as const],
  ),
);

const SETS = [
  { title: "Бухгалтерии", column: 0, prompts: PROMPTS.map((p) => p.name) },
  { title: "Управления торговлей", column: 1, prompts: ["sales_review"] },
] as const;

describe.each(SETS)("набор для маленького контекста: $title", ({ title, column, prompts }) => {
  const line = set(title);
  const short = new Set(line?.split(",") ?? []);
  const known = new Set(manifest.tools.map((t) => t.name));

  it("есть в CONNECTING.md, все имена существуют и только читают", () => {
    expect(short.size).toBeGreaterThanOrEqual(10);
    expect([...short].filter((n) => !known.has(n) || !n.startsWith("read."))).toEqual([]);
  });

  it("каждый инструмент проверен вживую в этой конфигурации (SUPPORT.md)", () => {
    expect([...short].filter((n) => !support.get(n)?.[column].startsWith("вживую"))).toEqual([]);
  });

  it("сценарии этой конфигурации доступны", () => {
    for (const p of PROMPTS.filter((x) => (prompts as readonly string[]).includes(x.name)))
      expect(
        promptTools(p).filter((t) => !short.has(t)),
        p.name,
      ).toEqual([]);
  });

  it("объём набора, названный в документации, — в пределах 10 % от настоящего", () => {
    const said = Number(new RegExp(`\\*\\*Набор для ${title}[\\s\\S]*?около\\s+(\\d+)\\s+тыс`).exec(doc)![1]);
    const real = size((n) => short.has(n));
    expect(Math.abs(said * 1_000 - real) / real).toBeLessThan(0.1);
  });
});

describe("объёмы всех инструментов, названные в документации", () => {
  it("в пределах 10 % от настоящих", () => {
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
  });
});
