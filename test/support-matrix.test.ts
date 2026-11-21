import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

/**
 * docs/SUPPORT.md — мерило «полной поддержки»: у каждого инструмента обе клетки (Бухгалтерия, УТ) заполнены одним из
 * пяти видов, а итог по видам совпадает с таблицами. Новый инструмент без строки или пустая клетка — красный тест.
 */
const KINDS = ["вживую", "отказ по замыслу", "исключение", "только подделка", "не сделано"] as const;
const DOC = readFileSync("docs/SUPPORT.md", "utf8");

async function toolNames(): Promise<string[]> {
  const fake = fake1C({ READ_ONLY: "true" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  return (await client.listTools()).tools.map((t) => t.name);
}

const rows = [...DOC.matchAll(/^\| `([a-z_.]+)` \| (.*) \| (.*) \|$/gm)].map((m) => ({
  tool: m[1]!,
  cells: [m[2]!.trim(), m[3]!.trim()] as const,
}));
const kindOf = (cell: string) => KINDS.find((k) => cell.startsWith(k));

describe("docs/SUPPORT.md — поддержка по конфигурациям", () => {
  it("каждый инструмент — ровно одна строка, лишних нет", async () => {
    const names = (await toolNames()).sort();
    expect(rows.map((r) => r.tool).sort()).toEqual(names);
  });

  it("обе клетки заполнены одним из пяти видов; у всех, кроме «не сделано», причина или база с датой", () => {
    for (const { tool, cells } of rows)
      for (const cell of cells) {
        const kind = kindOf(cell);
        expect(kind, `${tool}: «${cell}»`).toBeDefined();
        if (kind === "вживую") expect(cell, tool).toMatch(/^вживую \(.+, \d{2}\.\d{2}\.\d{4}.*\)$/);
        else if (kind === "исключение")
          expect(cell, tool).toMatch(/^исключение \(решение владельца \d{2}\.\d{2}\.\d{4}: .{5,}\)$/);
        else if (kind !== "не сделано") expect(cell, tool).toMatch(/^[^(]+ \(.{5,}\)$/);
      }
  });

  it("итог по видам совпадает с таблицами", () => {
    for (const [i, label] of ["Бухгалтерия 3.0", "Управление торговлей 11"].entries()) {
      const line = new RegExp(`^- ${label}: (.+)\\.$`, "m").exec(DOC)?.[1];
      expect(line, label).toBeDefined();
      const expected = KINDS.map(
        (k) => `${k} ${rows.filter((r) => kindOf(r.cells[i as 0 | 1]) === k).length}`,
      );
      expect(line).toBe(expected.join(", "));
    }
  });
});
