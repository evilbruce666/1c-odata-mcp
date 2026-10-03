import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

const entry = join(process.cwd(), "dist", "index.js");

// Проверяет собранный сервер (в CI build идёт до test); без dist — пропуск.
describe.skipIf(!existsSync(entry))("старт с ошибкой конфигурации", () => {
  it("под MCP-клиентом (stdin = pipe) сообщает причину в stderr и выходит с кодом 1", () => {
    const env: NodeJS.ProcessEnv = {
      PATH: process.env.PATH,
      TMPDIR: mkdtempSync(join(tmpdir(), "1c-start-")),
    };
    const r = spawnSync(process.execPath, [entry], { env, input: "", encoding: "utf8", timeout: 20_000 });
    expect(r.status).toBe(1);
    expect(r.stderr).toContain("не задана ни одна база");
    expect(r.stderr).not.toContain("sonic boom");
  });
});
