import { mkdtempSync, writeFileSync, chmodSync, statSync, rmSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { logFile } from "../src/logger.js";

// Лог stdio писался с правами по umask (обычно 0644): на Linux с общим /tmp его читали другие пользователи, а в нём
// тексты ошибок 1С с названиями и ИНН (аудит 25.11).
const posix = process.platform !== "win32";
let root = "";
afterEach(() => root && rmSync(root, { recursive: true, force: true }));

describe.runIf(posix)("файл лога — только владельцу", () => {
  it("новый каталог 0700, новый файл 0600", () => {
    root = mkdtempSync(join(tmpdir(), "log-mode-"));
    const path = join(root, "1c-odata-mcp", "server.log");
    const d = logFile(path);
    d.write("x\n");
    expect(statSync(join(root, "1c-odata-mcp")).mode & 0o777).toBe(0o700);
    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  it("уже существующие каталог 0755 и файл 0644 — права сужаются", () => {
    root = mkdtempSync(join(tmpdir(), "log-mode-"));
    const dir = join(root, "1c-odata-mcp");
    mkdirSync(dir, { mode: 0o755 });
    chmodSync(dir, 0o755);
    const path = join(dir, "server.log");
    writeFileSync(path, "old\n");
    chmodSync(path, 0o644);
    logFile(path);
    expect(statSync(dir).mode & 0o777).toBe(0o700);
    expect(statSync(path).mode & 0o777).toBe(0o600);
  });
});
