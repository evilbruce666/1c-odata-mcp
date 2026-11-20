import { describe, expect, it } from "vitest";
import { parseConfig, writableHint } from "../src/config/env.js";

/**
 * Подсказка «как включить запись» называет переменную, которую читает код (аудит документации: для одной базы
 * советовали ODATA_DB_DEFAULT_WRITABLE, а код читает ODATA_WRITABLE — пользователь упирался в тупик).
 */
const creds = { ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };

describe("подсказка включения записи", () => {
  it("одна база (ODATA_BASE_URL) — ODATA_WRITABLE, и именно он включает запись", () => {
    const off = parseConfig({ ODATA_BASE_URL: "https://x.invalid/odata/standard.odata/", ...creds });
    expect(writableHint(off.connections[0]!)).toBe("ODATA_WRITABLE=true");
    const on = parseConfig({
      ODATA_BASE_URL: "https://x.invalid/odata/standard.odata/",
      ODATA_WRITABLE: "true",
      ...creds,
    });
    expect(on.connections[0]!.writable).toBe(true);
    const wrong = parseConfig({
      ODATA_BASE_URL: "https://x.invalid/odata/standard.odata/",
      ODATA_DB_DEFAULT_WRITABLE: "true",
      ...creds,
    });
    expect(wrong.connections[0]!.writable).toBe(false);
  });

  it("несколько баз — ODATA_DB_<ИМЯ>_WRITABLE", () => {
    const cfg = parseConfig({
      ODATA_DB_BUH_BASE_URL: "https://x.invalid/buh/odata/standard.odata/",
      ODATA_DB_BUH_USERNAME: "u",
      ODATA_DB_BUH_PASSWORD: "p",
    });
    expect(writableHint(cfg.connections[0]!)).toBe("ODATA_DB_BUH_WRITABLE=true");
  });
});
