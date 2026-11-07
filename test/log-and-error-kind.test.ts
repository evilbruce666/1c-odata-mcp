import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import { logger } from "../src/logger.js";
import { ERROR_KIND_META, fail, guard } from "../src/tools/_shared.js";
import { InputError } from "../src/errors.js";
import { fake1C } from "./helpers/fake-1c.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("журнал не хранит параметры запроса на обычных уровнях", () => {
  it("warn при ошибке и повторе — без $filter", async () => {
    const f = fake1C({ READ_ONLY: "true", ODATA_RETRIES: "1" });
    const warn = vi.spyOn(logger, "warn").mockImplementation(() => undefined);
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response("{}", { status: 503 })),
    );
    const path = "Catalog_Контрагенты?$filter=ИНН eq '7701234567'";
    await expect(f.conn.client.getCollection(path)).rejects.toBeInstanceOf(ODataError);
    expect(warn).toHaveBeenCalled();
    const logged = JSON.stringify(warn.mock.calls);
    expect(logged).not.toContain("7701234567");
    expect(logged).not.toContain("filter");
    expect(logged).toContain("odata error");
  }, 15_000);

  it("в исходниках клиента warn не получает полный url", () => {
    const src = readFileSync(new URL("../src/odata/client.ts", import.meta.url), "utf8");
    expect(src).not.toMatch(/logger\.(warn|info|error)\(\{ url[,: ](?!\s*withoutQuery)/);
  });
});

describe("вид ошибки в _meta", () => {
  const kindOf = (r: { _meta?: Record<string, unknown> }) => r._meta?.[ERROR_KIND_META];
  it.each([
    [new ODataError({ kind: "auth", message: "нет" }), "auth"],
    [new InputError("плохо"), "input"],
    [new Error("сбой"), "internal"],
  ] as const)("%s → %s", async (error, kind) => {
    const r = await guard("t", () => Promise.reject(error));
    expect(r.isError).toBe(true);
    expect(kindOf(r)).toBe(kind);
  });
  it("отказ инструмента по умолчанию — refused, текст не меняется", () => {
    const r = fail("Сначала выполните dry-run");
    expect(kindOf(r)).toBe("refused");
    expect(r.content).toEqual([{ type: "text", text: "Сначала выполните dry-run" }]);
  });
});
