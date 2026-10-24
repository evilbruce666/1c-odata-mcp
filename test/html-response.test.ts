import { afterEach, describe, expect, it, vi } from "vitest";
import { ODataClient } from "../src/odata/client.js";
import { ODataError } from "../src/odata/errors.js";

// Публикация 1С (1cfresh) на отказ доступа отвечает HTML-страницей с кодом 200: это не сетевой
// сбой — без повторов и с понятным текстом вместо «Unexpected token '<'».
const PAGE =
  "<!DOCTYPE html><html><head><title>Доступ запрещён</title><style>p{}</style></head>" +
  "<body><p>У вас нет прав</p></body></html>";

function client(writable = false, retries = 2) {
  return new ODataClient(
    { name: "t", baseUrl: "http://localhost/odata/", username: "u", password: "p", writable } as never,
    { readOnly: !writable, writeJournalDir: "/nonexistent", retries, timeoutMs: 1000 } as never,
  );
}

function stub(body: string, contentType: string) {
  const fetch = vi.fn(
    async () => new Response(body, { status: 200, headers: { "Content-Type": contentType } }),
  );
  vi.stubGlobal("fetch", fetch);
  return fetch;
}

afterEach(() => vi.unstubAllGlobals());

describe("ответ 200 не в формате OData", () => {
  it("HTML-страница — ошибка parse с заголовком страницы, без повторов", async () => {
    const fetch = stub(PAGE, "text/html; charset=utf-8");
    const e = (await client()
      .getCollection("Catalog_X")
      .catch((x: unknown) => x)) as ODataError;
    expect(e).toBeInstanceOf(ODataError);
    expect(e.kind).toBe("parse");
    expect(e.retryable).toBe(false);
    expect(e.message).toContain("HTML-страницу");
    expect(e.message).toContain("Доступ запрещён");
    expect(e.message).not.toContain("<");
    expect(fetch).toHaveBeenCalledOnce();
  });

  it("HTML без заголовка Content-Type распознаётся по разметке", async () => {
    stub("<html><body>Ошибка публикации</body></html>", "application/octet-stream");
    const e = (await client()
      .getEntity("Catalog_X")
      .catch((x: unknown) => x)) as ODataError;
    expect(e.kind).toBe("parse");
    expect(e.message).toContain("Ошибка публикации");
  });

  it("при записи — предупреждение, что исход неизвестен", async () => {
    stub(PAGE, "text/html");
    const e = (await client(true)
      .action("Document_X(guid'1')/Post")
      .catch((x: unknown) => x)) as ODataError;
    expect(e.kind).toBe("parse");
    expect(e.message).toContain("исход неизвестен");
  });

  it("битый JSON — parse, а не повторяемая сетевая ошибка", async () => {
    const fetch = stub("{не json", "application/json");
    const e = (await client()
      .getCollection("Catalog_X")
      .catch((x: unknown) => x)) as ODataError;
    expect(e.kind).toBe("parse");
    expect(e.retryable).toBe(false);
    expect(fetch).toHaveBeenCalledOnce();
  });

  it("$metadata HTML-страницей — та же понятная ошибка", async () => {
    stub(PAGE, "text/html");
    const e = (await client()
      .getText("$metadata")
      .catch((x: unknown) => x)) as ODataError;
    expect(e.kind).toBe("parse");
    expect(e.message).toContain("Доступ запрещён");
  });

  it("обычный JSON и пустое тело действия — как раньше", async () => {
    stub('{"value":[]}', "application/json");
    expect(await client().getCollection("Catalog_X")).toEqual({ value: [] });
    stub("", "application/json");
    expect(await client(true).action("Document_X(guid'1')/Post")).toBeUndefined();
  });
});
