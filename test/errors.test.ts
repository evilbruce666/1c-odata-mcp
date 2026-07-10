import { describe, it, expect } from "vitest";
import { fromHttpStatus } from "../src/odata/errors.js";

describe("fromHttpStatus — классификация HTTP-ошибок OData", () => {
  it("401 → auth, не повторяется", () => {
    const e = fromHttpStatus(401, "url");
    expect(e.kind).toBe("auth");
    expect(e.retryable).toBe(false);
  });
  it("404 → not_found", () => {
    expect(fromHttpStatus(404, "url").kind).toBe("not_found");
  });
  it("400 → bad_request", () => {
    expect(fromHttpStatus(400, "url").kind).toBe("bad_request");
  });
  it("429 → rate_limit, повторяется", () => {
    const e = fromHttpStatus(429, "url");
    expect(e.kind).toBe("rate_limit");
    expect(e.retryable).toBe(true);
  });
  it("500 → server, повторяется", () => {
    const e = fromHttpStatus(500, "url");
    expect(e.kind).toBe("server");
    expect(e.retryable).toBe(true);
  });
  it("вытаскивает сообщение 1С из JSON-тела ошибки", () => {
    const body = JSON.stringify({ error: { message: { value: "Плохой фильтр" } } });
    expect(fromHttpStatus(400, "url", body).message).toContain("Плохой фильтр");
  });
  it("вытаскивает сообщение из ключа odata.error (OData v3, реальный формат 1С)", () => {
    const body = JSON.stringify({
      "odata.error": {
        code: "-1",
        message: { lang: "ru", value: "Не удалось записать: Счет-фактура выданный!" },
      },
    });
    expect(fromHttpStatus(500, "url", body).message).toContain("Не удалось записать");
  });
});
