import { describe, it, expect } from "vitest";
import { parseConfig } from "../src/config/env.js";

describe("parseConfig — одна база (обратная совместимость)", () => {
  it("ODATA_BASE_URL → база default, дефолты, read-only", () => {
    const c = parseConfig({
      ODATA_BASE_URL: "https://host/db/odata/standard.odata/",
      ODATA_USERNAME: "u",
      ODATA_PASSWORD: "p",
    });
    expect(c.connections).toHaveLength(1);
    expect(c.connections[0]!.name).toBe("default");
    expect(c.connections[0]!.writable).toBe(false);
    expect(c.defaultName).toBe("default");
    expect(c.behavior.readOnly).toBe(true);
    expect(c.behavior.pageSize).toBe(100);
  });
  it("добавляет завершающий слэш к URL", () => {
    const c = parseConfig({
      ODATA_BASE_URL: "https://host/db/odata/standard.odata",
      ODATA_USERNAME: "u",
      ODATA_PASSWORD: "p",
    });
    expect(c.connections[0]!.baseUrl.endsWith("/")).toBe(true);
  });
});

describe("parseConfig — несколько баз", () => {
  it("ODATA_DB_<ИМЯ>_* + per-base writable + default", () => {
    const c = parseConfig({
      ODATA_DB_IP_BASE_URL: "https://host/ip/odata/standard.odata/",
      ODATA_DB_IP_USERNAME: "u1",
      ODATA_DB_IP_PASSWORD: "p1",
      ODATA_DB_IP_WRITABLE: "true",
      ODATA_DB_OOO_BASE_URL: "https://host/ooo/odata/standard.odata/",
      ODATA_DB_OOO_USERNAME: "u2",
      ODATA_DB_OOO_PASSWORD: "p2",
      ODATA_DEFAULT_DB: "ooo",
      READ_ONLY: "false",
    });
    expect(c.connections.map((x) => x.name).sort()).toEqual(["ip", "ooo"]);
    expect(c.connections.find((x) => x.name === "ip")!.writable).toBe(true);
    expect(c.connections.find((x) => x.name === "ooo")!.writable).toBe(false);
    expect(c.defaultName).toBe("ooo");
    expect(c.behavior.readOnly).toBe(false);
  });
  it("без баз — бросает понятную ошибку", () => {
    expect(() => parseConfig({})).toThrow(/Некорректная конфигурация/);
  });
});
