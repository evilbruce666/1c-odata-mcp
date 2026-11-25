import { describe, it, expect } from "vitest";
import { parseConfig, unencryptedDatabases } from "../src/config/env.js";

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

  it("принимает отдельный локальный каталог журнала подтверждённых записей", () => {
    const c = parseConfig({
      ODATA_BASE_URL: "https://host/db/odata/standard.odata/",
      ODATA_USERNAME: "u",
      ODATA_PASSWORD: "p",
      ODATA_WRITE_JOURNAL_DIR: "/var/tmp/1c-write-journal",
    });
    expect(c.behavior.writeJournalDir).toBe("/var/tmp/1c-write-journal");
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

describe("адрес 1С без шифрования", () => {
  const names = (urls: Record<string, string>) => {
    const env: Record<string, string> = {};
    for (const [name, url] of Object.entries(urls)) {
      env[`ODATA_DB_${name}_BASE_URL`] = url;
      env[`ODATA_DB_${name}_USERNAME`] = "u";
      env[`ODATA_DB_${name}_PASSWORD`] = "p";
    }
    return unencryptedDatabases(parseConfig(env).connections);
  };
  it("http не на этой машине называется, https и loopback — нет", () => {
    expect(
      names({
        OFFICE: "http://192.168.1.5/buh/odata/standard.odata/",
        CLOUD: "https://1c.example/odata/standard.odata/",
        LOCAL: "http://localhost/odata/",
        LOOP: "http://127.0.0.1:8080/odata/",
        V6: "http://[::1]/odata/",
      }),
    ).toEqual(["office"]);
  });
});

describe("ошибки конфигурации понятны и не раскрывают пароль", () => {
  const base = { ODATA_USERNAME: "u", ODATA_PASSWORD: "p" };
  const message = (env: Record<string, string>) => {
    try {
      parseConfig(env);
    } catch (e) {
      return (e as Error).message;
    }
    return "";
  };

  // Логин и пароль в самом адресе 1С fetch не принимает: ошибка с адресом целиком уходила модели и в лог (аудит 25.11).
  it.each([
    "https://user:secret-pw@host/db/odata/standard.odata/",
    "https://user@host/db/odata/standard.odata/",
  ])("адрес с логином или паролем — отказ при старте без адреса в тексте: %s", (url) => {
    const m = message({ ...base, ODATA_BASE_URL: url });
    expect(m).toMatch(/без логина и пароля/);
    expect(m).toContain("ODATA_USERNAME");
    expect(m).not.toContain("secret-pw");
    expect(m).not.toContain("user@");
  });

  it("адрес с логином у одной из нескольких баз — отказ с именем базы", () => {
    const m = message({
      ODATA_DB_IP_BASE_URL: "https://a:b@host/ip/odata/standard.odata/",
      ODATA_DB_IP_USERNAME: "u",
      ODATA_DB_IP_PASSWORD: "p",
    });
    expect(m).toMatch(/база "ip".*без логина и пароля/);
  });

  // Без пароля zod писал «Invalid input: expected string, received undefined».
  it.each([
    ["ODATA_PASSWORD", { ODATA_BASE_URL: "https://host/db/odata/standard.odata/", ODATA_USERNAME: "u" }],
    ["ODATA_USERNAME", { ODATA_BASE_URL: "https://host/db/odata/standard.odata/", ODATA_PASSWORD: "p" }],
  ])("не задан %s — по-русски, с именем переменной", (name, env) => {
    const m = message(env);
    expect(m).toContain(name);
    expect(m).toContain("не задан");
    expect(m).not.toMatch(/Invalid input|expected string/);
  });

  it("сообщение не отсылает к .env: конфиг часто задан в блоке env клиента", () => {
    expect(message({})).not.toContain("(.env)");
  });
});

describe("некорректный адрес базы — прежнее понятное сообщение и полный список ошибок", () => {
  // Проверка логина в адресе разбирала URL и на невалидной строке: parseConfig падал с голым «Invalid URL» (ревью 25.11).
  it.each(["not a url", "server/base/odata/standard.odata", ""])("адрес %j", (url) => {
    let m = "";
    try {
      parseConfig({
        ODATA_BASE_URL: "https://h/a/odata/standard.odata/",
        ODATA_DB_B_BASE_URL: url,
        ODATA_DB_B_USERNAME: "u",
        READ_ONLY: "maybe",
      });
    } catch (e) {
      m = (e as Error).message;
    }
    expect(m).toMatch(/Некорректная конфигурация/);
    expect(m).toMatch(/база "b": baseUrl должен быть валидным URL/);
    expect(m).toContain("не задан ODATA_DB_B_PASSWORD");
    expect(m).toContain("READ_ONLY");
  });
});
