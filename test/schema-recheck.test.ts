import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { parseConfig } from "../src/config/env.js";
import { ServerContext } from "../src/context.js";
import {
  NotPublishedError,
  RECHECK_AFTER_MS,
  requireEntity,
  requirePublished,
} from "../src/odata/publication.js";
import { registerMetaTools } from "../src/tools/meta.js";
import { BP30_METADATA_XML, UT11_METADATA_XML } from "./helpers/metadata-snapshot.js";

const SET = "InformationRegister_НастройкиСистемыНалогообложения";
const without = BP30_METADATA_XML.replace(new RegExp(`<EntitySet Name="${SET}"[^>]*/>`), "");
const START = new Date("2030-01-01T10:00:00Z").getTime();

let xml: string;
let loads: number;
let failing: boolean;

function connect() {
  vi.stubGlobal("fetch", async () => {
    loads++;
    if (failing) throw new TypeError("fetch failed");
    return new Response(xml, { status: 200, headers: { "Content-Type": "application/xml" } });
  });
  return new ServerContext(
    parseConfig({
      ODATA_BASE_URL: "http://fake.test/odata/standard.odata/",
      ODATA_USERNAME: "u",
      ODATA_PASSWORD: "p",
      ODATA_RETRIES: "0",
    }),
  ).db();
}
const later = (ms: number) => vi.setSystemTime(Date.now() + ms);
const ask = (conn: ReturnType<typeof connect>) => requireEntity(conn, [SET], "Настройки");

beforeEach(() => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(START);
  xml = without;
  loads = 0;
  failing = false;
});
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("перечитывание схемы перед отказом «не опубликовано»", () => {
  it("объект добавили в «Состав» — следующий вызов через минуту его видит", async () => {
    expect(without).not.toBe(BP30_METADATA_XML);
    const conn = connect();
    await expect(ask(conn)).rejects.toThrow(NotPublishedError);
    expect(loads).toBe(1);
    xml = BP30_METADATA_XML;
    later(RECHECK_AFTER_MS + 1);
    await expect(ask(conn)).resolves.toBe(SET);
    expect(loads).toBe(2);
    // Найденное берётся из новой схемы без новых загрузок.
    await expect(ask(conn)).resolves.toBe(SET);
    expect(loads).toBe(2);
  });

  it("не чаще раза в минуту: подряд идущие отказы $metadata не тянут", async () => {
    const conn = connect();
    for (let i = 0; i < 3; i++) await expect(ask(conn)).rejects.toThrow(NotPublishedError);
    expect(loads).toBe(1);
    later(RECHECK_AFTER_MS + 1);
    await Promise.all([ask(conn), ask(conn)].map((p) => p.catch(() => undefined)));
    await expect(requirePublished(conn, "Catalog_НетТакого")).rejects.toThrow(NotPublishedError);
    expect(loads).toBe(2);
  });

  it("имя объекта от пользователя тоже перепроверяется", async () => {
    const conn = connect();
    await conn.getMetadata();
    xml = BP30_METADATA_XML;
    later(RECHECK_AFTER_MS + 1);
    await expect(requirePublished(conn, SET)).resolves.toBeUndefined();
  });

  it("describe_entity — им проверяют, появился ли объект, — тоже видит добавленный", async () => {
    const conn = connect();
    await conn.getMetadata();
    xml = BP30_METADATA_XML;
    later(RECHECK_AFTER_MS + 1);
    let handler!: (a: { entitySet: string }) => Promise<{ isError?: boolean }>;
    registerMetaTools(
      {
        registerTool: (name: string, _c: unknown, h: typeof handler) => {
          if (name === "read.schema.describe_entity") handler = h;
        },
      } as never,
      { db: () => conn, databases: () => [], defaultName: "x" } as never,
    );
    expect((await handler({ entitySet: SET })).isError).toBeFalsy();
    expect(loads).toBe(2);
  });

  it("сбой перечитывания — тот же отказ, прежняя схема цела, повтора сразу нет", async () => {
    const conn = connect();
    await conn.getMetadata();
    failing = true;
    later(RECHECK_AFTER_MS + 1);
    await expect(ask(conn)).rejects.toThrow(NotPublishedError);
    await expect(ask(conn)).rejects.toThrow(NotPublishedError);
    expect(loads).toBe(2);
    await expect(requireEntity(conn, ["Catalog_Организации"], "Организации")).resolves.toBe(
      "Catalog_Организации",
    );
    expect(loads).toBe(2);
  });

  it("в УТ отказ по бухучёту не перечитывает схему: объекта нет в самой конфигурации", async () => {
    xml = UT11_METADATA_XML;
    const conn = connect();
    await conn.getMetadata();
    later(RECHECK_AFTER_MS + 1);
    await expect(requireEntity(conn, ["AccountingRegister_Хозрасчетный"], "Проводки")).rejects.toThrow(
      NotPublishedError,
    );
    expect(loads).toBe(1);
  });
});
