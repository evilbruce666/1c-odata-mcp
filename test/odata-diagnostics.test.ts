import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { publicationHelp } from "../src/odata/publication.js";
import { readSavedSettlementSettings } from "../src/odata/saved-settlement-settings.js";
import { guard, ok } from "../src/tools/_shared.js";
import { fake1C } from "./helpers/fake-1c.js";
import { BP30_METADATA_XML, SNAPSHOT_REF } from "./helpers/metadata-snapshot.js";

const ENTITY = "Catalog_Контрагенты";
const WITHOUT_ENTITY = BP30_METADATA_XML.replace(/<EntitySet Name="Catalog_Контрагенты"[\s\S]*?\/>/, "");
const textOf = (result: CallToolResult) => JSON.stringify(result.content);
const response = (xml = BP30_METADATA_XML) => new Response(xml);
const denied = (status = 500) =>
  new Response(
    JSON.stringify({
      "odata.error": { message: { value: "Нарушение прав доступа! secret-record-data" } },
    }),
    { status },
  );
function tools(ctx: ReturnType<typeof fake1C>["ctx"]) {
  return (
    createServer(ctx) as unknown as {
      _registeredTools: Record<string, { handler: (args: object, extra: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
}
afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("точная диагностика OData", () => {
  it("сначала предлагает добавить в Состав и не выдаёт невидимость за установленную причину", () => {
    const message = publicationHelp([{ label: "Контрагенты", candidates: [ENTITY] }]);
    expect(message).toContain("отсутствуют в доступной OData-схеме");
    expect(message).toContain(ENTITY);
    expect(message.indexOf("Добавьте")).toBeLessThan(message.indexOf("другое имя"));
    expect(message).toContain("конфигурации");
    expect(message).toContain("ограничения прав");
  });
  it("ошибка настроек расчётов называет техническое имя регистра", async () => {
    const { conn } = fake1C();
    vi.spyOn(conn, "getMetadata").mockResolvedValue({ entities: new Map(), odataVersion: "3.0" });
    await expect(
      readSavedSettlementSettings(conn, {
        scope: "contract",
        organizationRef: SNAPSHOT_REF,
        counterpartyRef: SNAPSHOT_REF,
        contractRef: SNAPSHOT_REF,
        settlementKind: "РасчетыВВалютеРегламентированногоУчета",
      }),
    ).rejects.toThrow("InformationRegister_СчетаУчетаРасчетовСКонтрагентами");
  });
  it.each([
    ["GET", "", 401],
    ["GET", "", 403],
    ["GET", "", 500],
    ["POST", "", 500],
    ["PATCH", "", 500],
    ["POST", "/Post", 500],
    ["POST", "/Unpost", 500],
  ])("отказ %s %s (%s) сообщает только безопасную операцию и EntitySet", async (method, action, status) => {
    const { conn } = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true", ODATA_RETRIES: "2" });
    const fetchMock = vi.fn(async () => denied(status));
    vi.stubGlobal("fetch", fetchMock);
    const path = `${encodeURIComponent(ENTITY)}(guid'${SNAPSHOT_REF}')${action}?$filter=secret-filter`;
    const result = await guard("synthetic", async () =>
      ok(await conn.client.request(path, method, { secret: "payload" })),
    );
    const message = textOf(result);
    expect(result.isError).toBe(true);
    expect(message).toContain("Отказ получен при операции");
    expect(message).toContain(action ? `POST ${action.slice(1)}` : method);
    expect(message).toContain(ENTITY);
    for (const secret of [SNAPSHOT_REF, "secret", "payload", "https:", "$filter", "fake-1c.invalid"])
      expect(message).not.toContain(secret);
    expect(message).not.toContain("нужны права");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it("ограничивает контекст некорректного или чрезмерно длинного пути", async () => {
    const { conn } = fake1C();
    vi.stubGlobal("fetch", async () => denied(403));
    for (const path of [
      `${"X".repeat(1000)}?private=query`,
      "%FF/private-data",
      "Catalog_X%0Aprivate-data",
    ]) {
      const result = await guard("synthetic", async () => ok(await conn.client.request(path)));
      expect(textOf(result).length).toBeLessThan(600);
      expect(textOf(result)).not.toContain("private");
    }
  });
  it("отказ при чтении metadata тоже называет операцию без URL", async () => {
    const { ctx } = fake1C();
    vi.stubGlobal("fetch", async () => denied(403));
    const result = await tools(ctx)["read.system.health_check"]!.handler({}, {});
    expect(textOf(result)).toContain("GET");
    expect(textOf(result)).toContain("$metadata");
    expect(textOf(result)).not.toContain("https:");
  });
  it("404 отсутствующей записи не предлагает менять Состав", async () => {
    const { conn } = fake1C();
    vi.stubGlobal(
      "fetch",
      async () =>
        new Response(JSON.stringify({ "odata.error": { message: "Экземпляр сущности не найден" } }), {
          status: 404,
        }),
    );
    const result = await guard("synthetic", async () =>
      ok(await conn.client.getEntity(`${ENTITY}(guid'${SNAPSHOT_REF}')`)),
    );
    expect(textOf(result)).toContain("записи с таким ключом нет");
    expect(textOf(result)).not.toContain("Состав");
  });
});

describe("health_check перечитывает общую схему", () => {
  it("делает один свежий запрос на вызов и обновляет schema/tools cache", async () => {
    const { ctx, conn } = fake1C();
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(response(WITHOUT_ENTITY))
      .mockResolvedValueOnce(response())
      .mockResolvedValueOnce(response(WITHOUT_ENTITY));
    vi.stubGlobal("fetch", fetchMock);
    expect((await conn.available()).has(ENTITY)).toBe(false);
    const registered = tools(ctx);
    expect((await registered["read.system.health_check"]!.handler({}, {})).isError).toBeUndefined();
    expect((await conn.available()).has(ENTITY)).toBe(true);
    const listed = await registered["read.schema.list_entities"]!.handler({ search: ENTITY }, {});
    expect(textOf(listed)).toContain(ENTITY);
    expect(
      (await registered["read.schema.describe_entity"]!.handler({ entitySet: ENTITY }, {})).isError,
    ).toBeUndefined();
    expect(fetchMock).toHaveBeenCalledTimes(2);
    await registered["read.system.health_check"]!.handler({}, {});
    expect((await conn.available()).has(ENTITY)).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });
  it("сбой обновления не превращается в ok по старому кешу; следующее чтение восстанавливается", async () => {
    const { ctx, conn } = fake1C();
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(response())
      .mockResolvedValueOnce(denied(403))
      .mockResolvedValueOnce(response(WITHOUT_ENTITY));
    vi.stubGlobal("fetch", fetchMock);
    await conn.getMetadata();
    const result = await tools(ctx)["read.system.health_check"]!.handler({}, {});
    expect(result.isError).toBe(true);
    expect((await conn.available()).has(ENTITY)).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });
  it("старый незавершённый запрос не сбрасывает новый успешный кеш при отказе", async () => {
    const { ctx, conn } = fake1C();
    let rejectOld!: (error: Error) => void;
    const old = new Promise<Response>((_resolve, reject) => {
      rejectOld = reject;
    });
    const fetchMock = vi.fn().mockReturnValueOnce(old).mockResolvedValueOnce(response());
    vi.stubGlobal("fetch", fetchMock);
    const first = conn.getMetadata().catch(() => undefined);
    const result = await tools(ctx)["read.system.health_check"]!.handler({}, {});
    expect(result.isError).toBeUndefined();
    rejectOld(new Error("old request failed"));
    await first;
    expect((await conn.available()).has(ENTITY)).toBe(true);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
});

describe("детерминированный отказ 1С не повторяется (#53)", () => {
  it("GET с 500 «Операция не разрешена в предложении ГДЕ» отправляется один раз", async () => {
    const { conn } = fake1C({ ODATA_RETRIES: "2" });
    const fetchMock = vi.fn(
      async () =>
        new Response(
          JSON.stringify({
            "odata.error": { code: "-1", message: { value: 'Операция не разрешена в предложении "ГДЕ"' } },
          }),
          { status: 500 },
        ),
    );
    vi.stubGlobal("fetch", fetchMock);
    const result = await guard("synthetic", async () =>
      ok(await conn.client.getCollection(`${ENTITY}?$format=json`)),
    );
    expect(result.isError).toBe(true);
    expect(textOf(result)).toContain("ГДЕ");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});
