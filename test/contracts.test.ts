import { describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { createServer } from "../src/mcp/server.js";
import { ODataError } from "../src/odata/errors.js";
import { withDefaultContract, withEnumValues } from "../src/tools/write.js";

const ACT = "Document_РеализацияТоваровУслуг";
const cp = "11111111-1111-1111-1111-111111111111";
const org = "22222222-2222-2222-2222-222222222222";

const connWith = (contracts: Array<{ Ref_Key: string; Description: string }>) => {
  const getCollection = vi.fn(async (_path: string) => ({ value: contracts }));
  return {
    getCollection,
    conn: {
      available: async () => new Set(["Catalog_ДоговорыКонтрагентов"]),
      client: { getCollection },
    } as never,
  };
};

describe("договор по умолчанию — как форма 1С при выборе контрагента", () => {
  const payload = { Контрагент_Key: cp, Организация_Key: org };

  it("ровно один договор нужного вида → подставляется, с пояснением", async () => {
    const { conn, getCollection } = connWith([{ Ref_Key: "ctr-1", Description: "Договор 1" }]);
    const r = await withDefaultContract(conn, ACT, payload);
    expect(r.payload["ДоговорКонтрагента_Key"]).toBe("ctr-1");
    expect(r.notes[0]).toContain("подставлен единственный договор с покупателем");
    const url = decodeURIComponent(getCollection.mock.calls[0]![0]);
    expect(url).toContain(`Owner_Key eq guid'${cp}'`);
    expect(url).toContain("ВидДоговора eq 'СПокупателем'");
    expect(url).toContain(`Организация_Key eq guid'${org}'`);
  });

  it("несколько договоров → не угадываем, предупреждаем", async () => {
    const { conn } = connWith([
      { Ref_Key: "a", Description: "A" },
      { Ref_Key: "b", Description: "B" },
    ]);
    const r = await withDefaultContract(conn, ACT, payload);
    expect(r.payload["ДоговорКонтрагента_Key"]).toBeUndefined();
    expect(r.notes[0]).toContain("несколько договоров");
  });

  it("ни одного → предупреждаем, что 1С не проведёт", async () => {
    const { conn } = connWith([]);
    const r = await withDefaultContract(conn, ACT, payload);
    expect(r.payload["ДоговорКонтрагента_Key"]).toBeUndefined();
    expect(r.notes[0]).toContain("нет договора с покупателем");
  });

  it("договор указан, документ закупки → вид СПоставщиком; денежные документы не трогаем", async () => {
    const { conn, getCollection } = connWith([{ Ref_Key: "x", Description: "X" }]);
    expect(
      (await withDefaultContract(conn, ACT, { ...payload, ДоговорКонтрагента_Key: "given" })).notes,
    ).toEqual([]);
    await withDefaultContract(conn, "Document_ПоступлениеТоваровУслуг", payload);
    expect(decodeURIComponent(getCollection.mock.calls.at(-1)![0])).toContain(
      "ВидДоговора eq 'СПоставщиком'",
    );
    const calls = getCollection.mock.calls.length;
    expect((await withDefaultContract(conn, "Document_ПриходныйКассовыйОрдер", payload)).notes).toEqual([]);
    expect(getCollection.mock.calls.length).toBe(calls);
  });
});

describe("post_document: без договора — понятная ошибка до вызова 1С", () => {
  const ref = "33333333-3333-4333-8333-333333333333";
  const tool = (contract: string) => {
    const action = vi.fn(async () => ({}));
    const connection = {
      cfg: { name: "default", writable: true },
      available: async () => new Set([ACT]),
      getMetadata: async () => ({
        entities: new Map([
          [ACT, { properties: [{ name: "Контрагент_Key" }, { name: "ДоговорКонтрагента_Key" }] }],
        ]),
      }),
      client: { action, getEntity: async () => ({ Контрагент_Key: cp, ДоговорКонтрагента_Key: contract }) },
    };
    const t = (
      createServer({ db: () => connection } as never) as unknown as {
        _registeredTools: Record<
          string,
          { handler: (a: Record<string, unknown>, e: unknown) => Promise<CallToolResult> }
        >;
      }
    )._registeredTools["write.document.post_document"]!;
    return {
      action,
      run: () => t.handler({ database: "default", entitySet: ACT, ref, post: true, confirm: true }, {}),
    };
  };

  it("договор пуст → ошибка с подсказкой, Post не вызывается", async () => {
    const { action, run } = tool("00000000-0000-0000-0000-000000000000");
    const res = await run();
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toContain("не заполнен договор");
    expect(action).not.toHaveBeenCalled();
  });

  it("договор заполнен → проводим", async () => {
    const { action, run } = tool("44444444-4444-4444-8444-444444444444");
    const res = await run();
    expect(res.isError).toBeFalsy();
    expect(action).toHaveBeenCalledOnce();
  });
});

describe("ошибка значения перечисления — с допустимыми значениями", () => {
  it("дописывает члены перечисления из $metadata", async () => {
    const conn = {
      client: {
        getText: async () =>
          '<EnumType Name="ВидыОперацийПКО"><Member Name="ОплатаПокупателя"/><Member Name="ПрочийПриход"/></EnumType>',
      },
    } as never;
    const e = new ODataError({
      kind: "bad_request",
      message: "Перечисление 'ВидыОперацийПКО' не содержит элемент 'ПоступлениеОплатыОтПокупателя'.",
    });
    const r = (await withEnumValues(conn, e)) as ODataError;
    expect(r.message).toContain("Допустимые значения ВидыОперацийПКО: ОплатаПокупателя, ПрочийПриход.");
    expect(r.kind).toBe("bad_request");
    expect(await withEnumValues(conn, new Error("другое"))).toBeInstanceOf(Error);
  });
});
