import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { createServer } from "../src/mcp/server.js";
import { ODataError } from "../src/odata/errors.js";
import { parentMoveRefusal } from "../src/odata/parent-move.js";
import { fake1C } from "./helpers/fake-1c.js";

// Общие write.entity.* обходили проверки частных инструментов (аудит 26.11): move_to_folder и update_entity могли
// переместить группу в саму себя или в свою подгруппу и в обычный элемент или помеченную группу, а update_entity —
// записать ИНН, который уже есть у другого контрагента.
const CP = "Catalog_Контрагенты";
const EMPTY = "00000000-0000-0000-0000-000000000000";
const g = (n: number) => `cccccccc-0000-4000-8000-${String(n).padStart(12, "0")}`;
const [A, AA, AAA, B, ITEM, MARKED, OTHER, MISSING] = [1, 2, 3, 4, 5, 6, 7, 99].map(g) as [
  string,
  string,
  string,
  string,
  string,
  string,
  string,
  string,
];
type Row = Record<string, unknown>;
const folder = (Parent_Key: string, extra: Row = {}): Row => ({
  Parent_Key,
  IsFolder: true,
  DeletionMark: false,
  ...extra,
});
const element = (Parent_Key: string, extra: Row = {}): Row => ({
  ...folder(Parent_Key, extra),
  IsFolder: false,
});

afterEach(() => vi.restoreAllMocks());

function setup(schema: "bp30" | "ut11" = "bp30", rows?: Record<string, Row>) {
  const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" }, { schema });
  const db: Record<string, Row> = rows ?? {
    [A]: folder(EMPTY),
    [AA]: folder(A),
    [AAA]: folder(AA),
    [B]: folder(EMPTY),
    [ITEM]: element(A, { ИНН: "7700000001", КПП: "" }),
    [OTHER]: element(EMPTY, { ИНН: "7700000002", КПП: "", Description: "Другой" }),
    [MARKED]: folder(EMPTY, { DeletionMark: true }),
  };
  const selects: string[] = [];
  vi.spyOn(f.conn.client, "getEntity").mockImplementation(async (path: string) => {
    const ref = /guid'([0-9a-f-]{36})'/i.exec(decodeURIComponent(path))![1]!.toLowerCase();
    selects.push(decodeURIComponent(path));
    if (!db[ref]) throw new ODataError({ kind: "not_found", message: "нет объекта" });
    return { Ref_Key: ref, ...db[ref] } as never;
  });
  vi.spyOn(f.conn.client, "getCollection").mockImplementation(async (path: string) => {
    const inn = /ИНН eq '([^']*)'/.exec(decodeURIComponent(path).replace(/\+/g, " "))?.[1];
    const kpps = [
      ...decodeURIComponent(path)
        .replace(/\+/g, " ")
        .matchAll(/КПП eq '([^']*)'/g),
    ].map((m) => m[1]!);
    return {
      value: Object.entries(db)
        .filter(([, r]) => inn !== undefined && r["ИНН"] === inn && r["DeletionMark"] !== true)
        // Отбор по КПП делает 1С — подделка повторяет его.
        .filter(([, r]) => !kpps.length || kpps.includes(String(r["КПП"] ?? "")))
        .map(([Ref_Key, r]) => ({ Ref_Key, ...r })),
    } as never;
  });
  const patch = vi.spyOn(f.conn.client, "patch").mockImplementation(async (path: string, body: unknown) => {
    const ref = /guid'([0-9a-f-]{36})'/i.exec(path)![1]!.toLowerCase();
    Object.assign((db[ref] ??= {}), body as Row);
    return { Ref_Key: ref, ...db[ref] } as never;
  });
  const tools = (
    createServer(f.ctx) as unknown as {
      _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
    }
  )._registeredTools;
  const call = async (name: string, args: object) => {
    const r = await tools[name]!.handler(args, {});
    return { error: !!r.isError, text: (r.content[0] as { text: string }).text };
  };
  /** Два пути одного перемещения: частный инструмент и общий update_entity. */
  const movers = [
    (ref: string, to: string, confirm: boolean, entitySet = CP) =>
      call("write.entity.move_to_folder", { entitySet, ref, folder: to, confirm }),
    (ref: string, to: string, confirm: boolean, entitySet = CP) =>
      call("write.entity.update_entity", { entitySet, ref, fields: { Parent_Key: to }, confirm }),
  ];
  return { call, movers, patch, selects, conn: f.conn };
}

describe("перемещение в иерархии: move_to_folder и update_entity с Parent_Key", () => {
  it.each([
    ["группу в саму себя", A, A, /в себя/],
    ["группу в свою подгруппу", A, AA, /потомк|подгрупп/],
    ["группу в подгруппу второго уровня", A, AAA, /потомк|подгрупп/],
    ["в обычный элемент, а не в группу", B, ITEM, /не группа|не папка/],
    ["в помеченную на удаление группу", ITEM, MARKED, /помечен/],
    ["в несуществующую группу", ITEM, MISSING, /не найден/],
  ])("%s — отказ в предпросмотре и при подтверждении, без записи", async (_, ref, to, why) => {
    const t = setup();
    for (const move of t.movers)
      for (const confirm of [false, true]) {
        const r = await move(ref, to, confirm);
        expect(r.error, `confirm=${confirm}: ${r.text}`).toBe(true);
        expect(r.text).toMatch(why);
      }
    expect(t.patch).not.toHaveBeenCalled();
  });

  it.each([
    ["элемент в группу", ITEM, B],
    ["группу в другую группу", AA, B],
    ["группу с подгруппами в другую ветку", A, B],
    ["в корень", AA, EMPTY],
  ])("%s — по-прежнему перемещается", async (_, ref, to) => {
    for (const i of [0, 1]) {
      const t = setup();
      const preview = await t.movers[i]!(ref, to, false);
      expect(preview.error, preview.text).toBe(false);
      const done = await t.movers[i]!(ref, to, true);
      expect(done.error, done.text).toBe(false);
      expect(t.patch).toHaveBeenCalledTimes(1);
      // Описание: заменяется только родитель — состав переносимой группы и прочие реквизиты не трогаются.
      expect(t.patch.mock.calls[0]![1]).toEqual({ Parent_Key: to });
    }
  });

  // Ревью: update_entity с прежним родителем в помеченной группе рядом с другими полями отказывал — это не
  // перемещение. Тесты брали только настоящие перемещения.
  it("родитель не меняется, хотя группа помечена на удаление, — проходит", async () => {
    for (const i of [0, 1]) {
      const t = setup("bp30", { [ITEM]: element(MARKED) });
      for (const confirm of [false, true]) {
        const r = await t.movers[i]!(ITEM, MARKED, confirm);
        expect(r.error, r.text).toBe(false);
      }
      expect(t.patch).toHaveBeenCalledTimes(1);
    }
  });

  it("цикл, уже лежащий в данных, не зацикливает проверку: отказ", async () => {
    const t = setup("bp30", { [A]: folder(B), [B]: folder(A), [AA]: folder(EMPTY) });
    const r = await t.movers[0]!(AA, A, true);
    expect(r.error).toBe(true);
    expect(t.patch).not.toHaveBeenCalled();
  });

  // У партнёров УТ нет IsFolder: родитель — обычный элемент, а запрос поля IsFolder 1С отвергает (400). Общую
  // запись в партнёров УТ защита не пускает, поэтому проверка вызывается напрямую на схеме УТ.
  it("справочник без групп: родителем может быть элемент, IsFolder не запрашивается, цикл — отказ", async () => {
    const plain = (Parent_Key: string): Row => ({ Parent_Key, DeletionMark: false });
    const t = setup("ut11", { [A]: plain(EMPTY), [AA]: plain(A), [B]: plain(EMPTY) });
    const set = "Catalog_Партнеры";
    expect(await parentMoveRefusal(t.conn, set, B, A)).toBeUndefined();
    expect(await parentMoveRefusal(t.conn, set, A, AA)).toMatch(/потомк/);
    expect(t.selects.length).toBeGreaterThan(0);
    expect(t.selects.filter((s) => s.includes("IsFolder"))).toEqual([]);
  });
});

describe("update_entity: ИНН, который уже есть у другого контрагента", () => {
  const upd = (t: ReturnType<typeof setup>, fields: object, confirm: boolean, entitySet = CP, ref = ITEM) =>
    t.call("write.entity.update_entity", { entitySet, ref, fields, confirm });

  it("чужой ИНН — отказ в предпросмотре и при подтверждении, без записи", async () => {
    const t = setup();
    for (const confirm of [false, true]) {
      const r = await upd(t, { ИНН: "7700000002" }, confirm);
      expect(r.error, r.text).toBe(true);
      expect(r.text).toContain("уже есть");
    }
    expect(t.patch).not.toHaveBeenCalled();
  });

  it.each([
    ["свой прежний ИНН вместе с другим полем", { ИНН: "7700000001", Комментарий: "x" }],
    ["очистка ИНН (null)", { ИНН: null }],
    ["свободный ИНН числом", { ИНН: 7700000003 }],
    ["поля без ИНН и КПП", { Комментарий: "x" }],
  ])("%s — проходит", async (_, fields) => {
    const t = setup();
    const r = await upd(t, fields, false);
    expect(r.error, r.text).toBe(false);
  });

  it("не контрагенты — проверка не вмешивается", async () => {
    const t = setup();
    const r = await upd(t, { ИНН: "7700000002" }, false, "Catalog_Организации");
    expect(r.error, r.text).toBe(false);
  });

  it("обособленное подразделение (другой КПП) — проходит с предупреждением", async () => {
    const t = setup("bp30", {
      [ITEM]: element(EMPTY, { ИНН: "", КПП: "" }),
      [OTHER]: element(EMPTY, { ИНН: "7700000002", КПП: "770101001", Description: "Другой" }),
    });
    const r = await upd(t, { ИНН: "7700000002", КПП: "780101001" }, false);
    expect(r.error, r.text).toBe(false);
    expect(r.text).toContain("другим КПП");
  });
});
