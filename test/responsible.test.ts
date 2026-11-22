import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { createCore } from "../src/tools/write.js";
import { withResponsible } from "../src/odata/responsible-context.js";
import { withWriteOperation } from "../src/odata/write-operation-context.js";
import { randomUUID } from "node:crypto";
import { fake1C, type FakeSchema } from "./helpers/fake-1c.js";
import { SNAPSHOT_REF as id } from "./helpers/metadata-snapshot.js";

/**
 * responsibleRef у всех инструментов создания документов: поле по схеме документа («Ответственный», иначе
 * «Менеджер»), «Автор» не трогается; документ без обоих — отказ; пользователь проверяется до сборки документа.
 */
const U = (n: number) => `cccccccc-0000-4000-8000-${String(n).padStart(12, "0")}`;
const USERS = [
  { Ref_Key: U(1), Description: "Иванов", DeletionMark: false, Недействителен: false, Служебный: false },
  { Ref_Key: U(2), Description: "Служебный", DeletionMark: false, Недействителен: false, Служебный: true },
];

function fake(schema: FakeSchema) {
  return fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-resp-")),
    },
    { schema, data: (q) => (q.base === "Catalog_Пользователи" ? USERS : undefined) },
  );
}

async function client(schema: FakeSchema) {
  const f = fake(schema);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(f.ctx).connect(a);
  const c = new Client({ name: "t", version: "1" });
  await c.connect(b);
  const call = async (name: string, args: object) => {
    const r = (await c.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((x) => x.text).join("") };
  };
  return { call, fake: f, c };
}

const op = <T>(fn: () => Promise<T>) => withWriteOperation(randomUUID(), "test", fn);

describe("createCore: поле ответственного по схеме документа (УТ 11.5)", () => {
  it.each([
    ["Document_РеализацияТоваровУслуг", "Менеджер_Key"],
    ["Document_ЗаказКлиента", "Менеджер_Key"],
    ["Document_ПеремещениеТоваров", "Ответственный_Key"],
    ["Document_СписаниеБезналичныхДенежныхСредств", "Ответственный_Key"],
  ])("%s → %s, «Автор» не заполняется", async (set, field) => {
    const f = fake("ut11");
    const r = await withResponsible(U(1), () => op(() => createCore(f.conn, set, { Posted: false }, false)));
    const payload = r.structuredContent?.["payload"] as Record<string, unknown>;
    expect(payload[field]).toBe(U(1));
    expect(payload).not.toHaveProperty("Автор_Key");
  });

  it("ПКО УТ (есть только «Автор») — отказ, ничего не записано", async () => {
    const f = fake("ut11");
    await expect(
      withResponsible(U(1), () =>
        createCore(f.conn, "Document_ПриходныйКассовыйОрдер", { Posted: false }, false),
      ),
    ).rejects.toThrow(/нет реквизита «Ответственный» или «Менеджер»/);
    expect(f.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });

  it("без responsibleRef и для справочника — поле не трогается", async () => {
    const f = fake("ut11");
    const plain = await op(() =>
      createCore(f.conn, "Document_РеализацияТоваровУслуг", { Posted: false }, false),
    );
    expect(plain.structuredContent?.["payload"]).not.toHaveProperty("Менеджер_Key");
    const cat = await withResponsible(U(1), () =>
      op(() => createCore(f.conn, "Catalog_Контрагенты", { Description: "x" }, false)),
    );
    expect(cat.structuredContent?.["payload"]).not.toHaveProperty("Ответственный_Key");
  });
});

describe("responsibleRef через сервер (Бухгалтерия 3.0)", () => {
  const invoice = { counterpartyRef: id, lines: [{ nomenclatureRef: id, quantity: 1, price: 1 }] };

  it("счёт: предпросмотр с ответственным в нижнем регистре; повтор подтверждения без дубля", async () => {
    const { call, fake: f } = await client("bp30");
    const p = await call("write.sales.create_invoice", {
      ...invoice,
      responsibleRef: `{${U(1).toUpperCase()}}`,
    });
    expect(p.isError, p.text).toBeFalsy();
    expect((p.structuredContent["payload"] as Record<string, unknown>)["Ответственный_Key"]).toBe(U(1));
    const args = {
      ...invoice,
      responsibleRef: `{${U(1).toUpperCase()}}`,
      confirm: true,
      operationId: p.structuredContent["operationId"],
    };
    const c1 = await call("write.sales.create_invoice", args);
    const c2 = await call("write.sales.create_invoice", args);
    expect(c1.isError, c1.text).toBeFalsy();
    expect(c2.structuredContent["replayed"]).toBe(true);
    const posts = f.requests.filter((q) => q.method === "POST");
    expect(posts).toHaveLength(1);
  });

  it("служебный или несуществующий — отказ до сборки, без записи", async () => {
    const { call, fake: f } = await client("bp30");
    for (const [ref, text] of [
      [U(2), "служебный"],
      [U(99), "не найден"],
    ]) {
      const r = await call("write.sales.create_invoice", { ...invoice, responsibleRef: ref });
      expect(r.isError).toBe(true);
      expect(r.text).toContain(text);
    }
    expect(f.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });

  it("копия документа — ответственный заменяется на указанного", async () => {
    const { call } = await client("bp30");
    const r = await call("write.document.copy_document", {
      entitySet: "Document_РеализацияТоваровУслуг",
      ref: id,
      responsibleRef: U(1),
    });
    expect(r.isError, r.text).toBeFalsy();
    expect((r.structuredContent["payload"] as Record<string, unknown>)["Ответственный_Key"]).toBe(U(1));
  });

  it("у инструментов справочников responsibleRef нет; у документов — есть", async () => {
    const { c } = await client("bp30");
    const tools = new Map((await c.listTools()).tools.map((t) => [t.name, t]));
    const has = (n: string) => "responsibleRef" in ((tools.get(n)?.inputSchema.properties ?? {}) as object);
    for (const n of [
      "write.counterparty.create_counterparty",
      "write.catalog.create_contract",
      "write.entity.update_entity",
    ])
      expect(has(n), n).toBe(false);
    for (const n of [
      "write.sales.create_invoice",
      "write.purchase.create_purchase",
      "write.warehouse.create_transfer",
      "write.money.create_cash_receipt",
      "write.document.copy_document",
      "write.pricing.create_price_setting",
      "write.money.create_payout_order",
    ])
      expect(has(n), n).toBe(true);
  });
});
