import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery } from "./helpers/fake-1c.js";

/**
 * Контрагент с тем же ИНН уже есть — отказ со ссылкой (решение владельца 20.11.2026, Бухгалтерия и УТ). Обособленное
 * подразделение (свой КПП) — можно, с предупреждением. Повтор подтверждения после таймаута не считает дублем объект
 * своей же операции.
 */

const INN = "7700000001";
const A = "aaaaaaaa-0000-0000-0000-000000000001";

function setup(existing: Record<string, unknown>[] = []) {
  const rows = [...existing];
  const filters: string[] = [];
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-inn-")),
    },
    {
      data: (q: FakeQuery) => {
        const filter = q.params.get("$filter") ?? "";
        if (q.base !== "Catalog_Контрагенты" || !filter.includes("ИНН eq")) return undefined;
        filters.push(filter);
        const inn = /ИНН eq '([^']*)'/.exec(filter)?.[1];
        return rows.filter((r) => r["ИНН"] === inn && r["DeletionMark"] !== true);
      },
    },
  );
  const call = async (args: Record<string, unknown>, name = "write.counterparty.create_counterparty") => {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, rows, filters, fake };
}

describe("create_counterparty: дубль по ИНН", () => {
  it("тот же ИНН — отказ со ссылкой на существующего, ничего не пишется", async () => {
    const s = setup([{ Ref_Key: A, Description: "Есть такой", ИНН: INN, КПП: "770001001" }]);
    const r = await s.call({ name: "Новый", inn: INN });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(A);
    expect(r.text).toContain("Ничего не создано");
    expect(s.fake.requests.filter((q) => q.method === "POST")).toEqual([]);
    expect(s.filters[0]).toContain("DeletionMark eq false");
    expect(s.fake.violations).toEqual([]);
  });

  it("тот же ИНН и тот же КПП — отказ; ИНН с пробелами сравнивается без них", async () => {
    const s = setup([{ Ref_Key: A, Description: "Есть", ИНН: INN, КПП: "770001001" }]);
    const r = await s.call({ name: "Новый", inn: ` ${INN} `, kpp: "770001001" });
    expect(r.isError).toBe(true);
  });

  it("другой КПП — обособленное подразделение: предпросмотр с предупреждением", async () => {
    const s = setup([{ Ref_Key: A, Description: "Головное", ИНН: INN, КПП: "770001001" }]);
    const r = await s.call({ name: "Филиал", inn: INN, kpp: "770002002" });
    expect(r.isError, r.text).toBeFalsy();
    expect(JSON.stringify(r.structuredContent["notes"])).toContain("обособленное подразделение");
  });

  it("у найденного КПП пуст — подразделение не отличить: отказ", async () => {
    const s = setup([{ Ref_Key: A, Description: "Без КПП", ИНН: INN, КПП: "" }]);
    const r = await s.call({ name: "Новый", inn: INN, kpp: "770002002" });
    expect(r.isError).toBe(true);
  });

  it("нет совпадений или ИНН не задан — как раньше", async () => {
    const s = setup([{ Ref_Key: A, Description: "Другой", ИНН: "7700000002" }]);
    expect((await s.call({ name: "Новый", inn: INN })).isError).toBeFalsy();
    expect((await s.call({ name: "Без ИНН" })).isError).toBeFalsy();
  });

  it("повтор подтверждения с тем же operationId — не дубль сам себе: возвращает созданного", async () => {
    const s = setup();
    const preview = await s.call({ name: "Новый", inn: INN });
    const args = {
      name: "Новый",
      inn: INN,
      confirm: true,
      operationId: preview.structuredContent["operationId"],
    };
    const first = await s.call(args);
    expect(first.isError, first.text).toBeFalsy();
    // Теперь контрагент есть в 1С — проверка ИНН его находит.
    s.rows.push({ Ref_Key: first.structuredContent["ref"], Description: "Новый", ИНН: INN });
    const again = await s.call(args);
    expect(again.isError, again.text).toBeFalsy();
    expect(again.structuredContent).toMatchObject({ replayed: true, ref: first.structuredContent["ref"] });
    // А новая операция с тем же ИНН — уже дубль.
    expect((await s.call({ name: "Ещё", inn: INN })).isError).toBe(true);
  });

  it("update_counterparty: новый ИНН чужого контрагента — отказ; свой ИНН — не дубль", async () => {
    const B = "bbbbbbbb-0000-0000-0000-000000000002";
    const s = setup([
      { Ref_Key: A, Description: "Чужой", ИНН: INN },
      { Ref_Key: B, Description: "Свой", ИНН: "7700000002" },
    ]);
    const UPDATE = "write.counterparty.update_counterparty";
    const r = await s.call({ ref: B, inn: INN }, UPDATE);
    expect(r.isError).toBe(true);
    expect(r.text).toContain(A);
    expect((await s.call({ ref: B, inn: "7700000002" }, UPDATE)).isError).toBeFalsy();
    expect((await s.call({ ref: B, name: "Только имя" }, UPDATE)).isError).toBeFalsy();
  });
});
