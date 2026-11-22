import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeSchema } from "./helpers/fake-1c.js";
import { SNAPSHOT_REF } from "./helpers/metadata-snapshot.js";

/**
 * Пользователи 1С для выбора ответственного: поиск по частям имени, по умолчанию без помеченных, недействительных и
 * служебных; responsibleRef в записи проверяется до записи. Имена выдуманы, схемы — снимки БП 3.0 и УТ 11.5.
 */
const g = (n: number) => `aaaaaaaa-0000-4000-8000-${String(n).padStart(12, "0")}`;
const DEP = "bbbbbbbb-0000-4000-8000-000000000001";
const USERS = [
  {
    Ref_Key: g(1),
    Description: "Иванов Пётр Сергеевич",
    DeletionMark: false,
    Недействителен: false,
    Служебный: false,
    Подразделение_Key: DEP,
  },
  { Ref_Key: g(2), Description: "Иванова Анна", DeletionMark: false, Недействителен: true, Служебный: false },
  {
    Ref_Key: g(3),
    Description: "Обработчик обновления",
    DeletionMark: false,
    Недействителен: false,
    Служебный: true,
  },
  { Ref_Key: g(4), Description: "Петров Иван", DeletionMark: true, Недействителен: false, Служебный: false },
  {
    Ref_Key: g(5),
    Description: "Сидоров Олег",
    DeletionMark: false,
    Недействителен: false,
    Служебный: false,
  },
];

async function setup(schema: FakeSchema = "ut11") {
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-users-")),
    },
    {
      schema,
      data: (q) =>
        q.base === "Catalog_Пользователи"
          ? USERS
          : q.base === "Catalog_СтруктураПредприятия" || q.base === "Catalog_ПодразделенияОрганизаций"
            ? [{ Ref_Key: DEP, Description: "Склад" }]
            : undefined,
    },
  );
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const call = async (name: string, args: object) => {
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown> & { items?: Array<Record<string, unknown>> };
      content: { text: string }[];
    };
    return { ...r, text: r.content.map((c) => c.text).join("") };
  };
  return { call, fake };
}

describe("read.catalog.list_users", () => {
  it.each(["ut11", "bp30"] as const)(
    "%s: по умолчанию только действующие, с подразделением; поиск по частям имени",
    async (schema) => {
      const { call, fake } = await setup(schema);
      const all = await call("read.catalog.list_users", {});
      expect(all.isError, all.text).toBeFalsy();
      expect(all.structuredContent.items!.map((i) => i["ref"])).toEqual([g(1), g(5)]);
      expect(all.structuredContent.items![0]).toMatchObject({
        name: "Иванов Пётр Сергеевич",
        departmentRef: DEP,
        inactive: false,
        service: false,
      });
      // Название подразделения — если справочник опубликован (в снимке БП его нет — только ссылка).
      if (schema === "ut11") expect(all.structuredContent.items![0]).toMatchObject({ department: "Склад" });
      expect(String(all.structuredContent["note"])).toContain("Скрыто 3");
      // Части слов, регистр и ё не важны.
      const found = await call("read.catalog.list_users", { query: "иван петр" });
      expect(found.structuredContent.items!.map((i) => i["ref"])).toEqual([g(1)]);
      const inactive = await call("read.catalog.list_users", { query: "иванов", includeInactive: true });
      expect(inactive.structuredContent.items!.map((i) => [i["ref"], i["inactive"]])).toEqual([
        [g(1), false],
        [g(2), true],
      ]);
      expect(fake.violations).toEqual([]);
      // Ни контактов, ни физического лица в ответе.
      expect(JSON.stringify(all.structuredContent)).not.toMatch(
        /ФизическоеЛицо|КонтактнаяИнформация|Фотография/,
      );
    },
  );
});

describe("responsibleRef проверяется до записи", () => {
  const price = (responsibleRef: string) => ({
    priceKind: SNAPSHOT_REF,
    lines: [{ itemRef: SNAPSHOT_REF, price: 1 }],
    mechanism: "legacy",
    responsibleRef,
  });
  it.each([
    [g(2), "недействителен"],
    [g(3), "служебный"],
    [g(4), "помечен на удаление"],
    [g(99), "не найден"],
    ["не-гуид", "GUID"],
  ])("create_price_setting: %s — отказ «%s», ничего не записано", async (ref, text) => {
    const { call, fake } = await setup("ut11");
    const r = await call("write.pricing.create_price_setting", price(ref));
    expect(r.isError).toBe(true);
    expect(r.text).toContain(text);
    if (ref !== "не-гуид") expect(r.text).toContain("read.catalog.list_users");
    expect(fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });

  it("create_price_setting: действующий пользователь — предпросмотр с ответственным", async () => {
    const { call } = await setup("ut11");
    const r = await call("write.pricing.create_price_setting", price(g(1).toUpperCase()));
    expect(r.isError, r.text).toBeFalsy();
    expect((r.structuredContent["payload"] as Record<string, unknown>)["Ответственный_Key"]).toBe(
      g(1).toUpperCase(),
    );
  });

  it("create_payout_order (Бухгалтерия): недействительный ответственный — отказ до записи", async () => {
    const { call, fake } = await setup("bp30");
    const r = await call("write.money.create_payout_order", {
      counterpartyRef: g(80),
      amount: 1,
      number: "1",
      purposeText: "Оплата",
      recipientText: "Получатель",
      responsibleRef: g(2),
    });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("недействителен");
    expect(fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
  });
});
