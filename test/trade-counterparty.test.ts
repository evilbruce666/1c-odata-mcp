import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Контрагент Управления торговлей принадлежит партнёру: два инструмента — create_partner, затем create_counterparty
 * с partnerRef; каждый вызов — один объект своей операцией журнала. Критерий: подтверждение пишет ровно то, что показал
 * предпросмотр; повтор с тем же operationId не пишет ничего. Значения перечислений — с живой УТ 11.5.
 */

const P = "aaaaaaaa-0000-0000-0000-00000000000a";
const SET = "write.counterparty.create_counterparty";

interface Post {
  set: string;
  body: Record<string, unknown>;
}

function setup(
  o: {
    schema?: FakeSchema;
    partners?: Record<string, unknown>[];
    counterparties?: Record<string, unknown>[];
  } = {},
) {
  const posts: Post[] = [];
  const partners = [...(o.partners ?? [])];
  const fake = fake1C(
    {
      READ_ONLY: "false",
      ODATA_WRITABLE: "true",
      ODATA_WRITE_JOURNAL_DIR: mkdtempSync(join(tmpdir(), "1c-odata-tcp-")),
    },
    {
      schema: o.schema ?? "ut11",
      data: (q: FakeQuery) => {
        if (q.base === "Catalog_Партнеры" && !q.params.get("$filter")) return partners;
        // Дублей по ИНН нет.
        if (q.base === "Catalog_Контрагенты" && (q.params.get("$filter") ?? "").includes("ИНН eq")) return [];
        if (q.base === "Catalog_Контрагенты" && !q.params.get("$filter") && o.counterparties)
          return o.counterparties;
        return undefined;
      },
    },
  );
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = decodeURIComponent(
      typeof input === "string" ? input : input instanceof URL ? input.href : input.url,
    );
    if ((init?.method ?? "GET").toUpperCase() === "POST") {
      const set = /odata\/standard\.odata\/([^?(]+)/.exec(url)?.[1] ?? "";
      posts.push({ set, body: JSON.parse(String(init?.body)) as Record<string, unknown> });
    }
    return inner(input, init);
  });
  const call = async (args: Record<string, unknown>, name = SET) => {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent: Record<string, unknown>;
      content: { text: string }[];
    };
    return { ...r, sc: r.structuredContent ?? {}, text: r.content.map((c) => c.text).join("") };
  };
  return { call, posts, fake, partners };
}

const base = { name: "ZZ-TEST Ромашка", inn: "7700000001", partnerRef: P };
const PARTNER = "write.counterparty.create_partner";
const withPartner = () => ({
  partners: [{ Ref_Key: P, Description: "ZZ-TEST партнёр", DeletionMark: false }],
});

describe("create_counterparty в Управлении торговлей", () => {
  it("контрагент к партнёру: одна запись, тело — как в предпросмотре; повтор с тем же operationId — ни одной", async () => {
    const s = setup(withPartner());
    const preview = await s.call(base);
    expect(preview.isError, preview.text.slice(0, 400)).toBeFalsy();
    expect(preview.sc["payload"]).toMatchObject({
      Партнер_Key: P,
      ЮрФизЛицо: "ЮрЛицо",
      ЮридическоеФизическоеЛицо: "ЮридическоеЛицо",
      ИНН: base.inn,
    });
    const args = { ...base, confirm: true, operationId: preview.sc["operationId"] };
    const done = await s.call(args);
    expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
    const again = await s.call(args);
    expect(again.sc).toMatchObject({ replayed: true, ref: done.sc["ref"] });
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_Контрагенты"]);
    expect(s.posts[0]!.body).toMatchObject(preview.sc["payload"] as Record<string, unknown>);
    expect(s.fake.violations).toEqual([]);
  });

  it("без partnerRef — отказ с подсказкой про create_partner, записей нет", async () => {
    const s = setup();
    const r = await s.call({ name: "ZZ-TEST", inn: "7700000001" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("create_partner");
    expect(s.posts).toEqual([]);
  });

  it.each([
    ["партнёр помечен на удаление", [{ Ref_Key: P, Description: "Старый", DeletionMark: true }], "помечен"],
    [
      "партнёра нет",
      [{ Ref_Key: "bbbbbbbb-0000-0000-0000-00000000000b", Description: "Другой" }],
      "не найден",
    ],
  ] as const)("partnerRef: %s — отказ", async (_, partners, text) => {
    const s = setup({ partners: [...partners] });
    const r = await s.call({ name: "ZZ-TEST", inn: "7700000002", partnerRef: P });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(text);
    expect(s.posts).toEqual([]);
  });

  it("ИНН из 12 цифр без legalType — отказ; ИП и частное лицо — свои значения перечислений", async () => {
    const s = setup(withPartner());
    expect((await s.call({ ...base, inn: "770000000012" })).isError).toBe(true);
    const ip = await s.call({ ...base, inn: "770000000012", legalType: "ИндивидуальныйПредприниматель" });
    expect(ip.sc["payload"]).toMatchObject({
      ЮрФизЛицо: "ИндивидуальныйПредприниматель",
      ЮридическоеФизическоеЛицо: "ФизическоеЛицо",
    });
    const person = await s.call({ ...base, inn: undefined, legalType: "ФизическоеЛицо" });
    expect(person.sc["payload"]).toMatchObject({
      ЮрФизЛицо: "ФизЛицо",
      ЮридическоеФизическоеЛицо: "ФизическоеЛицо",
    });
    expect(s.fake.violations).toEqual([]);
  });

  it("в Бухгалтерии — как раньше: без партнёра, ИП — физическое лицо; partnerRef — в notes", async () => {
    const s = setup({ schema: "bp30" });
    const r = await s.call({ ...base, legalType: "ИндивидуальныйПредприниматель" });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["payload"]).toMatchObject({ ЮридическоеФизическоеЛицо: "ФизическоеЛицо" });
    expect(r.sc["payload"]).not.toHaveProperty("Партнер_Key");
    expect(JSON.stringify(r.sc["notes"])).toContain("partnerRef не применён");
  });

  it("create_partner: компания-клиент, одна запись как в предпросмотре; повтор — ни одной", async () => {
    const s = setup();
    const args = { name: "ZZ-TEST партнёр", relations: ["client"] };
    const preview = await s.call(args, PARTNER);
    expect(preview.isError, preview.text.slice(0, 300)).toBeFalsy();
    expect(preview.sc["willCreate"]).toBe("Catalog_Партнеры");
    expect(preview.sc["payload"]).toMatchObject({
      Description: "ZZ-TEST партнёр",
      ЮрФизЛицо: "Компания",
      Клиент: true,
      Поставщик: false,
      ПрочиеОтношения: false,
    });
    expect(String((preview.sc["payload"] as Record<string, unknown>)["ДатаРегистрации"])).toMatch(
      /T00:00:00$/,
    );
    const confirmArgs = { ...args, confirm: true, operationId: preview.sc["operationId"] };
    const done = await s.call(confirmArgs, PARTNER);
    expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
    expect((await s.call(confirmArgs, PARTNER)).sc).toMatchObject({ replayed: true, ref: done.sc["ref"] });
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_Партнеры"]);
    expect(s.posts[0]!.body).toMatchObject(preview.sc["payload"] as Record<string, unknown>);
    expect(s.fake.violations).toEqual([]);
  });

  it("create_partner: частное лицо-поставщик; без relations — отказ схемы; в Бухгалтерии — отказ", async () => {
    const s = setup();
    const r = await s.call(
      { name: "ZZ-TEST", privatePerson: true, relations: ["supplier", "other"] },
      PARTNER,
    );
    expect(r.sc["payload"]).toMatchObject({
      ЮрФизЛицо: "ЧастноеЛицо",
      Клиент: false,
      Поставщик: true,
      ПрочиеОтношения: true,
    });
    expect((await s.call({ name: "ZZ-TEST", relations: [] }, PARTNER)).isError).toBe(true);
    const bp = await setup({ schema: "bp30" }).call({ name: "ZZ-TEST", relations: ["client"] }, PARTNER);
    expect(bp.isError).toBe(true);
    expect(s.posts).toEqual([]);
  });

  it("create_partner: подтверждение после полуночи — дата регистрации дня предпросмотра", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    try {
      vi.setSystemTime(new Date(2026, 10, 20, 23, 59, 58));
      const s = setup();
      const args = { name: "ZZ-TEST", relations: ["client"] };
      const preview = await s.call(args, PARTNER);
      vi.setSystemTime(new Date(2026, 10, 21, 0, 0, 5));
      const done = await s.call({ ...args, confirm: true, operationId: preview.sc["operationId"] }, PARTNER);
      expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
      expect(s.posts[0]!.body["ДатаРегистрации"]).toBe("2026-11-20T00:00:00");
    } finally {
      vi.useRealTimers();
    }
  });

  it("update_counterparty в УТ проходит защиту: предпросмотр правки, записей нет", async () => {
    const s = setup();
    const r = await s.call(
      { ref: "cccccccc-0000-0000-0000-00000000000c", name: "ZZ-TEST новое имя" },
      "write.counterparty.update_counterparty",
    );
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["dryRun"]).toBe(true);
    expect(s.posts).toEqual([]);
    expect(s.fake.violations).toEqual([]);
  });

  it.each(["Catalog_Контрагенты", "Catalog_Партнеры"])(
    "mark_for_deletion в УТ: %s — проходит защиту (предпросмотр); склады — отказ",
    async (entitySet) => {
      const s = setup();
      const ref = "cccccccc-0000-0000-0000-00000000000c";
      const r = await s.call({ entitySet, ref, mark: true }, "write.entity.mark_for_deletion");
      expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
      expect(r.sc["dryRun"]).toBe(true);
      const other = await s.call(
        { entitySet: "Catalog_Склады", ref, mark: true },
        "write.entity.mark_for_deletion",
      );
      expect(other.isError).toBe(true);
      expect(other.text).toContain("только для Бухгалтерии");
      expect(s.fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
    },
  );

  const C = "dddddddd-0000-0000-0000-00000000000d";
  const bank = { ownerRef: C, accountNumber: "40702810000000000001", bik: "044525225" };

  it("create_bank_account в УТ: счёт контрагента, банк из классификатора; одна запись", async () => {
    const s = setup();
    const preview = await s.call(bank, "write.counterparty.create_bank_account");
    expect(preview.isError, preview.text.slice(0, 300)).toBeFalsy();
    expect(preview.sc["willCreate"]).toBe("Catalog_БанковскиеСчетаКонтрагентов");
    expect(preview.sc["payload"]).toMatchObject({
      Owner: C,
      Owner_Type: "StandardODATA.Catalog_Контрагенты",
      НомерСчета: bank.accountNumber,
    });
    expect(s.fake.requests.some((q) => q.resource.startsWith("Catalog_КлассификаторБанков"))).toBe(true);
    const done = await s.call(
      { ...bank, confirm: true, operationId: preview.sc["operationId"] },
      "write.counterparty.create_bank_account",
    );
    expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_БанковскиеСчетаКонтрагентов"]);
    expect(s.fake.violations).toEqual([]);
  });

  it("create_bank_account в УТ с makeMain — отказ: основного счёта у контрагента УТ нет", async () => {
    const s = setup();
    const r = await s.call({ ...bank, makeMain: true }, "write.counterparty.create_bank_account");
    expect(r.isError).toBe(true);
    expect(r.text).toContain("makeMain");
    expect(s.posts).toEqual([]);
  });

  it("create_contact_person в УТ: владелец — партнёр контрагента; дата связи — день предпросмотра", async () => {
    const s = setup({ counterparties: [{ Ref_Key: C, Партнер_Key: P }] });
    const args = { ownerRef: C, name: "ZZ-TEST Иванов Иван", position: "Директор" };
    const preview = await s.call(args, "write.counterparty.create_contact_person");
    expect(preview.isError, preview.text.slice(0, 300)).toBeFalsy();
    expect(preview.sc["willCreate"]).toBe("Catalog_КонтактныеЛицаПартнеров");
    expect(preview.sc["payload"]).toMatchObject({ Owner_Key: P, ДолжностьПоВизитке: "Директор" });
    const done = await s.call(
      { ...args, confirm: true, operationId: preview.sc["operationId"] },
      "write.counterparty.create_contact_person",
    );
    expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_КонтактныеЛицаПартнеров"]);
    expect(s.posts[0]!.body["ДатаРегистрацииСвязи"]).toBe(
      (preview.sc["payload"] as Record<string, unknown>)["ДатаРегистрацииСвязи"],
    );
    expect(s.fake.violations).toEqual([]);
  });

  it("create_contact_person в УТ: ссылка на партнёра тоже подходит; неизвестная — отказ; makeMain — отказ", async () => {
    const s = setup({
      counterparties: [{ Ref_Key: "eeeeeeee-0000-0000-0000-00000000000e", Партнер_Key: P }],
      partners: [{ Ref_Key: P, Description: "Партнёр" }],
    });
    const CP = "write.counterparty.create_contact_person";
    const byPartner = await s.call({ ownerRef: P, name: "ZZ-TEST" }, CP);
    expect(byPartner.isError, byPartner.text.slice(0, 300)).toBeFalsy();
    expect(byPartner.sc["payload"]).toMatchObject({ Owner_Key: P });
    const unknown = await s.call({ ownerRef: "ffffffff-0000-0000-0000-00000000000f", name: "ZZ-TEST" }, CP);
    expect(unknown.isError).toBe(true);
    expect(unknown.text).toContain("не найден");
    expect((await s.call({ ownerRef: P, name: "ZZ-TEST", makeMain: true }, CP)).isError).toBe(true);
    expect(s.posts).toEqual([]);
  });

  it("create_contact_person в УТ: подтверждение после полуночи совпадает с предпросмотром", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    try {
      vi.setSystemTime(new Date(2026, 10, 20, 23, 59, 58));
      const s = setup({ counterparties: [{ Ref_Key: C, Партнер_Key: P }] });
      const args = { ownerRef: C, name: "ZZ-TEST Петров" };
      const CP = "write.counterparty.create_contact_person";
      const preview = await s.call(args, CP);
      vi.setSystemTime(new Date(2026, 10, 21, 0, 0, 5));
      const done = await s.call({ ...args, confirm: true, operationId: preview.sc["operationId"] }, CP);
      expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
      expect(s.posts[0]!.body["ДатаРегистрацииСвязи"]).toBe("2026-11-20T00:00:00");
    } finally {
      vi.useRealTimers();
    }
  });

  it.each([
    ["write.counterparty.update_counterparty", { name: "X" }],
    ["write.catalog.update_nomenclature", { name: "X" }],
    ["write.entity.move_to_folder", { entitySet: "Catalog_Номенклатура", folder: P }],
  ] as const)(
    "%s: ref с «/../» — отказ схемы, ничего не отправлено (ревью: обход защиты)",
    async (name, args) => {
      const s = setup();
      const ref = `${P})/../Catalog_Организации(guid'${P}')?x=`;
      const r = await s.call({ ...args, ref, confirm: true }, name);
      expect(r.isError).toBe(true);
      expect(s.fake.requests.filter((q) => q.method !== "GET")).toEqual([]);
    },
  );
});
