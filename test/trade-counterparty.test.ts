import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Контрагент Управления торговлей = партнёр + контрагент. Критерий: подтверждение создаёт ровно то, что показал
 * предпросмотр (партнёр, затем контрагент с его Ref_Key); повтор не создаёт второго; если 1С отказала контрагенту
 * после партнёра, повтор с partnerRef не создаёт второго партнёра. Значения перечислений — с живой УТ 11.5.
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
    rejectCounterparty?: number;
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
  let rejects = o.rejectCounterparty ?? 0;
  const inner = globalThis.fetch;
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = decodeURIComponent(
      typeof input === "string" ? input : input instanceof URL ? input.href : input.url,
    );
    if ((init?.method ?? "GET").toUpperCase() === "POST") {
      const set = /odata\/standard\.odata\/([^?(]+)/.exec(url)?.[1] ?? "";
      posts.push({ set, body: JSON.parse(String(init?.body)) as Record<string, unknown> });
      if (set === "Catalog_Контрагенты" && rejects > 0) {
        rejects--;
        return new Response(
          JSON.stringify({
            "odata.error": { code: "-1", message: { lang: "ru", value: "Synthetic rejection" } },
          }),
          { status: 400, headers: { "content-type": "application/json" } },
        );
      }
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

const base = { name: "ZZ-TEST Ромашка", inn: "7700000001", relations: ["client"] };

async function confirmed(s: ReturnType<typeof setup>, args: Record<string, unknown>) {
  const preview = await s.call(args);
  expect(preview.isError, preview.text.slice(0, 400)).toBeFalsy();
  const done = await s.call({ ...args, confirm: true, operationId: preview.sc["operationId"] });
  return { preview, done };
}

describe("create_counterparty в Управлении торговлей", () => {
  it("предпросмотр: партнёр и контрагент, Партнер_Key — Ref_Key будущего партнёра; записей нет", async () => {
    const s = setup();
    const r = await s.call(base);
    expect(r.isError, r.text.slice(0, 400)).toBeFalsy();
    const partner = r.sc["partner"] as { willCreate: string; ref: string; payload: Record<string, unknown> };
    expect(partner.willCreate).toBe("Catalog_Партнеры");
    expect(partner.payload).toMatchObject({
      Description: base.name,
      ЮрФизЛицо: "Компания",
      Клиент: true,
      Поставщик: false,
      ПрочиеОтношения: false,
    });
    expect(String(partner.payload["ДатаРегистрации"])).toMatch(/^\d{4}-\d\d-\d\dT00:00:00$/);
    expect(partner.ref).toMatch(/^[0-9a-f-]{36}$/);
    expect(partner.ref).not.toBe("00000000-0000-0000-0000-000000000000");
    expect(r.sc["payload"]).toMatchObject({
      Партнер_Key: partner.ref,
      ЮрФизЛицо: "ЮрЛицо",
      ЮридическоеФизическоеЛицо: "ЮридическоеЛицо",
      ИНН: base.inn,
    });
    expect(s.posts).toEqual([]);
    expect(s.fake.violations).toEqual([]);
  });

  it("подтверждение: две записи — партнёр, затем контрагент с его Ref_Key; совпадают с предпросмотром", async () => {
    const s = setup();
    const { preview, done } = await confirmed(s, base);
    expect(done.isError, done.text.slice(0, 400)).toBeFalsy();
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_Партнеры", "Catalog_Контрагенты"]);
    const partner = preview.sc["partner"] as { ref: string; payload: Record<string, unknown> };
    expect(s.posts[0]!.body).toMatchObject({ Ref_Key: partner.ref, ...partner.payload });
    expect(s.posts[1]!.body).toMatchObject({ Партнер_Key: partner.ref });
    expect(done.sc).toMatchObject({ created: true, partnerRef: partner.ref, partnerCreated: true });
    expect(s.fake.violations).toEqual([]);
  });

  it("повтор подтверждения с тем же operationId — ни одной новой записи", async () => {
    const s = setup();
    const preview = await s.call(base);
    const args = { ...base, confirm: true, operationId: preview.sc["operationId"] };
    const first = await s.call(args);
    const again = await s.call(args);
    expect(first.isError, first.text).toBeFalsy();
    expect(again.sc).toMatchObject({ replayed: true, ref: first.sc["ref"] });
    expect(s.posts).toHaveLength(2);
  });

  it("1С отказала контрагенту после партнёра: ссылка на партнёра; повтор с partnerRef — без второго партнёра", async () => {
    const s = setup({ rejectCounterparty: 1 });
    const { done } = await confirmed(s, base);
    expect(done.isError).toBe(true);
    const partnerRef = /partnerRef=([0-9a-f-]{36})/.exec(done.text)?.[1];
    expect(partnerRef, done.text).toBeTruthy();
    // Партнёр теперь есть в 1С.
    s.partners.push({ Ref_Key: partnerRef, Description: base.name, DeletionMark: false });
    const retry = await confirmed(s, { ...base, relations: undefined, partnerRef });
    expect(retry.done.isError, retry.done.text.slice(0, 300)).toBeFalsy();
    expect(s.posts.filter((p) => p.set === "Catalog_Партнеры")).toHaveLength(1);
    expect(s.posts.at(-1)).toMatchObject({ set: "Catalog_Контрагенты", body: { Партнер_Key: partnerRef } });
  });

  it("partnerRef существующего партнёра — одна запись, relations не нужны", async () => {
    const s = setup({ partners: [{ Ref_Key: P, Description: "Есть", DeletionMark: false }] });
    const { preview, done } = await confirmed(s, {
      name: "ZZ-TEST Второе юрлицо",
      inn: "7700000002",
      partnerRef: P,
    });
    expect(preview.sc["partner"]).toBeUndefined();
    expect(done.isError, done.text.slice(0, 300)).toBeFalsy();
    expect(s.posts.map((p) => p.set)).toEqual(["Catalog_Контрагенты"]);
    expect(s.posts[0]!.body).toMatchObject({ Партнер_Key: P });
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

  it("без relations и partnerRef — отказ с подсказкой", async () => {
    const r = await setup().call({ name: "ZZ-TEST", inn: "7700000001" });
    expect(r.isError).toBe(true);
    expect(r.text).toContain("relations");
  });

  it("ИНН из 12 цифр без legalType — отказ; ИП и частное лицо — свои значения перечислений", async () => {
    const s = setup();
    expect((await s.call({ ...base, inn: "770000000012" })).isError).toBe(true);
    const ip = await s.call({ ...base, inn: "770000000012", legalType: "ИндивидуальныйПредприниматель" });
    expect(ip.sc["payload"]).toMatchObject({
      ЮрФизЛицо: "ИндивидуальныйПредприниматель",
      ЮридическоеФизическоеЛицо: "ФизическоеЛицо",
    });
    expect((ip.sc["partner"] as { payload: object }).payload).toMatchObject({ ЮрФизЛицо: "Компания" });
    const person = await s.call({
      ...base,
      inn: undefined,
      legalType: "ФизическоеЛицо",
      relations: ["supplier"],
    });
    expect(person.sc["payload"]).toMatchObject({
      ЮрФизЛицо: "ФизЛицо",
      ЮридическоеФизическоеЛицо: "ФизическоеЛицо",
    });
    expect((person.sc["partner"] as { payload: object }).payload).toMatchObject({
      ЮрФизЛицо: "ЧастноеЛицо",
      Клиент: false,
      Поставщик: true,
    });
    expect(s.fake.violations).toEqual([]);
  });

  it("в Бухгалтерии — как раньше: без партнёра, ИП — физическое лицо", async () => {
    const s = setup({ schema: "bp30" });
    const r = await s.call({ ...base, legalType: "ИндивидуальныйПредприниматель" });
    expect(r.isError, r.text.slice(0, 300)).toBeFalsy();
    expect(r.sc["partner"]).toBeUndefined();
    expect(r.sc["payload"]).toMatchObject({ ЮридическоеФизическоеЛицо: "ФизическоеЛицо" });
    expect(r.sc["payload"]).not.toHaveProperty("Партнер_Key");
    expect(JSON.stringify(r.sc["notes"])).toContain("partnerRef и relations не применены");
  });

  it("подтверждение после полуночи: дата регистрации — день предпросмотра, запись совпадает с ним", async () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    try {
      vi.setSystemTime(new Date(2026, 10, 20, 23, 59, 58));
      const s = setup();
      const preview = await s.call(base);
      vi.setSystemTime(new Date(2026, 10, 21, 0, 0, 5));
      const done = await s.call({ ...base, confirm: true, operationId: preview.sc["operationId"] });
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
});
