import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { CATALOGS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { listOrganizations, resolveOrganization } from "../odata/orgs.js";
import { ensurePublished, requireEntity } from "../odata/publication.js";
import { accountsByCode, nomenclatureAccounts, type NomAccounts } from "../odata/accounting.js";
import { fetchAll } from "../odata/pagination.js";
import { contains } from "../odata/query.js";
import type { ODataEntity } from "../types/odata.js";

/** Тип ссылки на номенклатуру в табличной части (полиморфная ссылка 1С). */
const NOMENCLATURE_TYPE = "StandardODATA.Catalog_Номенклатура";

/** Виды договоров (Enum_ВидыДоговоровКонтрагентов). */
const CONTRACT_KINDS = [
  "СПокупателем",
  "СПоставщиком",
  "Прочее",
  "СКомиссионером",
  "СКомитентом",
] as const;

/** Ставки НДС (Enum_СтавкиНДС), подмножество ходовых. */
const VAT_RATES = ["БезНДС", "НДС0", "НДС5", "НДС7", "НДС10", "НДС20", "НДС22"] as const;

/** Дата → формат 1С Edm.DateTime ('YYYY-MM-DDTHH:mm:ss', без зоны). */
function odataDate(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, "");
}

/** Резолвит организацию: по названию, либо авто, если в базе ровно одна. */
async function resolveOrg(
  conn: Connection,
  organization: string | undefined,
): Promise<{ key: string; name: string }> {
  if (organization) {
    const o = await resolveOrganization(conn, organization);
    return { key: o.ref, name: o.name };
  }
  const orgs = await listOrganizations(conn);
  if (orgs.length === 1) {
    const only = orgs[0] as { ref: string; name: string };
    return { key: only.ref, name: only.name };
  }
  throw new Error(
    `В базе несколько организаций — укажите organization. Доступные: ${orgs.map((o) => o.name).join(", ")}`,
  );
}

/** Резолвит склад по названию; если не задан и склад один — берёт его; иначе undefined. */
async function resolveWarehouse(conn: Connection, name: string | undefined): Promise<string | undefined> {
  const set = resolveEntity(CATALOGS.warehouses, await conn.available());
  if (!set) return undefined;
  if (name) {
    const { rows } = await fetchAll(
      conn.client,
      set,
      { filter: contains("Description", name), select: ["Ref_Key", "Description"] },
      5,
      5,
    );
    const first = rows[0];
    if (!first) throw new Error(`Склад "${name}" не найден (см. list_entities / справочник Склады).`);
    return String(first["Ref_Key"]);
  }
  const { rows } = await fetchAll(conn.client, set, { select: ["Ref_Key"] }, 2, 2);
  return rows.length === 1 ? String((rows[0] as ODataEntity)["Ref_Key"]) : undefined;
}

/** Строит и создаёт/предпросматривает товарный документ (поступление/реализация). */
async function createGoodsDoc(
  conn: Connection,
  entitySet: string,
  p: {
    orgKey: string;
    counterpartyRef: string;
    contractRef?: string | undefined;
    warehouseKey?: string | undefined;
    date?: string | undefined;
    sumIncludesVat: boolean;
    lines: Array<{ nomenclatureRef: string; quantity: number; price: number; vatRate: string }>;
    settlement?: string | undefined; // СчетУчетаРасчетовСКонтрагентом_Key (шапка)
    lineAccountsFor: (nomRef: string, vatRate: string) => Record<string, string>; // счета строки
  },
  confirm: boolean,
) {
  const rows = p.lines.map((l, i) =>
    clean({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.quantity,
      Цена: l.price,
      Сумма: Math.round(l.quantity * l.price * 100) / 100,
      СтавкаНДС: l.vatRate,
      ...p.lineAccountsFor(l.nomenclatureRef, l.vatRate),
    }),
  );
  const total = Math.round(rows.reduce((s, r) => s + (r["Сумма"] as number), 0) * 100) / 100;
  const payload = clean({
    Date: odataDate(p.date ? new Date(`${p.date}T00:00:00`) : new Date()),
    Posted: false,
    Организация_Key: p.orgKey,
    Контрагент_Key: p.counterpartyRef,
    ДоговорКонтрагента_Key: p.contractRef,
    Склад_Key: p.warehouseKey,
    СчетУчетаРасчетовСКонтрагентом_Key: p.settlement,
    СуммаВключаетНДС: p.sumIncludesVat,
    СуммаДокумента: total,
    Товары: rows,
  });
  return createOrPreview(conn, entitySet, payload, confirm);
}

/** Первый найденный Ref среди кодов-кандидатов. */
function pickAccount(map: Map<string, string>, ...codes: string[]): string | undefined {
  for (const c of codes) {
    const r = map.get(c);
    if (r) return r;
  }
  return undefined;
}

/**
 * Резолвит счета учёта по каждой номенклатуре: сперва из регистра «Счета учёта
 * номенклатуры», недостающее добивает дефолтами по кодам плана счетов.
 */
async function resolveLineAccounts(
  conn: Connection,
  orgKey: string,
  nomRefs: string[],
  defaults: NomAccounts,
): Promise<Map<string, NomAccounts>> {
  const map = new Map<string, NomAccounts>();
  for (const ref of new Set(nomRefs)) {
    const reg = await nomenclatureAccounts(conn, orgKey, ref);
    map.set(ref, {
      goods: reg?.goods ?? defaults.goods,
      incomingVat: reg?.incomingVat ?? defaults.incomingVat,
      outgoingVat: reg?.outgoingVat ?? defaults.outgoingVat,
      income: reg?.income ?? defaults.income,
      expense: reg?.expense ?? defaults.expense,
    });
  }
  return map;
}

const confirmField = z
  .boolean()
  .default(false)
  .describe(
    "false (по умолчанию) — только предпросмотр (dry-run), запись НЕ выполняется. " +
      "true — выполнить создание. Сначала всегда показывайте dry-run и получайте согласие пользователя.",
  );

/** Убирает undefined-поля, чтобы не слать их в 1С. */
function clean(obj: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined && v !== ""));
}

const resolveSet = (conn: Connection, candidates: readonly string[], human: string): Promise<string> =>
  requireEntity(conn, candidates, human);

/**
 * Общий путь создания: при confirm=false возвращает предпросмотр (ничего не пишет),
 * при confirm=true выполняет POST. Гард записи (READ_ONLY + WRITABLE) — в клиенте.
 */
async function createOrPreview(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
  confirm: boolean,
) {
  if (!confirm) {
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      writableBase: conn.cfg.writable,
      willCreate: entitySet,
      payload,
      note: conn.cfg.writable
        ? "Предпросмотр. Чтобы создать запись, повторите вызов с confirm=true."
        : `Предпросмотр. ВНИМАНИЕ: запись в базу "${conn.cfg.name}" сейчас запрещена — включите ODATA_DB_${conn.cfg.name.toUpperCase()}_WRITABLE=true (и READ_ONLY=false).`,
    });
  }
  const created = await conn.client.create<ODataEntity>(entitySet, payload);
  return ok({
    created: true,
    database: conn.cfg.name,
    entitySet,
    ref: created["Ref_Key"],
    code: created["Code"],
    description: created["Description"],
  });
}

/** Общий путь изменения (PATCH): dry-run при confirm=false, иначе применяет. */
async function patchOrPreview(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
) {
  const guid = ref.replace(/[{}']/g, "");
  if (!confirm) {
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      willPatch: `${entitySet}(${guid})`,
      fields,
      note: conn.cfg.writable
        ? "Предпросмотр изменения. Чтобы применить, повторите с confirm=true."
        : `Предпросмотр. ВНИМАНИЕ: запись в базу "${conn.cfg.name}" запрещена — включите ODATA_DB_${conn.cfg.name.toUpperCase()}_WRITABLE=true (и READ_ONLY=false).`,
    });
  }
  const updated = await conn.client.patch<ODataEntity>(`${entitySet}(guid'${guid}')?$format=json`, fields);
  return ok({
    updated: true,
    database: conn.cfg.name,
    entitySet,
    ref: updated["Ref_Key"] ?? guid,
    description: updated["Description"],
  });
}

export function registerWriteTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "create_counterparty",
    {
      title: "Создать контрагента",
      description:
        "Создаёт нового контрагента в справочнике. ПО УМОЛЧАНИЮ это предпросмотр (dry-run) — " +
        "сначала покажите пользователю, что будет создано, и только после явного согласия " +
        "вызовите повторно с confirm=true. Запись идёт в боевую базу, поэтому без подтверждения не создавайте.",
      inputSchema: {
        database: databaseField,
        name: z.string().min(1).describe("Наименование контрагента (Description)"),
        inn: z.string().optional().describe("ИНН"),
        kpp: z.string().optional().describe("КПП"),
        fullName: z.string().optional().describe("Полное наименование"),
        legalType: z
          .enum(["ЮридическоеЛицо", "ФизическоеЛицо"])
          .optional()
          .describe("Юридическое или физическое лицо"),
        confirm: confirmField,
      },
    },
    ({ database, name, inn, kpp, fullName, legalType, confirm }) =>
      guard("create_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Контрагенты");
        const payload = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
          ЮридическоеФизическоеЛицо: legalType,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "mark_for_deletion",
    {
      title: "Пометить на удаление",
      description:
        "Ставит/снимает пометку на удаление у объекта (как в 1С) — мягкое удаление, " +
        "физически данные не стираются. По умолчанию предпросмотр (dry-run); для применения " +
        "повторите с confirm=true. Полезно, чтобы убрать ошибочно созданную запись.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string().describe("Имя объекта, напр. Catalog_Контрагенты"),
        ref: z.string().describe("Ref_Key объекта (GUID)"),
        mark: z.boolean().default(true).describe("true — пометить на удаление, false — снять пометку"),
        confirm: confirmField,
      },
    },
    ({ database, entitySet, ref, mark, confirm }) =>
      guard("mark_for_deletion", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const guid = ref.replace(/[{}']/g, "");
        const path = `${entitySet}(guid'${guid}')?$format=json`;
        if (!confirm) {
          return ok({
            dryRun: true,
            database: conn.cfg.name,
            willPatch: `${entitySet}(${guid})`,
            payload: { DeletionMark: mark },
            note: "Предпросмотр. Чтобы применить пометку, повторите с confirm=true.",
          });
        }
        const updated = await conn.client.patch<ODataEntity>(path, { DeletionMark: mark });
        return ok({
          updated: true,
          database: conn.cfg.name,
          ref: updated["Ref_Key"] ?? guid,
          deletionMark: updated["DeletionMark"] ?? mark,
        });
      }),
  );

  server.registerTool(
    "create_nomenclature",
    {
      title: "Создать номенклатуру",
      description:
        "Создаёт новую позицию номенклатуры (товар/услуга). ПО УМОЛЧАНИЮ предпросмотр (dry-run): " +
        "покажите пользователю payload и только после согласия повторите с confirm=true. " +
        "Запись идёт в боевую базу.",
      inputSchema: {
        database: databaseField,
        name: z.string().min(1).describe("Наименование номенклатуры (Description)"),
        fullName: z.string().optional().describe("Полное наименование"),
        confirm: confirmField,
      },
    },
    ({ database, name, fullName, confirm }) =>
      guard("create_nomenclature", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.nomenclature, "Номенклатура");
        const payload = clean({ Description: name, НаименованиеПолное: fullName });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "create_contract",
    {
      title: "Создать договор",
      description:
        "Заводит договор контрагента (подчинённый справочник): владелец — контрагент, " +
        "плюс организация и вид договора. По умолчанию предпросмотр (dry-run); создание — при confirm=true. " +
        "Ref контрагента берётся из find_counterparty, организация — по названию (list_organizations).",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента (владелец договора)"),
        name: z.string().min(1).describe("Наименование/номер договора (Description)"),
        kind: z
          .enum(CONTRACT_KINDS)
          .default("СПокупателем")
          .describe("Вид договора"),
        confirm: confirmField,
      },
    },
    ({ database, organization, counterpartyRef, name, kind, confirm }) =>
      guard("create_contract", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.contracts, "Договоры контрагентов");
        const org = await resolveOrg(conn, organization);
        const payload = clean({
          Description: name,
          Owner_Key: counterpartyRef,
          ВидДоговора: kind,
          Организация_Key: org.key,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "create_invoice",
    {
      title: "Создать счёт покупателю",
      description:
        "Создаёт счёт на оплату покупателю (документ) с табличной частью «Товары». " +
        "Документ создаётся НЕПРОВЕДЁННЫМ (черновик) — провести можно вручную в 1С или инструментом post_document. " +
        "По умолчанию предпросмотр (dry-run); создание — при confirm=true. " +
        "Контрагент и (опц.) договор — по Ref_Key; позиции — по Ref_Key номенклатуры.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора (необязательно)"),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        lines: z
          .array(
            z.object({
              nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
              quantity: z.number().positive().describe("Количество"),
              price: z.number().nonnegative().describe("Цена за единицу"),
              vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
            }),
          )
          .min(1)
          .describe("Позиции счёта"),
        confirm: confirmField,
      },
    },
    ({ database, organization, counterpartyRef, contractRef, date, sumIncludesVat, lines, confirm }) =>
      guard("create_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.customerInvoice, "Документ «Счёт на оплату покупателю»");
        const org = await resolveOrg(conn, organization);
        const docDate = odataDate(date ? new Date(`${date}T00:00:00`) : new Date());

        const rows = lines.map((l, i) => {
          const сумма = Math.round(l.quantity * l.price * 100) / 100;
          return {
            LineNumber: i + 1,
            Номенклатура: l.nomenclatureRef,
            Номенклатура_Type: NOMENCLATURE_TYPE,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: сумма,
            СтавкаНДС: l.vatRate,
          };
        });
        const total = Math.round(rows.reduce((s, r) => s + r.Сумма, 0) * 100) / 100;

        const payload = clean({
          Date: docDate,
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: total,
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "post_document",
    {
      title: "Провести / отменить проведение документа",
      description:
        "Проводит документ (влияет на учёт) или отменяет проведение. По умолчанию предпросмотр (dry-run); " +
        "применение — при confirm=true. ВНИМАНИЕ: проведение меняет учётные данные — используйте осознанно.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_СчетНаОплатуПокупателю"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        post: z.boolean().default(true).describe("true — провести, false — отменить проведение"),
        confirm: confirmField,
      },
    },
    ({ database, entitySet, ref, post, confirm }) =>
      guard("post_document", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const guid = ref.replace(/[{}']/g, "");
        const action = post ? "Post" : "Unpost";
        const path = `${entitySet}(guid'${guid}')/${action}?$format=json`;
        if (!confirm) {
          return ok({
            dryRun: true,
            database: conn.cfg.name,
            willCall: `${entitySet}(${guid})/${action}`,
            note: "Предпросмотр. Чтобы применить, повторите с confirm=true.",
          });
        }
        await conn.client.action(path);
        return ok({ done: true, database: conn.cfg.name, ref: guid, action });
      }),
  );

  server.registerTool(
    "update_entity",
    {
      title: "Изменить объект (общий)",
      description:
        "Изменяет произвольные поля объекта 1С через PATCH по Ref_Key. Имена полей — технические " +
        "(узнать через describe_entity), напр. {\"ИНН\":\"7701234567\"}. По умолчанию предпросмотр (dry-run); " +
        "применение — при confirm=true. Мощный инструмент: меняйте только те поля, что указали.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string().describe("Имя объекта, напр. Catalog_Контрагенты"),
        ref: z.string().describe("Ref_Key объекта (GUID)"),
        fields: z
          .record(z.union([z.string(), z.number(), z.boolean(), z.null()]))
          .describe("Поля для изменения: { техническоеИмя: значение }"),
        confirm: confirmField,
      },
    },
    ({ database, entitySet, ref, fields, confirm }) =>
      guard("update_entity", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        if (Object.keys(fields).length === 0) return fail("Не заданы поля для изменения.");
        return patchOrPreview(conn, entitySet, ref, fields, confirm);
      }),
  );

  server.registerTool(
    "update_counterparty",
    {
      title: "Изменить контрагента",
      description:
        "Изменяет реквизиты контрагента по Ref_Key (наименование/ИНН/КПП/полное имя). " +
        "По умолчанию предпросмотр (dry-run); применение — при confirm=true. Указывайте только меняемые поля.",
      inputSchema: {
        database: databaseField,
        ref: z.string().describe("Ref_Key контрагента"),
        name: z.string().optional().describe("Новое наименование (Description)"),
        inn: z.string().optional().describe("ИНН"),
        kpp: z.string().optional().describe("КПП"),
        fullName: z.string().optional().describe("Полное наименование"),
        confirm: confirmField,
      },
    },
    ({ database, ref, name, inn, kpp, fullName, confirm }) =>
      guard("update_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Справочник «Контрагенты»");
        const fields = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
        });
        if (Object.keys(fields).length === 0) return fail("Не задано ни одного поля для изменения.");
        return patchOrPreview(conn, set, ref, fields, confirm);
      }),
  );

  // Общая схема входов для товарных документов (поступление/реализация).
  const goodsDocInput = {
    database: databaseField,
    organization: organizationField,
    counterpartyRef: z.string().describe("Ref_Key контрагента"),
    contractRef: z.string().optional().describe("Ref_Key договора (необязательно)"),
    warehouse: z.string().optional().describe("Название склада (если в базе несколько)"),
    date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
    sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
    lines: z
      .array(
        z.object({
          nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
          quantity: z.number().positive().describe("Количество"),
          price: z.number().nonnegative().describe("Цена за единицу"),
          vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
        }),
      )
      .min(1)
      .describe("Позиции документа"),
    confirm: confirmField,
  };

  server.registerTool(
    "create_purchase",
    {
      title: "Создать поступление от поставщика",
      description:
        "Создаёт документ «Поступление товаров и услуг» (закупка у поставщика) с табличной частью «Товары». " +
        "Документ НЕПРОВЕДЁННЫЙ; провести — вручную в 1С или post_document (тогда 1С сформирует проводки Дт 41/19 Кт 60). " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — поставщик, договор — вида «СПоставщиком».",
      inputSchema: goodsDocInput,
    },
    ({ database, organization, counterpartyRef, contractRef, warehouse, date, sumIncludesVat, lines, confirm }) =>
      guard("create_purchase", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.purchases, "Документ «Поступление товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        // Счета учёта: из регистра по номенклатуре, дефолты по кодам (Дт 41 Кт 60, +19 при НДС).
        const codes = await accountsByCode(conn, ["41.01", "41", "60.01", "60", "19.03", "19"]);
        const defaults: NomAccounts = {
          goods: pickAccount(codes, "41.01", "41"),
          incomingVat: pickAccount(codes, "19.03", "19"),
        };
        const accMap = await resolveLineAccounts(conn, org.key, lines.map((l) => l.nomenclatureRef), defaults);
        const lineAccountsFor = (nomRef: string, vat: string): Record<string, string> => {
          const a = accMap.get(nomRef);
          return clean({
            СчетУчета_Key: a?.goods,
            ...(vat !== "БезНДС" ? { СчетУчетаНДС_Key: a?.incomingVat } : {}),
          }) as Record<string, string>;
        };
        return createGoodsDoc(
          conn,
          set,
          { orgKey: org.key, counterpartyRef, contractRef, warehouseKey, date, sumIncludesVat, lines, settlement: pickAccount(codes, "60.01", "60"), lineAccountsFor },
          confirm,
        );
      }),
  );

  server.registerTool(
    "create_shipment",
    {
      title: "Создать реализацию покупателю",
      description:
        "Создаёт документ «Реализация товаров и услуг» (отгрузка покупателю) с табличной частью «Товары». " +
        "Документ НЕПРОВЕДЁННЫЙ; провести — вручную в 1С или post_document (тогда 1С сформирует проводки Дт 62 Кт 90, Дт 90 Кт 41 и др.). " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — покупатель, договор — вида «СПокупателем».",
      inputSchema: goodsDocInput,
    },
    ({ database, organization, counterpartyRef, contractRef, warehouse, date, sumIncludesVat, lines, confirm }) =>
      guard("create_shipment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        // Счета учёта: из регистра по номенклатуре, дефолты по кодам (Дт 62 Кт 90.01, Дт 90.02 Кт 41, +90.03 при НДС).
        const codes = await accountsByCode(conn, ["41.01", "41", "62.01", "62", "90.01.1", "90.01", "90.02.1", "90.02", "90.03"]);
        const defaults: NomAccounts = {
          goods: pickAccount(codes, "41.01", "41"),
          income: pickAccount(codes, "90.01.1", "90.01"),
          expense: pickAccount(codes, "90.02.1", "90.02"),
          outgoingVat: pickAccount(codes, "90.03"),
        };
        const accMap = await resolveLineAccounts(conn, org.key, lines.map((l) => l.nomenclatureRef), defaults);
        const lineAccountsFor = (nomRef: string, vat: string): Record<string, string> => {
          const a = accMap.get(nomRef);
          return clean({
            СчетУчета_Key: a?.goods,
            СчетДоходов_Key: a?.income,
            СчетРасходов_Key: a?.expense,
            ...(vat !== "БезНДС" ? { СчетУчетаНДСПоРеализации_Key: a?.outgoingVat } : {}),
          }) as Record<string, string>;
        };
        return createGoodsDoc(
          conn,
          set,
          { orgKey: org.key, counterpartyRef, contractRef, warehouseKey, date, sumIncludesVat, lines, settlement: pickAccount(codes, "62.01", "62"), lineAccountsFor },
          confirm,
        );
      }),
  );
}
