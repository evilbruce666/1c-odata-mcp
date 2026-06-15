import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, guard, databaseField, organizationField } from "./_shared.js";
import { CATALOGS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { listOrganizations, resolveOrganization } from "../odata/orgs.js";
import { ensurePublished, requireEntity } from "../odata/publication.js";
import { accountsByCode } from "../odata/accounting.js";
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

/** Счета учёта для заполнения документа (через OData автозаполнение 1С не работает). */
interface DocAccounts {
  settlement?: string | undefined; // СчетУчетаРасчетовСКонтрагентом_Key (шапка)
  line: Record<string, string>; // постоянные счета строки (СчетУчета_Key и т.п.)
  vatField?: string; // имя поля НДС-счёта в строке
  vatAccount?: string | undefined; // ref НДС-счёта (если ставка ≠ БезНДС)
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
    accounts: DocAccounts;
  },
  confirm: boolean,
) {
  const rows = p.lines.map((l, i) => {
    const vat =
      l.vatRate !== "БезНДС" && p.accounts.vatField && p.accounts.vatAccount
        ? { [p.accounts.vatField]: p.accounts.vatAccount }
        : {};
    return clean({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.quantity,
      Цена: l.price,
      Сумма: Math.round(l.quantity * l.price * 100) / 100,
      СтавкаНДС: l.vatRate,
      ...p.accounts.line,
      ...vat,
    });
  });
  const total = Math.round(rows.reduce((s, r) => s + (r["Сумма"] as number), 0) * 100) / 100;
  const payload = clean({
    Date: odataDate(p.date ? new Date(`${p.date}T00:00:00`) : new Date()),
    Posted: false,
    Организация_Key: p.orgKey,
    Контрагент_Key: p.counterpartyRef,
    ДоговорКонтрагента_Key: p.contractRef,
    Склад_Key: p.warehouseKey,
    СчетУчетаРасчетовСКонтрагентом_Key: p.accounts.settlement,
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
        // Счета учёта для проводок Дт 41 Кт 60 (+19 при НДС).
        const acc = await accountsByCode(conn, ["41.01", "41", "60.01", "60", "19.03", "19"]);
        const accounts = {
          settlement: pickAccount(acc, "60.01", "60"),
          line: clean({ СчетУчета_Key: pickAccount(acc, "41.01", "41") }) as Record<string, string>,
          vatField: "СчетУчетаНДС_Key",
          vatAccount: pickAccount(acc, "19.03", "19"),
        };
        return createGoodsDoc(
          conn,
          set,
          { orgKey: org.key, counterpartyRef, contractRef, warehouseKey, date, sumIncludesVat, lines, accounts },
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
        // Счета учёта для проводок Дт 62 Кт 90.01, Дт 90.02 Кт 41 (+90.03 при НДС).
        const acc = await accountsByCode(conn, ["41.01", "41", "62.01", "62", "90.01.1", "90.01", "90.02.1", "90.02", "90.03"]);
        const accounts = {
          settlement: pickAccount(acc, "62.01", "62"),
          line: clean({
            СчетУчета_Key: pickAccount(acc, "41.01", "41"),
            СчетДоходов_Key: pickAccount(acc, "90.01.1", "90.01"),
            СчетРасходов_Key: pickAccount(acc, "90.02.1", "90.02"),
          }) as Record<string, string>,
          vatField: "СчетУчетаНДСПоРеализации_Key",
          vatAccount: pickAccount(acc, "90.03"),
        };
        return createGoodsDoc(
          conn,
          set,
          { orgKey: org.key, counterpartyRef, contractRef, warehouseKey, date, sumIncludesVat, lines, accounts },
          confirm,
        );
      }),
  );
}
