import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { CATALOGS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { listOrganizations, resolveOrganization } from "../odata/orgs.js";
import { ensurePublished, requireEntity } from "../odata/publication.js";
import { accountsByCode, nomenclatureAccounts, type NomAccounts } from "../odata/accounting.js";
import {
  contactKindsForCounterparties,
  resolveBankByBik,
  resolveAdditionalProperty,
  type ContactKinds,
} from "../odata/refdata.js";
import { fetchAll } from "../odata/pagination.js";
import { and, buildQuery, cmp, contains, odataGuid, odataString } from "../odata/query.js";
import type { ODataEntity } from "../types/odata.js";

/** Тип ссылки на номенклатуру в табличной части (полиморфная ссылка 1С). */
const NOMENCLATURE_TYPE = "StandardODATA.Catalog_Номенклатура";
const COUNTERPARTY_TYPE = "StandardODATA.Catalog_Контрагенты";

/** Строит строки табличной части «КонтактнаяИнформация» из телефона/email/адреса. */
function buildContactRows(
  kinds: ContactKinds,
  data: { phone?: string | undefined; email?: string | undefined; address?: string | undefined },
): Array<Record<string, unknown>> {
  const rows: Array<Record<string, unknown>> = [];
  if (data.phone && kinds.phone)
    rows.push({
      Тип: kinds.phone.тип,
      Вид_Key: kinds.phone.key,
      Представление: data.phone,
      НомерТелефона: data.phone,
    });
  if (data.email && kinds.email)
    rows.push({
      Тип: kinds.email.тип,
      Вид_Key: kinds.email.key,
      Представление: data.email,
      АдресЭП: data.email,
    });
  if (data.address && kinds.address)
    rows.push({ Тип: kinds.address.тип, Вид_Key: kinds.address.key, Представление: data.address });
  return rows.map((r, i) => ({ LineNumber: i + 1, ...r }));
}

/** Дополнительные поля карточки контрагента: контактная информация + ОГРН (доп.реквизит). */
async function counterpartyExtras(
  conn: Connection,
  data: {
    phone?: string | undefined;
    email?: string | undefined;
    address?: string | undefined;
    ogrn?: string | undefined;
  },
): Promise<{ fields: Record<string, unknown>; notes: string[] }> {
  const fields: Record<string, unknown> = {};
  const notes: string[] = [];
  if (data.phone || data.email || data.address) {
    const kinds = await contactKindsForCounterparties(conn);
    const rows = buildContactRows(kinds, data);
    if (rows.length) fields["КонтактнаяИнформация"] = rows;
    if (data.phone && !kinds.phone) notes.push("Вид «Телефон» не найден — телефон не записан.");
    if (data.email && !kinds.email) notes.push("Вид «Email» не найден — email не записан.");
    if (data.address && !kinds.address) notes.push("Вид «Адрес» не найден — адрес не записан.");
  }
  if (data.ogrn) {
    const prop = await resolveAdditionalProperty(conn, "ОГРН");
    if (prop)
      fields["ДополнительныеРеквизиты"] = [{ LineNumber: 1, Свойство_Key: prop, Значение: data.ogrn }];
    else
      notes.push(
        "ОГРН не записан: в этой базе нет дополнительного реквизита «ОГРН» для контрагентов. Заведите его в 1С (Администрирование → Дополнительные реквизиты), либо вносите ОГРН вручную.",
      );
  }
  return { fields, notes };
}

/** Виды договоров (Enum_ВидыДоговоровКонтрагентов). */
const CONTRACT_KINDS = ["СПокупателем", "СПоставщиком", "Прочее", "СКомиссионером", "СКомитентом"] as const;

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

/**
 * Резолвит элемент справочника по точному коду или части наименования.
 * Публикацию проверяет (requireEntity). Бросает Error, если не найдено.
 */
async function resolveCatalogItem(
  conn: Connection,
  candidates: readonly string[],
  label: string,
  query: string,
): Promise<{ ref: string; code?: string; name?: string }> {
  const set = await requireEntity(conn, candidates, label);
  const byCode = await fetchAll(
    conn.client,
    set,
    { filter: cmp("Code", "eq", odataString(query)), select: ["Ref_Key", "Code", "Description"] },
    3,
    3,
  );
  const rows = byCode.rows.length
    ? byCode.rows
    : (
        await fetchAll(
          conn.client,
          set,
          { filter: contains("Description", query), select: ["Ref_Key", "Code", "Description"] },
          5,
          5,
        )
      ).rows;
  const first = rows[0];
  if (!first) throw new Error(`${label}: «${query}» не найдено (по коду или наименованию).`);
  return {
    ref: String(first["Ref_Key"]),
    code: first["Code"] ? String(first["Code"]) : undefined,
    name: first["Description"] ? String(first["Description"]) : undefined,
  };
}

interface GoodsLine {
  nomenclatureRef: string;
  quantity: number;
  price: number;
  vatRate: string;
}
type LineAccountsFor = (nomRef: string, vatRate: string) => Record<string, string>;

const lineSum = (l: GoodsLine): number => Math.round(l.quantity * l.price * 100) / 100;
const rowsTotal = (rows: Array<Record<string, unknown>>): number =>
  Math.round(rows.reduce((s, r) => s + (r["Сумма"] as number), 0) * 100) / 100;

/** Строки табличной части «Товары» для поступления/реализации (Номенклатура_Key + счета). */
function buildGoodsRows(
  lines: GoodsLine[],
  lineAccountsFor: LineAccountsFor,
): Array<Record<string, unknown>> {
  return lines.map((l, i) =>
    clean({
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      Количество: l.quantity,
      Цена: l.price,
      Сумма: lineSum(l),
      СтавкаНДС: l.vatRate,
      ...lineAccountsFor(l.nomenclatureRef, l.vatRate),
    }),
  );
}

/** Строки «Товары» для счёта покупателю (полиморфная ссылка Номенклатура+_Type). */
function buildInvoiceRows(lines: GoodsLine[]): Array<Record<string, unknown>> {
  return lines.map((l, i) =>
    clean({
      LineNumber: i + 1,
      Номенклатура: l.nomenclatureRef,
      Номенклатура_Type: NOMENCLATURE_TYPE,
      Количество: l.quantity,
      Цена: l.price,
      Сумма: lineSum(l),
      СтавкаНДС: l.vatRate,
    }),
  );
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
    lines: GoodsLine[];
    settlement?: string | undefined; // СчетУчетаРасчетовСКонтрагентом_Key (шапка)
    lineAccountsFor: LineAccountsFor; // счета строки
  },
  confirm: boolean,
) {
  const rows = buildGoodsRows(p.lines, p.lineAccountsFor);
  const payload = clean({
    Date: odataDate(p.date ? new Date(`${p.date}T00:00:00`) : new Date()),
    Posted: false,
    Организация_Key: p.orgKey,
    Контрагент_Key: p.counterpartyRef,
    ДоговорКонтрагента_Key: p.contractRef,
    Склад_Key: p.warehouseKey,
    СчетУчетаРасчетовСКонтрагентом_Key: p.settlement,
    СуммаВключаетНДС: p.sumIncludesVat,
    СуммаДокумента: rowsTotal(rows),
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

/**
 * Готовит счета учёта для товарного документа: счёт расчётов (шапка) и функцию
 * счетов строки по номенклатуре. kind различает поступление (Дт 41 Кт 60) и
 * реализацию (Дт 62 Кт 90, Дт 90 Кт 41).
 */
async function goodsAccounts(
  conn: Connection,
  orgKey: string,
  lines: GoodsLine[],
  kind: "purchase" | "shipment" | "service",
): Promise<{ settlement?: string | undefined; lineAccountsFor: LineAccountsFor }> {
  const nomRefs = lines.map((l) => l.nomenclatureRef);
  if (kind === "purchase") {
    const codes = await accountsByCode(conn, ["41.01", "41", "60.01", "60", "19.03", "19"]);
    const defaults: NomAccounts = {
      goods: pickAccount(codes, "41.01", "41"),
      incomingVat: pickAccount(codes, "19.03", "19"),
    };
    const accMap = await resolveLineAccounts(conn, orgKey, nomRefs, defaults);
    return {
      settlement: pickAccount(codes, "60.01", "60"),
      lineAccountsFor: (nomRef, vat) => {
        const a = accMap.get(nomRef);
        return clean({
          СчетУчета_Key: a?.goods,
          ...(vat !== "БезНДС" ? { СчетУчетаНДС_Key: a?.incomingVat } : {}),
        }) as Record<string, string>;
      },
    };
  }
  // shipment (товары) и service (услуги) — продажа: Дт 62 Кт 90, у товаров ещё Дт 90 Кт 41.
  const codes = await accountsByCode(conn, [
    "41.01",
    "41",
    "62.01",
    "62",
    "90.01.1",
    "90.01",
    "90.02.1",
    "90.02",
    "90.03",
  ]);
  const defaults: NomAccounts = {
    goods: pickAccount(codes, "41.01", "41"),
    income: pickAccount(codes, "90.01.1", "90.01"),
    expense: pickAccount(codes, "90.02.1", "90.02"),
    outgoingVat: pickAccount(codes, "90.03"),
  };
  const accMap = await resolveLineAccounts(conn, orgKey, nomRefs, defaults);
  return {
    settlement: pickAccount(codes, "62.01", "62"),
    lineAccountsFor: (nomRef, vat) => {
      const a = accMap.get(nomRef);
      const vatField = vat !== "БезНДС" ? { СчетУчетаНДСПоРеализации_Key: a?.outgoingVat } : {};
      // Услуги — без счёта учёта 41 (нет склада); товары — со счётом учёта.
      const goods = kind === "service" ? {} : { СчетУчета_Key: a?.goods };
      return clean({
        ...goods,
        СчетДоходов_Key: a?.income,
        СчетРасходов_Key: a?.expense,
        ...vatField,
      }) as Record<string, string>;
    },
  };
}

/** Читает документ: организация, проведён ли, и текущие строки «Товары» как GoodsLine[]. */
async function getDocInfo(
  conn: Connection,
  entitySet: string,
  guid: string,
): Promise<{ orgKey: string; posted: boolean; lines: GoodsLine[] }> {
  const doc = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
  const rows = (doc["Товары"] as Array<Record<string, unknown>>) ?? [];
  const lines: GoodsLine[] = rows.map((r) => ({
    nomenclatureRef: String(r["Номенклатура_Key"] ?? r["Номенклатура"] ?? ""),
    quantity: Number(r["Количество"] ?? 0),
    price: Number(r["Цена"] ?? 0),
    vatRate: String(r["СтавкаНДС"] ?? "БезНДС"),
  }));
  return { orgKey: String(doc["Организация_Key"] ?? ""), posted: doc["Posted"] === true, lines };
}

/** Собирает строки табличной части под тип документа (счёт/поступление/реализация). */
async function buildSectionRows(
  conn: Connection,
  entitySet: string,
  orgKey: string,
  lines: GoodsLine[],
): Promise<{ rows: Array<Record<string, unknown>>; total: number } | undefined> {
  const inList = (arr: readonly string[]): boolean => arr.includes(entitySet);
  if (inList(DOCUMENTS.customerInvoice)) {
    const rows = buildInvoiceRows(lines);
    return { rows, total: rowsTotal(rows) };
  }
  if (inList(DOCUMENTS.purchases) || inList(DOCUMENTS.sales)) {
    const kind = inList(DOCUMENTS.purchases) ? "purchase" : "shipment";
    const { lineAccountsFor } = await goodsAccounts(conn, orgKey, lines, kind);
    const rows = buildGoodsRows(lines, lineAccountsFor);
    return { rows, total: rowsTotal(rows) };
  }
  return undefined;
}

/** Банковский счёт организации (для документов оплаты). По названию или первый. */
async function resolveOrgBankAccount(
  conn: Connection,
  orgKey: string,
  query: string | undefined,
): Promise<string | undefined> {
  const set = resolveEntity(CATALOGS.bankAccounts, await conn.available());
  if (!set) return undefined;
  // Владелец банковского счёта — полиморфная ссылка Owner (+ Owner_Type), не Owner_Key.
  const filter = and(
    cmp("Owner", "eq", odataGuid(orgKey)),
    query ? contains("Description", query) : undefined,
  );
  try {
    const { rows } = await fetchAll(conn.client, set, { filter, select: ["Ref_Key"] }, 5, 5);
    return rows[0] ? String(rows[0]["Ref_Key"]) : undefined;
  } catch {
    // Не удалось подобрать — не критично: 1С подставит счёт организации по умолчанию.
    return undefined;
  }
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
  notes?: string[],
) {
  const extra = notes?.length ? { notes } : {};
  if (!confirm) {
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      writableBase: conn.cfg.writable,
      willCreate: entitySet,
      payload,
      ...extra,
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
    ...extra,
  });
}

/** Общий путь изменения (PATCH): dry-run при confirm=false, иначе применяет. */
async function patchOrPreview(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
  notes?: string[],
) {
  const guid = ref.replace(/[{}']/g, "");
  const extra = notes?.length ? { notes } : {};
  if (!confirm) {
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      willPatch: `${entitySet}(${guid})`,
      fields,
      ...extra,
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
    ...extra,
  });
}

/**
 * Создаёт подчинённый объект (банковский счёт / контактное лицо) и, при makeMain,
 * проставляет его основным у владельца. dry-run при confirm=false.
 */
async function createSubordinate(
  conn: Connection,
  set: string,
  payload: Record<string, unknown>,
  confirm: boolean,
  main?: { ownerSet: string; ownerRef: string; field: string },
) {
  if (!confirm) return createOrPreview(conn, set, payload, false);
  const created = await conn.client.create<ODataEntity>(set, payload);
  const ref = String(created["Ref_Key"] ?? "");
  if (main && ref) {
    const g = main.ownerRef.replace(/[{}']/g, "");
    await conn.client.patch(`${main.ownerSet}(guid'${g}')?$format=json`, { [main.field]: ref });
  }
  return ok({
    created: true,
    database: conn.cfg.name,
    entitySet: set,
    ref,
    description: created["Description"],
    ...(main ? { setAsMain: true } : {}),
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
        phone: z.string().optional().describe("Телефон (в контактную информацию)"),
        email: z.string().optional().describe("Email (в контактную информацию)"),
        address: z.string().optional().describe("Адрес текстом (в контактную информацию, юр. адрес)"),
        ogrn: z.string().optional().describe("ОГРН (пишется, если в базе настроен доп.реквизит «ОГРН»)"),
        confirm: confirmField,
      },
    },
    ({ database, name, inn, kpp, fullName, legalType, phone, email, address, ogrn, confirm }) =>
      guard("create_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Контрагенты");
        const { fields, notes } = await counterpartyExtras(conn, { phone, email, address, ogrn });
        const payload = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
          ЮридическоеФизическоеЛицо: legalType,
          ...fields,
        });
        return createOrPreview(conn, set, payload, confirm, notes);
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
        kind: z.enum(CONTRACT_KINDS).default("СПокупателем").describe("Вид договора"),
        currency: z
          .string()
          .optional()
          .describe("Валюта взаиморасчётов — код (напр. 643/840) или название (RUB/USD)"),
        priceType: z
          .string()
          .optional()
          .describe("Тип цен — название или код (из справочника Типы цен номенклатуры)"),
        confirm: confirmField,
      },
    },
    ({ database, organization, counterpartyRef, name, kind, currency, priceType, confirm }) =>
      guard("create_contract", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.contracts, "Договоры контрагентов");
        const org = await resolveOrg(conn, organization);

        let cur: { ref: string; code?: string } | undefined;
        if (currency)
          cur = await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", currency);
        const pt = priceType
          ? await resolveCatalogItem(
              conn,
              CATALOGS.priceTypes,
              "Справочник «Типы цен номенклатуры»",
              priceType,
            )
          : undefined;
        const foreign = cur ? cur.code !== "643" : undefined; // 643 = рубль

        const payload = clean({
          Description: name,
          Owner_Key: counterpartyRef,
          ВидДоговора: kind,
          Организация_Key: org.key,
          ВалютаВзаиморасчетов_Key: cur?.ref,
          ТипЦен_Key: pt?.ref,
          ...(foreign !== undefined ? { Валютный: foreign } : {}),
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
        const set = await requireEntity(
          conn,
          DOCUMENTS.customerInvoice,
          "Документ «Счёт на оплату покупателю»",
        );
        const org = await resolveOrg(conn, organization);
        const rows = buildInvoiceRows(lines);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
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
        '(узнать через describe_entity), напр. {"ИНН":"7701234567"}. По умолчанию предпросмотр (dry-run); ' +
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
        phone: z.string().optional().describe("Телефон (контактная информация)"),
        email: z.string().optional().describe("Email (контактная информация)"),
        address: z.string().optional().describe("Адрес текстом (юр. адрес)"),
        ogrn: z.string().optional().describe("ОГРН (если в базе настроен доп.реквизит)"),
        confirm: confirmField,
      },
    },
    ({ database, ref, name, inn, kpp, fullName, phone, email, address, ogrn, confirm }) =>
      guard("update_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Справочник «Контрагенты»");
        const fields: Record<string, unknown> = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
        });
        const notes: string[] = [];
        // Контакты — со слиянием: заменяем только заданные виды, прочие сохраняем.
        if (phone || email || address) {
          const kinds = await contactKindsForCounterparties(conn);
          const guid = ref.replace(/[{}']/g, "");
          const existing =
            ((await conn.client.getEntity(`${set}(guid'${guid}')${buildQuery({})}`))[
              "КонтактнаяИнформация"
            ] as Array<Record<string, unknown>>) ?? [];
          const replacedTypes = new Set(
            [phone && kinds.phone?.тип, email && kinds.email?.тип, address && kinds.address?.тип].filter(
              Boolean,
            ),
          );
          const kept = existing.filter((r) => !replacedTypes.has(String(r["Тип"])));
          const merged = [...kept, ...buildContactRows(kinds, { phone, email, address })].map((r, i) => ({
            ...r,
            LineNumber: i + 1,
          }));
          fields["КонтактнаяИнформация"] = merged;
        }
        if (ogrn) {
          const prop = await resolveAdditionalProperty(conn, "ОГРН");
          if (prop)
            fields["ДополнительныеРеквизиты"] = [{ LineNumber: 1, Свойство_Key: prop, Значение: ogrn }];
          else notes.push("ОГРН не записан: в базе нет доп.реквизита «ОГРН» для контрагентов.");
        }
        if (Object.keys(fields).length === 0) return fail("Не задано ни одного поля для изменения.");
        return patchOrPreview(conn, set, ref, fields, confirm, notes);
      }),
  );

  server.registerTool(
    "update_nomenclature",
    {
      title: "Изменить номенклатуру",
      description:
        "Изменяет реквизиты номенклатуры по Ref_Key (наименование/полное имя/артикул). " +
        "По умолчанию предпросмотр (dry-run); применение — при confirm=true. Указывайте только меняемые поля.",
      inputSchema: {
        database: databaseField,
        ref: z.string().describe("Ref_Key номенклатуры"),
        name: z.string().optional().describe("Новое наименование (Description)"),
        fullName: z.string().optional().describe("Полное наименование"),
        article: z.string().optional().describe("Артикул"),
        confirm: confirmField,
      },
    },
    ({ database, ref, name, fullName, article, confirm }) =>
      guard("update_nomenclature", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.nomenclature, "Справочник «Номенклатура»");
        const fields = clean({ Description: name, НаименованиеПолное: fullName, Артикул: article });
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
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      sumIncludesVat,
      lines,
      confirm,
    }) =>
      guard("create_purchase", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.purchases, "Документ «Поступление товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "purchase");
        return createGoodsDoc(
          conn,
          set,
          {
            orgKey: org.key,
            counterpartyRef,
            contractRef,
            warehouseKey,
            date,
            sumIncludesVat,
            lines,
            settlement,
            lineAccountsFor,
          },
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
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      sumIncludesVat,
      lines,
      confirm,
    }) =>
      guard("create_shipment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "shipment");
        return createGoodsDoc(
          conn,
          set,
          {
            orgKey: org.key,
            counterpartyRef,
            contractRef,
            warehouseKey,
            date,
            sumIncludesVat,
            lines,
            settlement,
            lineAccountsFor,
          },
          confirm,
        );
      }),
  );

  server.registerTool(
    "update_document_lines",
    {
      title: "Изменить строки документа",
      description:
        "Заменяет табличную часть «Товары» существующего документа новым набором строк и пересчитывает сумму. " +
        "Поддержаны: счёт покупателю, поступление, реализация. Документ должен быть НЕПРОВЕДЁННЫМ " +
        "(если проведён — сначала отмените проведение через post_document). По умолчанию dry-run; применение — при confirm=true.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа"),
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
          .describe("Новый полный набор строк (заменяет прежние)"),
        confirm: confirmField,
      },
    },
    ({ database, entitySet, ref, lines, confirm }) =>
      guard("update_document_lines", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) {
          return fail(
            "Документ проведён. Сначала отмените проведение (post_document с post=false), затем меняйте строки.",
          );
        }
        const built = await buildSectionRows(conn, entitySet, info.orgKey, lines);
        if (!built) {
          return fail(
            "Поддерживаются: счёт покупателю, поступление товаров и услуг, реализация товаров и услуг.",
          );
        }
        return patchOrPreview(
          conn,
          entitySet,
          ref,
          { Товары: built.rows, СуммаДокумента: built.total },
          confirm,
        );
      }),
  );

  const lineObject = z.object({
    nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
    quantity: z.number().positive().describe("Количество"),
    price: z.number().nonnegative().describe("Цена за единицу"),
    vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
  });

  server.registerTool(
    "add_document_line",
    {
      title: "Добавить строку в документ",
      description:
        "Добавляет одну позицию в табличную часть «Товары» существующего НЕПРОВЕДЁННОГО документа " +
        "(счёт/поступление/реализация), сохраняя прежние строки. dry-run/confirm.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string(),
        ref: z.string(),
        line: lineObject,
        confirm: confirmField,
      },
    },
    ({ database, entitySet, ref, line, confirm }) =>
      guard("add_document_line", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) return fail("Документ проведён. Сначала отмените проведение, затем меняйте строки.");
        const built = await buildSectionRows(conn, entitySet, info.orgKey, [...info.lines, line]);
        if (!built) return fail("Поддерживаются: счёт, поступление, реализация.");
        return patchOrPreview(
          conn,
          entitySet,
          ref,
          { Товары: built.rows, СуммаДокумента: built.total },
          confirm,
        );
      }),
  );

  server.registerTool(
    "remove_document_line",
    {
      title: "Удалить строку документа",
      description:
        "Удаляет строку (по номеру LineNumber, начиная с 1) из табличной части «Товары» существующего " +
        "НЕПРОВЕДЁННОГО документа. Нельзя удалить последнюю строку. dry-run/confirm.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string(),
        ref: z.string(),
        lineNumber: z.number().int().positive().describe("Номер строки (с 1)"),
        confirm: confirmField,
      },
    },
    ({ database, entitySet, ref, lineNumber, confirm }) =>
      guard("remove_document_line", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) return fail("Документ проведён. Сначала отмените проведение, затем меняйте строки.");
        if (lineNumber > info.lines.length) return fail(`В документе всего ${info.lines.length} строк(и).`);
        const kept = info.lines.filter((_, i) => i + 1 !== lineNumber);
        if (kept.length === 0)
          return fail("Нельзя удалить последнюю строку — в документе должна остаться хотя бы одна позиция.");
        const built = await buildSectionRows(conn, entitySet, info.orgKey, kept);
        if (!built) return fail("Поддерживаются: счёт, поступление, реализация.");
        return patchOrPreview(
          conn,
          entitySet,
          ref,
          { Товары: built.rows, СуммаДокумента: built.total },
          confirm,
        );
      }),
  );

  server.registerTool(
    "create_act",
    {
      title: "Создать акт (реализация услуг)",
      description:
        "Создаёт «Реализация (акт, накладная)» с табличной частью УСЛУГИ (без склада/остатков). " +
        "Документ НЕПРОВЕДЁННЫЙ; проводки при проведении Дт 62 Кт 90.01 (без 41). " +
        "dry-run/confirm. Контрагент — покупатель; позиции — услуги-номенклатура.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        lines: z.array(lineObject).min(1).describe("Позиции-услуги"),
        confirm: confirmField,
      },
    },
    ({ database, organization, counterpartyRef, contractRef, date, sumIncludesVat, lines, confirm }) =>
      guard("create_act", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "service");
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: "Услуги", // иначе табличная часть «Услуги» не сохраняется
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Услуги: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "create_payment",
    {
      title: "Создать оплату от покупателя",
      description:
        "Создаёт «Поступление на расчётный счёт» (оплата от покупателя) на сумму, по контрагенту и договору. " +
        "Документ НЕПРОВЕДЁННЫЙ; при проведении проводка Дт 51 Кт 62. " +
        "Банковский счёт организации берётся по названию или первый. dry-run/confirm.",
      inputSchema: {
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().describe("Ref_Key договора"),
        amount: z.number().positive().describe("Сумма оплаты"),
        bankAccount: z
          .string()
          .optional()
          .describe("Название банковского счёта организации (если несколько)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      },
    },
    ({ database, organization, counterpartyRef, contractRef, amount, bankAccount, date, confirm }) =>
      guard("create_payment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.bankIn, "Документ «Поступление на расчётный счёт»");
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, bankAccount);
        const codes = await accountsByCode(conn, ["62.01", "62", "62.02"]);
        const settle = pickAccount(codes, "62.01", "62");
        const advance = pickAccount(codes, "62.02");
        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: "ОплатаПокупателя",
          Организация_Key: org.key,
          СчетОрганизации_Key: bank,
          Контрагент: counterpartyRef,
          ...(cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {}),
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          СчетУчетаРасчетовСКонтрагентом_Key: settle,
          РасшифровкаПлатежа: [
            clean({
              LineNumber: 1,
              ДоговорКонтрагента_Key: contractRef,
              СпособПогашенияЗадолженности: "Автоматически", // иначе платёж не распределяется и нет проводок
              СуммаПлатежа: amount,
              СтавкаНДС: "БезНДС",
              СчетУчетаРасчетовСКонтрагентом_Key: settle,
              СчетУчетаРасчетовПоАвансам_Key: advance,
            }),
          ],
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "create_bank_account",
    {
      title: "Создать банковский счёт контрагента",
      description:
        "Заводит расчётный счёт контрагента (подчинённый справочник): банк по БИК + номер счёта. " +
        "Можно сделать счёт основным. По умолчанию dry-run; создание — при confirm=true. " +
        "Ref контрагента — из find_counterparty.",
      inputSchema: {
        database: databaseField,
        ownerRef: z.string().describe("Ref_Key контрагента-владельца счёта"),
        accountNumber: z.string().min(1).describe("Номер расчётного счёта"),
        bik: z.string().min(1).describe("БИК банка (ищется в справочнике «Банки»)"),
        currency: z.string().optional().describe("Валюта (код/название; по умолчанию рубль)"),
        label: z.string().optional().describe("Наименование счёта (по умолчанию — номер счёта)"),
        makeMain: z.boolean().default(false).describe("Сделать основным банковским счётом контрагента"),
        confirm: confirmField,
      },
    },
    ({ database, ownerRef, accountNumber, bik, currency, label, makeMain, confirm }) =>
      guard("create_bank_account", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, CATALOGS.bankAccounts, "Справочник «Банковские счета»");
        const bank = await resolveBankByBik(conn, bik);
        const cur = currency
          ? await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", currency)
          : await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", "643");
        const payload = clean({
          Description: label ?? accountNumber,
          Owner: ownerRef,
          Owner_Type: COUNTERPARTY_TYPE,
          НомерСчета: accountNumber,
          Банк_Key: bank.ref,
          ВалютаДенежныхСредств_Key: cur.ref,
        });
        return createSubordinate(
          conn,
          set,
          payload,
          confirm,
          makeMain
            ? {
                ownerSet: await resolveSet(conn, CATALOGS.counterparties, "Контрагенты"),
                ownerRef,
                field: "ОсновнойБанковскийСчет_Key",
              }
            : undefined,
        );
      }),
  );

  server.registerTool(
    "create_contact_person",
    {
      title: "Создать контактное лицо (директор и т.п.)",
      description:
        "Заводит контактное лицо контрагента (директор, бухгалтер, менеджер): ФИО + должность. " +
        "Можно сделать основным контактным лицом. По умолчанию dry-run; создание — при confirm=true.",
      inputSchema: {
        database: databaseField,
        ownerRef: z.string().describe("Ref_Key контрагента"),
        name: z.string().min(1).describe("ФИО (наименование контактного лица)"),
        lastName: z.string().optional().describe("Фамилия"),
        firstName: z.string().optional().describe("Имя"),
        position: z.string().optional().describe("Должность (напр. «Директор»)"),
        makeMain: z.boolean().default(false).describe("Сделать основным контактным лицом контрагента"),
        confirm: confirmField,
      },
    },
    ({ database, ownerRef, name, lastName, firstName, position, makeMain, confirm }) =>
      guard("create_contact_person", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, ["Catalog_КонтактныеЛица"], "Справочник «Контактные лица»");
        const payload = clean({
          Description: name,
          Фамилия: lastName,
          Имя: firstName,
          Должность: position,
          ОбъектВладелец: ownerRef,
          ОбъектВладелец_Type: COUNTERPARTY_TYPE,
        });
        return createSubordinate(
          conn,
          set,
          payload,
          confirm,
          makeMain
            ? {
                ownerSet: await resolveSet(conn, CATALOGS.counterparties, "Контрагенты"),
                ownerRef,
                field: "ОсновноеКонтактноеЛицо_Key",
              }
            : undefined,
        );
      }),
  );
}
