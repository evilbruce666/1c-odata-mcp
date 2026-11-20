import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { EMPTY_REF } from "../odata/pricing.js";
import { buildQuery } from "../odata/query.js";
import { currentWriteOperationId } from "../odata/write-operation-context.js";
import { fail } from "./_shared.js";
import { createCore } from "./write.js";

/**
 * Контрагент Управления торговлей: юрлицо (Catalog_Контрагенты) всегда принадлежит партнёру (Catalog_Партнеры) —
 * бизнес-сущности, у которой отношения «клиент / поставщик / прочие». Без партнёра контрагента в УТ не выбрать ни в
 * одном документе. Поэтому в УТ два шага: create_partner, затем create_counterparty с partnerRef.
 */

export const PARTNERS = "Catalog_Партнеры";

/** Схема УТ: у контрагента есть партнёр, справочник партнёров опубликован. КА и ERP с Хозрасчётным — тоже сюда. */
export async function tradeCounterpartySchema(
  conn: Connection,
  set = "Catalog_Контрагенты",
): Promise<boolean> {
  const meta = await conn.getMetadata();
  const hasPartner = !!meta.entities.get(set)?.properties.some((p) => p.name === "Партнер_Key");
  return hasPartner && (await conn.available()).has(PARTNERS);
}

/** У контрагента есть поле партнёра, а справочник партнёров не опубликован: писать нельзя — отказ с советом. */
export async function partnerFieldWithoutCatalog(conn: Connection, set: string): Promise<boolean> {
  const meta = await conn.getMetadata();
  const hasPartner = !!meta.entities.get(set)?.properties.some((p) => p.name === "Партнер_Key");
  return hasPartner && !(await conn.available()).has(PARTNERS);
}

export type LegalType = "ЮридическоеЛицо" | "ФизическоеЛицо" | "ИндивидуальныйПредприниматель";
export type Relation = "client" | "supplier" | "other";

/**
 * Вид контрагента в УТ. Значения перечислений в OData не опубликованы — взяты с живой УТ 11.5 (копия, чтение,
 * 20.11.2026): ЮрФизЛицо — ЮрЛицо / ФизЛицо / ИндивидуальныйПредприниматель; ЮридическоеФизическоеЛицо — ЮридическоеЛицо
 * (юрлица) / ФизическоеЛицо (физлица и ИП); партнёр — Компания, а у частного лица — ЧастноеЛицо.
 */
const LEGAL: Record<LegalType, { kind: string; legal: string }> = {
  ЮридическоеЛицо: { kind: "ЮрЛицо", legal: "ЮридическоеЛицо" },
  ИндивидуальныйПредприниматель: {
    kind: "ИндивидуальныйПредприниматель",
    legal: "ФизическоеЛицо",
  },
  ФизическоеЛицо: { kind: "ФизЛицо", legal: "ФизическоеЛицо" },
};

/** Дата регистрации партнёра — день предпросмотра (из журнала), чтобы подтверждение совпало с ним байт в байт. */
export function dayOf(iso: string | undefined): string {
  const d = iso ? new Date(iso) : new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T00:00:00`;
}

export interface TradeCounterpartyInput {
  name: string;
  inn?: string | undefined;
  kpp?: string | undefined;
  fullName?: string | undefined;
  legalType?: LegalType | undefined;
  partnerRef?: string | undefined;
}

/**
 * Контрагент УТ — к существующему партнёру (partnerRef). Один вызов — один объект: партнёра создаёт
 * write.counterparty.create_partner своей операцией. Составная запись двух объектов одним подтверждением давала
 * второго партнёра при потерянном ответе (два раунда ревью) — механизм убран.
 */
export async function createTradeCounterparty(
  conn: Connection,
  set: string,
  input: TradeCounterpartyInput,
  extraFields: Record<string, unknown>,
  confirm: boolean,
  notes: string[],
): Promise<CallToolResult> {
  const partnerRef = input.partnerRef?.replace(/[{}']/g, "");
  if (!partnerRef)
    return fail(
      "В Управлении торговлей контрагент принадлежит партнёру. Укажите partnerRef — Ref_Key партнёра; нового партнёра " +
        "сначала создайте write.counterparty.create_partner. Ничего не создано.",
    );
  const inn = input.inn?.trim();
  const legalType = input.legalType ?? (inn && /^\d{10}$/.test(inn) ? "ЮридическоеЛицо" : undefined);
  if (!legalType)
    return fail(
      "Укажите legalType: ЮридическоеЛицо, ИндивидуальныйПредприниматель или ФизическоеЛицо. По ИНН из 10 цифр это " +
        "юрлицо, а ИНН из 12 цифр бывает и у ИП, и у частного лица. Ничего не создано.",
    );
  const legal = LEGAL[legalType];
  let partner: ODataEntity;
  try {
    partner = await conn.client.getEntity<ODataEntity>(
      `${PARTNERS}(guid'${partnerRef}')${buildQuery({ select: ["Ref_Key", "Description", "DeletionMark"] })}`,
    );
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found")
      return fail(`Партнёр ${partnerRef} не найден в этой базе. Ничего не создано.`);
    throw e;
  }
  if (partner["DeletionMark"] === true)
    return fail(`Партнёр «${String(partner["Description"])}» помечен на удаление. Ничего не создано.`);

  const payload = Object.fromEntries(
    Object.entries({
      Description: input.name,
      НаименованиеПолное: input.fullName ?? input.name,
      ИНН: inn,
      КПП: input.kpp?.trim(),
      ЮрФизЛицо: legal.kind,
      ЮридическоеФизическоеЛицо: legal.legal,
      Партнер_Key: partnerRef,
      ...extraFields,
    }).filter(([, v]) => v !== undefined && v !== ""),
  );
  return createCore(conn, set, payload, confirm, notes, (created) =>
    // Головной контрагент — сам контрагент (так у всех, кроме обособленных подразделений): 1С заполняет его при
    // записи (проверено на копии); не заполнила — сказать, а не молчать.
    !created["ГоловнойКонтрагент_Key"] || created["ГоловнойКонтрагент_Key"] === EMPTY_REF
      ? "1С не заполнила головного контрагента — откройте карточку в 1С и сохраните её."
      : undefined,
  );
}

export interface PartnerInput {
  name: string;
  fullName?: string | undefined;
  privatePerson: boolean;
  relations: Relation[];
}

/** Партнёр УТ — одна запись своей операцией журнала; контрагента к нему — create_counterparty с partnerRef. */
export async function createPartner(
  conn: Connection,
  input: PartnerInput,
  confirm: boolean,
): Promise<CallToolResult> {
  if (!(await conn.available()).has(PARTNERS))
    return fail(
      "Справочник «Партнеры» (Catalog_Партнеры) не опубликован в OData: партнёры есть в Управлении торговлей, КА и " +
        "ERP. Ничего не создано.",
    );
  const relations = new Set(input.relations);
  const payload = {
    Description: input.name,
    НаименованиеПолное: input.fullName ?? input.name,
    ЮрФизЛицо: input.privatePerson ? "ЧастноеЛицо" : "Компания",
    Клиент: relations.has("client"),
    Поставщик: relations.has("supplier"),
    ПрочиеОтношения: relations.has("other"),
    ДатаРегистрации: await previewDay(conn),
  };
  return createCore(
    conn,
    PARTNERS,
    payload,
    confirm,
    ["Контрагента (юрлицо) к партнёру — write.counterparty.create_counterparty с partnerRef = ref партнёра."],
    () => undefined,
  );
}

/** День предпросмотра этой операции: при подтверждении — из журнала, в предпросмотре — сегодня. */
export async function previewDay(conn: Connection): Promise<string> {
  const operationId = currentWriteOperationId();
  const entry = operationId ? await conn.client.operationEntry(operationId) : undefined;
  return dayOf(entry?.createdAt);
}

/** Партнёр по ссылке на контрагента или на самого партнёра (в УТ контактные лица принадлежат партнёру). */
export async function partnerOf(conn: Connection, ref: string): Promise<string | undefined> {
  const guid = ref.replace(/[{}']/g, "");
  try {
    const c = await conn.client.getEntity<ODataEntity>(
      `Catalog_Контрагенты(guid'${guid}')${buildQuery({ select: ["Партнер_Key"] })}`,
    );
    const key = String(c["Партнер_Key"] ?? "");
    if (key && key !== EMPTY_REF) return key;
  } catch (e) {
    if (!(e instanceof ODataError && e.kind === "not_found")) throw e;
  }
  try {
    await conn.client.getEntity<ODataEntity>(
      `${PARTNERS}(guid'${guid}')${buildQuery({ select: ["Ref_Key"] })}`,
    );
    return guid;
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found") return undefined;
    throw e;
  }
}
