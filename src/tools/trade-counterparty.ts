import { createHash } from "node:crypto";
import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { EMPTY_REF } from "../odata/pricing.js";
import { buildQuery } from "../odata/query.js";
import {
  currentWriteOperationId,
  currentWriteRequestHash,
  withWriteOperation,
} from "../odata/write-operation-context.js";
import { fail } from "./_shared.js";
import { createCore } from "./write.js";

/**
 * Контрагент Управления торговлей: юрлицо (Catalog_Контрагенты) всегда принадлежит партнёру (Catalog_Партнеры) —
 * бизнес-сущности, у которой отношения «клиент / поставщик / прочие». Без партнёра контрагента в УТ не выбрать ни в
 * одном документе. Поэтому create_counterparty в УТ создаёт оба объекта или берёт готового партнёра (partnerRef).
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
const LEGAL: Record<LegalType, { kind: string; legal: string; partner: string }> = {
  ЮридическоеЛицо: { kind: "ЮрЛицо", legal: "ЮридическоеЛицо", partner: "Компания" },
  ИндивидуальныйПредприниматель: {
    kind: "ИндивидуальныйПредприниматель",
    legal: "ФизическоеЛицо",
    partner: "Компания",
  },
  ФизическоеЛицо: { kind: "ФизЛицо", legal: "ФизическоеЛицо", partner: "ЧастноеЛицо" },
};

/** Подоперация журнала для партнёра: детерминированный UUID от operationId контрагента. */
export function derivedOperationId(operationId: string, part: string): string {
  const h = createHash("sha256").update(`${operationId}:${part}`).digest("hex");
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-5${h.slice(13, 16)}-8${h.slice(17, 20)}-${h.slice(20, 32)}`;
}

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
  relations?: Relation[] | undefined;
}

export async function createTradeCounterparty(
  conn: Connection,
  set: string,
  input: TradeCounterpartyInput,
  extraFields: Record<string, unknown>,
  confirm: boolean,
  notes: string[],
): Promise<CallToolResult> {
  const inn = input.inn?.trim();
  const legalType = input.legalType ?? (inn && /^\d{10}$/.test(inn) ? "ЮридическоеЛицо" : undefined);
  if (!legalType)
    return fail(
      "Укажите legalType: ЮридическоеЛицо, ИндивидуальныйПредприниматель или ФизическоеЛицо. По ИНН из 10 цифр это " +
        "юрлицо, а ИНН из 12 цифр бывает и у ИП, и у частного лица. Ничего не создано.",
    );
  const legal = LEGAL[legalType];

  let partnerKey: string | undefined;
  let partnerPreview: Record<string, unknown> | undefined;
  const partnerRef = input.partnerRef?.replace(/[{}']/g, "");
  if (partnerRef) {
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
    partnerKey = partnerRef;
    if (input.relations?.length)
      notes.push("relations не применены: отношения существующего партнёра этим инструментом не меняются.");
  } else if (!input.relations?.length) {
    return fail(
      "В Управлении торговлей контрагент принадлежит партнёру. Укажите relations (client — клиент, supplier — " +
        "поставщик, other — прочие отношения), чтобы создать нового партнёра, или partnerRef — существующего. " +
        "Ничего не создано.",
    );
  }

  const operationId = currentWriteOperationId();
  const requestHash = currentWriteRequestHash();
  const partnerOp = operationId ? derivedOperationId(operationId, "partner") : undefined;
  const partnerPayload = async () => {
    const entry = partnerOp ? await conn.client.operationEntry(partnerOp) : undefined;
    const relations = new Set(input.relations);
    return {
      Description: input.name,
      НаименованиеПолное: input.fullName ?? input.name,
      ЮрФизЛицо: legal.partner,
      Клиент: relations.has("client"),
      Поставщик: relations.has("supplier"),
      ПрочиеОтношения: relations.has("other"),
      ДатаРегистрации: dayOf(entry?.createdAt),
    };
  };
  const inPartnerOp = <T>(fn: () => Promise<T>) =>
    partnerOp && requestHash ? withWriteOperation(partnerOp, requestHash, fn) : fn();

  if (!partnerKey && !confirm) {
    partnerPreview = await partnerPayload();
    await inPartnerOp(() => conn.client.prepareCreate(PARTNERS, partnerPreview!));
    partnerKey = (partnerOp && (await conn.client.operationEntry(partnerOp))?.refKey) || EMPTY_REF;
    notes.push(
      "Будет создан и партнёр (Catalog_Партнеры) — см. partner. Контрагент без партнёра в документах УТ не выбирается.",
    );
  }
  let createdPartner: string | undefined;
  if (!partnerKey && confirm) {
    const payload = await partnerPayload();
    const partner = await inPartnerOp(() => conn.client.create<ODataEntity>(PARTNERS, payload));
    createdPartner = String(partner["Ref_Key"]);
    partnerKey = createdPartner;
  }

  const payload = Object.fromEntries(
    Object.entries({
      Description: input.name,
      НаименованиеПолное: input.fullName ?? input.name,
      ИНН: inn,
      КПП: input.kpp?.trim(),
      ЮрФизЛицо: legal.kind,
      ЮридическоеФизическоеЛицо: legal.legal,
      Партнер_Key: partnerKey,
      ...extraFields,
    }).filter(([, v]) => v !== undefined && v !== ""),
  );
  let result: CallToolResult;
  try {
    result = await createCore(conn, set, payload, confirm, notes, (created) =>
      // Головной контрагент — сам контрагент (так у всех, кроме обособленных подразделений): 1С заполняет его при
      // записи; не заполнила — сказать, а не молчать.
      !created["ГоловнойКонтрагент_Key"] || created["ГоловнойКонтрагент_Key"] === EMPTY_REF
        ? "1С не заполнила головного контрагента — откройте карточку в 1С и сохраните её."
        : undefined,
    );
  } catch (e) {
    if (!createdPartner) throw e;
    const message = e instanceof Error ? e.message : String(e);
    return fail(
      `Партнёр создан (ref ${createdPartner}), а контрагент — нет: ${message} Повторите подтверждение с тем же ` +
        `operationId, если исход неизвестен; если 1С отказала — новый предпросмотр с partnerRef=${createdPartner}: ` +
        "второй партнёр не создастся.",
    );
  }
  if (result.isError || !result.structuredContent) return result;
  const sc = result.structuredContent as Record<string, unknown>;
  const extra = partnerPreview
    ? { partner: { willCreate: PARTNERS, ref: partnerKey, payload: partnerPreview } }
    : { partnerRef: partnerKey, ...(createdPartner ? { partnerCreated: true } : {}) };
  const structuredContent = { ...sc, ...extra };
  return {
    ...result,
    structuredContent,
    content: [{ type: "text", text: JSON.stringify(structuredContent, null, 2) }],
  };
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
