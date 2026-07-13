/**
 * Zod-схемы structuredContent для outputSchema инструментов (см. src/tools/*.ts).
 *
 * Все схемы — единый плоский объект (.passthrough()), НЕ z.union()/дискриминированные
 * варианты: MCP SDK требует outputSchema.type === "object" на верхнем уровне (иначе
 * при вызове ломается с "Cannot read properties of undefined (reading '_zod')" —
 * проверено вживую). Ветки dry-run/confirmed одного инструмента поэтому описаны как
 * одна схема со всеми полями .optional() — каждая ветка использует своё подмножество.
 * .passthrough() — намеренно: лишнее (забытое здесь или добавленное позже) поле не
 * должно ронять вызов инструмента ошибкой валидации.
 */

import { z } from "zod";

/** Справочник «Контрагенты» — единая форма для find_counterparty/get_counterparty. */
export const counterpartySchema = z
  .object({
    ref: z.string(),
    name: z.string(),
    code: z.string().optional(),
    inn: z.string().optional(),
    kpp: z.string().optional(),
    fullName: z.string().optional(),
    isFolder: z.boolean().optional(),
  })
  .passthrough();

/** Сводка по документу (поиск/история взаиморасчётов). */
export const documentSummarySchema = z
  .object({
    ref: z.string(),
    type: z.string(),
    entitySet: z.string(),
    number: z.string().optional(),
    date: z.string().optional(),
    posted: z.boolean().optional(),
    deletionMark: z.boolean().optional(),
    organization: z.string().optional(),
    counterparty: z.string().optional(),
    amount: z.number().optional(),
  })
  .passthrough();

/** Обёртка усечённого списка (withTruncationNote в _shared.ts). */
export function truncatedList<T extends z.ZodTypeAny>(item: T) {
  return z
    .object({
      rows: z.array(item),
      count: z.number(),
      truncated: z.boolean(),
      note: z.string().optional(),
    })
    .passthrough();
}

/**
 * Динамическая сущность OData (get_document) — форма зависит от типа документа.
 * НЕ голый z.record(): такой schema на верхнем уровне outputSchema тоже ломает SDK
 * так же, как z.union() — "Cannot read properties of undefined (reading '_zod')"
 * (проверено вживую). z.object({}).passthrough() даёт корректный type:"object".
 */
export const odataEntitySchema = z.object({}).passthrough();

// ─── Запись: общие формы (createOrPreview/createSubordinate, patchOrPreview,
//     mark_for_deletion, post_document) — см. docs плана, все поля optional ───

/** create_* через createOrPreview()/createSubordinate(): dry-run ИЛИ created. */
export const createResultSchema = z
  .object({
    dryRun: z.boolean().optional(),
    database: z.string().optional(),
    writableBase: z.boolean().optional(),
    willCreate: z.string().optional(),
    payload: z.record(z.string(), z.unknown()).optional(),
    created: z.boolean().optional(),
    entitySet: z.string().optional(),
    ref: z.string().optional(),
    code: z.string().optional(),
    description: z.string().optional(),
    notes: z.array(z.string()).optional(),
    note: z.string().optional(),
  })
  .passthrough();

/** update_* через patchOrPreview(): dry-run ИЛИ updated. */
export const patchResultSchema = z
  .object({
    dryRun: z.boolean().optional(),
    database: z.string().optional(),
    willPatch: z.string().optional(),
    fields: z.record(z.string(), z.unknown()).optional(),
    updated: z.boolean().optional(),
    entitySet: z.string().optional(),
    ref: z.string().optional(),
    description: z.string().optional(),
    notes: z.array(z.string()).optional(),
    note: z.string().optional(),
  })
  .passthrough();

/** write.entity.mark_for_deletion: свой вариант (без entitySet/description, есть deletionMark). */
export const markForDeletionResultSchema = z
  .object({
    dryRun: z.boolean().optional(),
    database: z.string().optional(),
    willPatch: z.string().optional(),
    payload: z.record(z.string(), z.unknown()).optional(),
    note: z.string().optional(),
    updated: z.boolean().optional(),
    ref: z.string().optional(),
    deletionMark: z.boolean().optional(),
  })
  .passthrough();

/** write.document.post_document: свой вариант (willCall/done/action). */
export const postDocumentResultSchema = z
  .object({
    dryRun: z.boolean().optional(),
    database: z.string().optional(),
    willCall: z.string().optional(),
    note: z.string().optional(),
    done: z.boolean().optional(),
    ref: z.string().optional(),
    action: z.string().optional(),
  })
  .passthrough();

// ─── Чтение: по одной bespoke-схеме на инструмент ───

export const listDatabasesResultSchema = z
  .object({
    default: z.string().optional(),
    databases: z.array(
      z.object({ name: z.string(), label: z.string().optional(), isDefault: z.boolean() }).passthrough(),
    ),
  })
  .passthrough();

export const healthCheckResultSchema = z
  .object({
    status: z.string(),
    database: z.string(),
    label: z.string().optional(),
    odataVersion: z.string().optional(),
    entityCount: z.number(),
    baseUrl: z.string(),
    readOnly: z.boolean(),
  })
  .passthrough();

export const listOrganizationsResultSchema = z
  .object({
    count: z.number(),
    organizations: z.array(
      z.object({ ref: z.string(), name: z.string(), inn: z.string().optional() }).passthrough(),
    ),
  })
  .passthrough();

export const listEntitiesResultSchema = z
  .object({
    database: z.string(),
    odataVersion: z.string().optional(),
    groups: z.record(
      z.string(),
      z.array(z.object({ entitySet: z.string(), name: z.string() }).passthrough()),
    ),
  })
  .passthrough();

export const describeEntityResultSchema = z
  .object({
    entitySet: z.string(),
    class: z.string(),
    keys: z.array(z.string()),
    fields: z.array(z.object({ name: z.string(), type: z.string(), nullable: z.boolean() }).passthrough()),
    relations: z.array(z.string()),
  })
  .passthrough();

/** get_customer_history/get_supplier_history: {counterparty, period} + спред truncatedList. */
export const counterpartyHistoryResultSchema = z
  .object({
    counterparty: z.string(),
    period: z.object({ from: z.string().optional(), to: z.string().optional() }).passthrough(),
    rows: z.array(documentSummarySchema),
    count: z.number(),
    truncated: z.boolean(),
    note: z.string().optional(),
  })
  .passthrough();

/** search_documents: спред truncatedList + опциональная заметка о клиентском фильтре контрагента. */
export const searchDocumentsResultSchema = z
  .object({
    rows: z.array(documentSummarySchema),
    count: z.number(),
    truncated: z.boolean(),
    note: z.string().optional(),
    counterpartyFilter: z.string().optional(),
  })
  .passthrough();

const scanSchema = z
  .object({
    documentsScanned: z.number().optional(),
    rowsScanned: z.number().optional(),
    windows: z.number().optional(),
    elapsedMs: z.number(),
  })
  .passthrough();

export const getSalesResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    period: z.object({ from: z.string(), to: z.string() }),
    total: z.number(),
    byDocument: z.record(z.string(), z.number()),
    scan: scanSchema,
  })
  .passthrough();

export const getCashflowResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    period: z.object({ from: z.string(), to: z.string() }),
    inflow: z.number(),
    outflow: z.number(),
    net: z.number(),
    byDocument: z.record(z.string(), z.number()),
    scan: scanSchema,
  })
  .passthrough();

export const getDebtorsResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    asOf: z.string().optional(),
    accounts: z.array(z.string()),
    totalReceivable: z.number(),
    count: z.number(),
    debtors: z.array(
      z.object({ counterparty: z.string(), ref: z.string(), amount: z.number() }).passthrough(),
    ),
    scan: scanSchema,
  })
  .passthrough();

export const getInventoryResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    asOf: z.string().optional(),
    accounts: z.array(z.string()),
    totalAmount: z.number(),
    count: z.number(),
    items: z.array(
      z.object({ item: z.string(), ref: z.string(), quantity: z.number(), amount: z.number() }).passthrough(),
    ),
    scan: scanSchema,
  })
  .passthrough();

const breakdownGroupSchema = z
  .object({
    label: z.string(),
    count: z.number(),
    sum: z.number(),
    inflow: z.number().optional(),
    outflow: z.number().optional(),
  })
  .passthrough();

/** get_payments_breakdown: приход/расход по виду операции/контрагенту/статье ДДС. */
export const paymentsBreakdownResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    period: z.object({ from: z.string(), to: z.string() }),
    direction: z.string(),
    filters: z.record(z.string(), z.unknown()),
    groupBy: z.string(),
    documents: z.number(),
    inflow: z.number(),
    outflow: z.number(),
    net: z.number(),
    groups: z.array(breakdownGroupSchema),
    scan: scanSchema,
  })
  .passthrough();

/** get_deal_history: хронология по сделке/договору. */
export const dealHistoryResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    period: z.object({ from: z.string().optional(), to: z.string().optional() }).optional(),
    filters: z.record(z.string(), z.unknown()),
    events: z.number(),
    inflow: z.number(),
    outflow: z.number(),
    net: z.number(),
    items: z.array(
      z
        .object({
          date: z.string(),
          entitySet: z.string(),
          number: z.string().optional(),
          direction: z.string(),
          amount: z.number(),
          operation: z.string().optional(),
          counterparty: z.string().optional(),
          purpose: z.string().optional(),
        })
        .passthrough(),
    ),
    note: z.string().optional(),
    scan: scanSchema,
  })
  .passthrough();

/** get_taxes_paid: уплаченные налоги/взносы за период. */
export const taxesPaidResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    period: z.object({ from: z.string(), to: z.string() }),
    filters: z.record(z.string(), z.unknown()),
    groupBy: z.string(),
    documents: z.number(),
    total: z.number(),
    groups: z.array(breakdownGroupSchema),
    scan: scanSchema,
  })
  .passthrough();

/** get_sales_breakdown/get_purchases_breakdown: общий aggregate() в sales.ts. */
export const salesBreakdownResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    period: z.object({ from: z.string(), to: z.string() }),
    filters: z.record(z.string(), z.unknown()),
    groupBy: z.string(),
    entitySet: z.string(),
    documents: z.number(),
    total: z.number(),
    groups: z.array(breakdownGroupSchema),
    scan: scanSchema,
  })
  .passthrough();

/** get_organization_card: реквизиты организации, почти все поля опциональны. */
export const organizationCardResultSchema = z
  .object({
    database: z.string(),
    ref: z.string(),
    name: z.string().optional(),
    fullName: z.string().optional(),
    shortName: z.string().optional(),
    legalType: z.string().optional(),
    inn: z.string().optional(),
    kpp: z.string().optional(),
    ogrn: z.string().optional(),
    registrationDate: z.string().optional(),
    okved: z.object({ code: z.string(), name: z.string().optional() }).passthrough().optional(),
    taxAuthority: z
      .object({ code: z.string().optional(), name: z.string().optional() })
      .passthrough()
      .optional(),
    contacts: z.array(z.object({ kind: z.string(), value: z.string() }).passthrough()).optional(),
    bankAccount: z
      .object({
        accountNumber: z.string().optional(),
        bank: z.object({ name: z.string(), bik: z.string() }).passthrough().optional(),
        currency: z.object({ name: z.string(), code: z.string() }).passthrough().optional(),
      })
      .passthrough()
      .optional(),
    director: z.object({ fullName: z.string() }).passthrough().optional(),
    accountant: z.object({ fullName: z.string() }).passthrough().optional(),
    notes: z.array(z.string()).optional(),
  })
  .passthrough();
