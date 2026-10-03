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
    operationId: z.string().uuid().optional(),
    replayed: z.boolean().optional(),
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
    operationId: z.string().uuid().optional(),
    replayed: z.boolean().optional(),
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
    postings: z
      .object({
        count: z.number(),
        debitTotal: z.number(),
        creditTotal: z.number(),
        byCorrespondence: z.array(z.record(z.string(), z.unknown())),
      })
      .optional(),
    postingsNote: z.string().optional(),
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
    writeJournal: z
      .object({
        dir: z.string(),
        writable: z.boolean(),
        uncertainOperations: z.number(),
        uncertainIds: z.array(z.string()).optional(),
        error: z.string().optional(),
        note: z.string(),
      })
      .optional(),
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

/** Сальдо/обороты ОСВ в рублях (Dr/Cr раздельно, развёрнуто по аналитике). */
const turnoverSumsSchema = z
  .object({
    openingDebit: z.number(),
    openingCredit: z.number(),
    debitTurnover: z.number(),
    creditTurnover: z.number(),
    closingDebit: z.number(),
    closingCredit: z.number(),
  })
  .passthrough();

export const getAccountTurnoverResultSchema = z
  .object({
    database: z.string(),
    organization: z.string().optional(),
    account: z.string(),
    period: z.object({ from: z.string(), to: z.string() }),
    openingDebit: z.number(),
    openingCredit: z.number(),
    debitTurnover: z.number(),
    creditTurnover: z.number(),
    closingDebit: z.number(),
    closingCredit: z.number(),
    openingNet: z.number().optional(),
    closingNet: z.number().optional(),
    consistent: z.boolean(),
    accounts: z.array(
      turnoverSumsSchema.extend({ code: z.string(), description: z.string(), ref: z.string() }).passthrough(),
    ),
    note: z.string().optional(),
    scan: scanSchema,
  })
  .passthrough();

const postingDimensionSchema = z
  .object({
    index: z.number(),
    type: z.string().optional(),
    ref: z.string().optional(),
    value: z.string().optional(),
  })
  .passthrough();

/** Сторона проводки (Дт или Кт); null — у стороны нет счёта (забалансовая проводка). */
const postingSideSchema = z
  .object({
    accountCode: z.string(),
    accountName: z.string(),
    accountRef: z.string(),
    dimensions: z.array(postingDimensionSchema).optional(),
    divisionRef: z.string().optional(),
  })
  .passthrough()
  .nullable();

/** get_document_postings: проводки одного регистратора из регистра Хозрасчетный. */
export const getDocumentPostingsResultSchema = z
  .object({
    database: z.string(),
    document: z
      .object({
        entitySet: z.string(),
        ref: z.string(),
        number: z.string().optional(),
        date: z.string().optional(),
        posted: z.boolean().optional(),
        deletionMark: z.boolean().optional(),
        organization: z.string().optional(),
        organizationRef: z.string().optional(),
        operation: z.string().optional(),
        state: z.string().optional(),
        comment: z.string().optional(),
      })
      .passthrough(),
    postingsCount: z.number(),
    debitTotal: z.number(),
    creditTotal: z.number(),
    postings: z.array(
      z
        .object({
          period: z.string(),
          lineNumber: z.number().optional(),
          active: z.boolean().optional(),
          debit: postingSideSchema,
          credit: postingSideSchema,
          amount: z.number(),
          quantityDebit: z.number().optional(),
          quantityCredit: z.number().optional(),
          currencyAmountDebit: z.number().optional(),
          currencyAmountCredit: z.number().optional(),
          organizationRef: z.string().optional(),
          content: z.string().optional(),
        })
        .passthrough(),
    ),
    byCorrespondence: z.array(
      z
        .object({
          debitAccount: z.string().nullable(),
          creditAccount: z.string().nullable(),
          amount: z.number(),
          entries: z.number(),
        })
        .passthrough(),
    ),
    note: z.string().optional(),
    source: z.object({ entitySet: z.string(), filter: z.string() }).passthrough(),
    scan: scanSchema,
  })
  .passthrough();

/** read.audit.get_document_history — хронология документа (только доказуемое). */
const derivedTimestampSchema = z
  .object({
    value: z.string(),
    source: z.string(),
    confidence: z.enum(["verified", "derived", "uncertain"]),
    description: z.string().optional(),
  })
  .passthrough();

export const getDocumentHistoryResultSchema = z
  .object({
    database: z.string(),
    document: z
      .object({
        entitySet: z.string(),
        ref: z.string(),
        number: z.string().optional(),
        posted: z.boolean().optional(),
        deletionMark: z.boolean().optional(),
        organization: z.string().optional(),
        organizationRef: z.string().optional(),
        operation: z.string().optional(),
        state: z.string().optional(),
        dataVersion: z.string().optional(),
      })
      .passthrough(),
    timestamps: z
      .object({
        documentDate: z.string().optional(),
        refCreatedAt: derivedTimestampSchema.optional(),
        // executedAt / modifiedAt через стандартный OData недоступны — в ответе их нет.
      })
      .passthrough(),
    responsible: z
      .object({
        ref: z.string(),
        name: z.string().optional(),
        resolution: z.enum(["resolved", "not_found", "catalog_not_published", "lookup_failed"]),
        source: z.string(),
        detail: z.string().optional(),
      })
      .passthrough()
      .optional(),
    accountingMovements: z
      .object({
        status: z.enum(["exists", "none", "unsupported", "error"]),
        exists: z.boolean().optional(),
        source: z.string().optional(),
        filter: z.string().optional(),
        detail: z.string().optional(),
      })
      .passthrough(),
    evidence: z.array(
      z
        .object({
          sourceEntity: z.string(),
          sourceRef: z.string().optional(),
          sourceField: z.string(),
          timestamp: z.string().optional(),
          description: z.string(),
          details: z.record(z.string(), z.unknown()).optional(),
        })
        .passthrough(),
    ),
    limitations: z.array(z.string()),
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

/** write.operation.status: состояние операции создания и результат сверки с 1С. */
export const operationStatusSchema = z
  .object({
    operationId: z.string().uuid(),
    database: z.string().optional(),
    status: z.enum([
      "not_in_journal",
      "prepared",
      "succeeded",
      "rejected",
      "found_reconciled",
      "not_found",
      "not_applied",
      "unverifiable",
    ]),
    entitySet: z.string().optional(),
    ref: z.string().optional(),
    number: z.string().optional(),
    note: z.string().optional(),
  })
  .passthrough();
