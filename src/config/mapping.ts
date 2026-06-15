/**
 * Маппинг бизнес-понятий на объекты 1С:Бухгалтерия предприятия 3.0.
 *
 * Это «ожидаемые» имена. Реальные имена объектов берутся из $metadata
 * при старте; если объект из этого списка не найден в базе, инструмент
 * сообщит об этом понятно, а не упадёт. Кандидаты перечислены по
 * убыванию вероятности — выбирается первый, присутствующий в карте.
 */

export const CATALOGS = {
  counterparties: ["Catalog_Контрагенты"],
  nomenclature: ["Catalog_Номенклатура"],
  organizations: ["Catalog_Организации"],
  contracts: ["Catalog_ДоговорыКонтрагентов", "Catalog_Договоры"],
  warehouses: ["Catalog_Склады", "Catalog_СкладыМеста"],
  bankAccounts: ["Catalog_БанковскиеСчета"],
} as const;

export const DOCUMENTS = {
  sales: ["Document_РеализацияТоваровУслуг", "Document_РеализацияТоваровИУслуг"],
  customerInvoice: ["Document_СчетНаОплатуПокупателю", "Document_СчетПокупателю"],
  purchases: ["Document_ПоступлениеТоваровУслуг", "Document_ПоступлениеТоваровИУслуг"],
  bankIn: ["Document_ПоступлениеНаРасчетныйСчет"],
  bankOut: ["Document_СписаниеСРасчетногоСчета"],
  paymentOrder: ["Document_ПлатежноеПоручение"],
  cashIn: ["Document_ПриходныйКассовыйОрдер"],
  cashOut: ["Document_РасходныйКассовыйОрдер"],
} as const;

export const REGISTERS = {
  /** Регистр бухгалтерии (главная книга) — сальдо/обороты по счетам. */
  accounting: ["AccountingRegister_Хозрасчетный"],
  /** Остатки товаров по складам. */
  stock: ["AccumulationRegister_ТоварыНаСкладах", "AccumulationRegister_ТоварыОрганизаций"],
} as const;

/** Счета плана счетов БП 3.0, нужные для аналитики. */
export const ACCOUNTS = {
  receivables: ["62.01", "62.02", "62"], // дебиторка покупателей
  payables: ["60.01", "60.02", "60"], // кредиторка поставщиков
  revenue: ["90.01", "90.01.1"], // выручка
  bank: ["51"], // расчётные счета
  cash: ["50", "50.01"], // касса
} as const;

/** Общие поля документов БП 3.0 (для $select и нормализации). */
export const DOC_FIELDS = {
  ref: "Ref_Key",
  number: "Number",
  date: "Date",
  posted: "Posted",
  deletionMark: "DeletionMark",
  amount: "СуммаДокумента",
  counterparty: "Контрагент_Key",
  organization: "Организация_Key",
} as const;

/** Поля справочника контрагентов. */
export const COUNTERPARTY_FIELDS = {
  ref: "Ref_Key",
  name: "Description",
  code: "Code",
  inn: "ИНН",
  kpp: "КПП",
  fullName: "НаименованиеПолное",
  isFolder: "IsFolder",
} as const;

/** Возвращает первое имя-кандидата, присутствующее в карте сущностей. */
export function resolveEntity(
  candidates: readonly string[],
  available: ReadonlySet<string>,
): string | undefined {
  return candidates.find((name) => available.has(name));
}
