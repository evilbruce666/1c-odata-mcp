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
  currencies: ["Catalog_Валюты"],
  priceTypes: ["Catalog_ТипыЦенНоменклатуры", "Catalog_ТипыЦен"],
  cashflowItems: ["Catalog_СтатьиДвиженияДенежныхСредств"],
  physicalPersons: ["Catalog_ФизическиеЛица"],
  nomenclatureGroups: ["Catalog_НоменклатурныеГруппы"],
} as const;

export const DOCUMENTS = {
  sales: ["Document_РеализацияТоваровУслуг", "Document_РеализацияТоваровИУслуг"],
  customerInvoice: ["Document_СчетНаОплатуПокупателю", "Document_СчетПокупателю"],
  purchases: ["Document_ПоступлениеТоваровУслуг", "Document_ПоступлениеТоваровИУслуг"],
  supplierInvoice: ["Document_СчетНаОплатуПоставщика"],
  bankIn: ["Document_ПоступлениеНаРасчетныйСчет"],
  bankOut: ["Document_СписаниеСРасчетногоСчета"],
  paymentOrder: ["Document_ПлатежноеПоручение"],
  cashIn: ["Document_ПриходныйКассовыйОрдер"],
  cashOut: ["Document_РасходныйКассовыйОрдер"],
  // Счета-фактуры (фаза 3) — создаются на основании реализации/поступления.
  issuedInvoice: ["Document_СчетФактураВыданный"],
  receivedInvoice: ["Document_СчетФактураПолученный"],
  // Акт об оказании услуг (фаза 4) — доходы/расходы по номенклатурной группе.
  servicesAct: ["Document_АктОбОказанииПроизводственныхУслуг"],
  // Товарные складские документы (фаза 1).
  returnFromCustomer: ["Document_ВозвратТоваровОтПокупателя"],
  returnToSupplier: ["Document_ВозвратТоваровПоставщику"],
  transfer: ["Document_ПеремещениеТоваров"],
  surplus: ["Document_ОприходованиеТоваров"],
  writeoff: ["Document_СписаниеТоваров"],
  inventory: ["Document_ИнвентаризацияТоваровНаСкладе"],
} as const;

export const REGISTERS = {
  /** Регистр бухгалтерии (главная книга) — сальдо/обороты по счетам. */
  accounting: ["AccountingRegister_Хозрасчетный"],
  /** Остатки товаров по складам. */
  stock: ["AccumulationRegister_ТоварыНаСкладах", "AccumulationRegister_ТоварыОрганизаций"],
  /** Периодический регистр: директор/гл.бухгалтер/кассир организации (может быть не опубликован). */
  responsiblePersons: [
    "InformationRegister_ОтветственныеЛицаОрганизаций",
    "InformationRegister_ОтветственныеЛицаОрганизации",
  ],
} as const;

/**
 * Префиксы счетов плана счетов БП 3.0 для аналитики через регистр Хозрасчетный.
 * Берётся как startswith(Code, prefix) — покрывает все субсчета группы.
 */
export const ACCOUNT_PREFIX = {
  receivables: ["62"], // расчёты с покупателями (дебиторка = Dr)
  payables: ["60"], // расчёты с поставщиками (кредиторка = Cr)
  inventory: ["41", "10", "43"], // товары / материалы / готовая продукция
  revenue: ["90.01"], // выручка
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
