/**
 * Доменные (бизнес-) типы, которые инструменты возвращают наружу.
 * Это «причёсанное» представление поверх сырых сущностей OData —
 * именно его видит модель/пользователь, без GUID-шумихи и технических полей.
 */

export interface Counterparty {
  ref: string; // Ref_Key
  name: string; // Description
  code?: string; // Code
  inn?: string; // ИНН
  kpp?: string; // КПП
  fullName?: string; // НаименованиеПолное
  isFolder?: boolean; // IsFolder (группа справочника)
}

export interface DocumentSummary {
  ref: string;
  type: string; // человекочитаемый тип, напр. "Реализация товаров и услуг"
  entitySet: string; // техническое имя Document_*
  number?: string; // Number
  date?: string; // Date (ISO)
  posted?: boolean; // Posted
  deletionMark?: boolean; // DeletionMark
  organization?: string;
  counterparty?: string;
  amount?: number; // СуммаДокумента
}

/** Строка остатка (товары на складах / денежные средства). */
export interface BalanceRow {
  dimensions: Record<string, string>; // напр. { Номенклатура, Склад }
  measures: Record<string, number>; // напр. { КоличествоBalance, СуммаBalance }
}

/** Строка задолженности (дебиторка/кредиторка) по контрагенту. */
export interface DebtRow {
  counterparty: string;
  account: string; // счёт, напр. "62.01"
  amount: number; // сальдо
  currency?: string;
}

/** Агрегат продаж за период. */
export interface SalesSummary {
  periodFrom: string;
  periodTo: string;
  total: number;
  byCounterparty?: Array<{ counterparty: string; amount: number }>;
  documents?: DocumentSummary[];
}

/** Движение денежных средств за период. */
export interface CashflowSummary {
  periodFrom: string;
  periodTo: string;
  inflow: number;
  outflow: number;
  net: number;
  byAccount?: Array<{ account: string; inflow: number; outflow: number }>;
}
