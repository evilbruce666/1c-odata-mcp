import type { Connection } from "../context.js";
import { accountingRegister } from "../config/accounting.js";
import { resolveEntity } from "../config/mapping.js";

/**
 * Признак базы «торговля без бухучёта»: регистра бухгалтерии нет, а регистры оперативного учёта (склад, продажи)
 * есть — так выглядит Управление торговлей 11. В КА 2 и ERP 2 есть и то и другое, они сюда не попадают: там работают
 * инструменты бухучёта. Признак — по опубликованным объектам, а не по имени конфигурации: OData его не отдаёт.
 */

/** Регистры оперативного учёта, по которым узнаётся торговая конфигурация. Имена — как в stock.ts и item-sales.ts. */
const TRADE_REGISTERS = [
  "AccumulationRegister_ТоварыНаСкладах",
  "AccumulationRegister_ВыручкаИСебестоимостьПродаж",
] as const;

export async function tradeWithoutAccounting(conn: Connection): Promise<boolean> {
  const available = await conn.available();
  return (
    !resolveEntity(accountingRegister(conn.cfg), available) && TRADE_REGISTERS.some((r) => available.has(r))
  );
}

/** Что в такой базе есть вместо бухучёта: подсказка для отказов, ведёт к инструментам, которые там работают. */
export const TRADE_TOOLS_HINT =
  "Для Управления торговлей: остатки на складах — read.analytics.get_warehouse_stock, продажи в штуках и выручка " +
  "по регистру продаж (с розницей) — read.analytics.get_sales_register, долги клиентов и поставщикам — " +
  "read.analytics.get_partner_settlements, деньги на счетах и в кассах — read.analytics.get_cash_balance, документы — " +
  "read.document.search_documents.";

/** Текст отказа инструмента бухучёта в торговой базе. */
export const tradeRefusal = (what: string): string =>
  `${what} В этой базе нет регистра бухгалтерии «Хозрасчётный»: похоже, это Управление торговлей, бухгалтерского ` +
  `учёта в ней нет, и публикация тут не поможет. ${TRADE_TOOLS_HINT}`;
