import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { nextDay, num } from "./accounting.js";
import { requireEntity } from "./publication.js";
import { refOf } from "./settlements-aging.js";

/**
 * Расчёты с партнёрами в Управлении торговлей: регистры «Расчёты с клиентами / поставщиками по срокам». Это
 * управленческие расчёты по партнёрам и объектам расчётов, а не счёт 62/60 по договорам: в КА и ERP они могут
 * расходиться с бухучётом, поэтому инструмент отдельный от get_debtors / get_creditors.
 *
 * Регистр «по срокам» выбран вместо основного: в основном сумма только в валюте расчётов, а знак у поставщиков
 * обратный (−638 млн при долге 638 млн на базе УТ 11.5). «По срокам» даёт долг и предоплату отдельно, в
 * регламентированной валюте, с плановой датой погашения; итог по каждому ключу партнёра совпал с основным регистром.
 */

export type Side = "clients" | "suppliers";

const REGISTERS: Record<Side, { set: string; label: string }> = {
  clients: {
    set: "AccumulationRegister_РасчетыСКлиентамиПоСрокам",
    label: "Регистр накопления «Расчёты с клиентами по срокам»",
  },
  suppliers: {
    set: "AccumulationRegister_РасчетыСПоставщикамиПоСрокам",
    label: "Регистр накопления «Расчёты с поставщиками по срокам»",
  },
};

const KEY_NAV = "АналитикаУчетаПоПартнерам";
export const PARTNERS_CATALOG = "Catalog_Партнеры";

export interface SettlementRow {
  partnerRef?: string;
  organizationRef?: string;
  /** Плановая дата погашения YYYY-MM-DD; undefined — не задана. */
  dueDate?: string;
  debtCents: number;
  prepaymentCents: number;
}

/** Строки остатков «по срокам» с партнёром и организацией из ключа аналитики (тем же запросом, через $expand). */
/** Регистр стороны расчётов; не опубликован — отказ с его названием. */
export const requireSettlements = (conn: Connection, side: Side): Promise<string> =>
  requireEntity(conn, [REGISTERS[side].set], REGISTERS[side].label);

export async function readSettlements(
  conn: Connection,
  side: Side,
  asOf?: string,
): Promise<{ rows: SettlementRow[]; scanned: number }> {
  const { set, label } = REGISTERS[side];
  const register = await requireEntity(conn, [set], label);
  const path = `${register}/Balance${asOf ? `(Period=datetime'${nextDay(asOf)}T00:00:00')` : ""}`;
  const { rows } = await fetchAllForAggregation(
    conn,
    path,
    {
      select: [
        `${KEY_NAV}_Key`,
        `${KEY_NAV}/Партнер_Key`,
        `${KEY_NAV}/Организация_Key`,
        "ДатаПлановогоПогашения",
        "ДолгРеглBalance",
        "ПредоплатаРеглBalance",
      ],
      expand: [KEY_NAV],
    },
    `расчёты на ${asOf ?? "сейчас"}`,
  );
  return {
    scanned: rows.length,
    rows: rows.map((r) => {
      const key = (r[KEY_NAV] ?? {}) as ODataEntity;
      const partnerRef = refOf(key["Партнер_Key"]);
      const organizationRef = refOf(key["Организация_Key"]);
      const due =
        typeof r["ДатаПлановогоПогашения"] === "string" ? r["ДатаПлановогоПогашения"].slice(0, 10) : "";
      return {
        ...(partnerRef ? { partnerRef } : {}),
        ...(organizationRef ? { organizationRef } : {}),
        ...(due && !due.startsWith("0001") ? { dueDate: due } : {}),
        debtCents: Math.round(num(r["ДолгРеглBalance"]) * 100),
        prepaymentCents: Math.round(num(r["ПредоплатаРеглBalance"]) * 100),
      };
    }),
  };
}

/** Дней между датами YYYY-MM-DD (b − a). */
export function daysBetween(a: string, b: string): number {
  return Math.round((Date.parse(`${b}T00:00:00Z`) - Date.parse(`${a}T00:00:00Z`)) / 86_400_000);
}

export const AGE_BUCKETS = [
  "notDue",
  "days1to30",
  "days31to90",
  "days91to180",
  "over180",
  "noDueDate",
] as const;
export type AgeBucket = (typeof AGE_BUCKETS)[number];

/** Корзина срока долга на дату: в срок, просрочка по интервалам или срок не задан. */
export function ageBucket(dueDate: string | undefined, on: string): AgeBucket {
  if (!dueDate) return "noDueDate";
  const late = daysBetween(dueDate, on);
  if (late <= 0) return "notDue";
  if (late <= 30) return "days1to30";
  if (late <= 90) return "days31to90";
  if (late <= 180) return "days91to180";
  return "over180";
}
