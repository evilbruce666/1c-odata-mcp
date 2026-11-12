import type { Connection } from "../context.js";
import { resolveEntity } from "../config/mapping.js";
import type { ODataEntity } from "../types/odata.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { nextDay, num, resolveNames } from "./accounting.js";
import { cmp, odataGuid } from "./query.js";
import { refOf } from "./settlements-aging.js";

/**
 * Деньги в Управлении торговлей: остатки регистров денежных средств вместо счетов 50/51/57. Тот же факт, что
 * get_cash_balance в Бухгалтерии, поэтому инструмент один, а источник выбирается по базе. Суммы — в регламентированной
 * валюте (СуммаРегл), как у счетов Хозрасчётного; сумма в валюте счёта — отдельно, если отличается.
 */

interface CashRegister {
  set: string;
  /** Подпись источника в ответе. */
  label: string;
  /** Вид строки: bank, cash, pos (кассы ККМ), transit. */
  kind: "bank" | "cash" | "pos" | "transit";
  /** Измерение-держатель и справочник для названий. */
  holder: string;
  holderCatalog?: string;
}

const REGISTERS: readonly CashRegister[] = [
  {
    set: "AccumulationRegister_ДенежныеСредстваБезналичные",
    label: "Безналичные",
    kind: "bank",
    holder: "БанковскийСчет_Key",
    holderCatalog: "Catalog_БанковскиеСчетаОрганизаций",
  },
  {
    set: "AccumulationRegister_ДенежныеСредстваНаличные",
    label: "Наличные в кассах",
    kind: "cash",
    holder: "Касса_Key",
    holderCatalog: "Catalog_Кассы",
  },
  {
    set: "AccumulationRegister_ДенежныеСредстваВКассахККМ",
    label: "В кассах ККМ",
    kind: "pos",
    holder: "КассаККМ_Key",
    holderCatalog: "Catalog_КассыККМ",
  },
  {
    set: "AccumulationRegister_ДенежныеСредстваВПути",
    label: "В пути",
    kind: "transit",
    holder: "ВидПереводаДенежныхСредств",
  },
];

/** Опубликован ли хоть один регистр денег УТ. */
export async function tradeCashPublished(conn: Connection): Promise<boolean> {
  const available = await conn.available();
  return REGISTERS.some((r) => available.has(r.set));
}

const toCents = (v: unknown): number => Math.round(num(v) * 100);

export async function tradeCashBalance(
  conn: Connection,
  o: {
    organizationRef?: string | undefined;
    organizationName?: string | undefined;
    asOf?: string | undefined;
  },
) {
  const t0 = Date.now();
  const available = await conn.available();
  const used = REGISTERS.filter((r) => available.has(r.set));
  type Line = { reg: CashRegister; holder: string; cents: number; ownCents: number };
  const lines: Line[] = [];
  let scanned = 0;
  for (const reg of used) {
    const params = [
      ...(o.asOf ? [`Period=datetime'${nextDay(o.asOf)}T00:00:00'`] : []),
      ...(o.organizationRef
        ? [`Condition='${cmp("Организация_Key", "eq", odataGuid(o.organizationRef)).replace(/'/g, "''")}'`]
        : []),
    ];
    const path = `${reg.set}/Balance${params.length ? `(${params.join(",")})` : ""}`;
    const { rows } = await fetchAllForAggregation(
      conn,
      path,
      { select: [reg.holder, "СуммаBalance", "СуммаРеглBalance"] },
      `деньги на ${o.asOf ?? "сейчас"}`,
    );
    scanned += rows.length;
    const byHolder = new Map<string, Line>();
    for (const r of rows as ODataEntity[]) {
      const holder = String(r[reg.holder] ?? "");
      const l = byHolder.get(holder) ?? { reg, holder, cents: 0, ownCents: 0 };
      l.cents += toCents(r["СуммаРеглBalance"]);
      l.ownCents += toCents(r["СуммаBalance"]);
      byHolder.set(holder, l);
    }
    lines.push(...[...byHolder.values()].filter((l) => l.cents !== 0 || l.ownCents !== 0));
  }

  const names = new Map<string, string>();
  for (const reg of used) {
    const catalog = reg.holderCatalog && resolveEntity([reg.holderCatalog], available);
    const refs = lines.filter((l) => l.reg === reg).flatMap((l) => refOf(l.holder) ?? []);
    if (catalog && refs.length) for (const [k, v] of await resolveNames(conn, catalog, refs)) names.set(k, v);
  }

  lines.sort((a, b) => used.indexOf(a.reg) - used.indexOf(b.reg) || b.cents - a.cents);
  const sum = (kinds: CashRegister["kind"][]) =>
    lines.filter((l) => kinds.includes(l.reg.kind)).reduce((t, l) => t + l.cents, 0) / 100;
  const negative = lines.filter((l) => l.cents < 0);
  const notes = [
    "Управление торговлей: остатки регистров денежных средств, а не счетов бухучёта. available — расчётные счета, " +
      "кассы и кассы ККМ; inTransit — деньги в пути (эквайринг, переводы между счетами и кассами); special в УТ не ведётся, 0. " +
      "Суммы — в регламентированной валюте; currencyAmount — в валюте счёта или кассы, если она другая.",
  ];
  if (negative.length)
    notes.push(`Отрицательные остатки (${negative.length}) — проверить в учёте: итоги с их учётом.`);
  const transit = sum(["transit"]);
  if (transit < 0) {
    // В регистре «в пути» отправка перевода — приход, получение — расход. Минус значит: получение отражено, а отправка
    // нет (на базе УТ 11.5 так накопилось −484 млн). Цифра регистра верна, но это не деньги в дороге.
    notes.push(
      `Деньги в пути отрицательные (${transit}): переводы получены, а их отправка в регистре не отражена — в учёте не ` +
        "закрыты пары документов (см. строки kind=transit по видам перевода). Остатки касс и счетов-получателей при этом " +
        "могут быть завышены; сверьте с отчётом 1С «Остатки денежных средств».",
    );
  }

  return {
    database: conn.cfg.name,
    ...(o.organizationName ? { organization: o.organizationName } : {}),
    ...(o.asOf ? { asOf: o.asOf } : {}),
    source: "registers",
    accounts: used.map((r) => `${r.label} (${r.set})`),
    available: sum(["bank", "cash", "pos"]),
    special: 0,
    inTransit: sum(["transit"]),
    balances: lines.map((l) => {
      const ref = refOf(l.holder);
      return {
        account: l.reg.label,
        kind: l.reg.kind,
        ...(l.holder
          ? { holder: (ref && names.get(ref)) || l.holder, ...(ref ? { holderRef: ref } : {}) }
          : {}),
        amount: l.cents / 100,
        ...(l.ownCents !== l.cents ? { currencyAmount: l.ownCents / 100 } : {}),
      };
    }),
    note: notes.join(" "),
    scan: { rowsScanned: scanned, elapsedMs: Date.now() - t0 },
  };
}
