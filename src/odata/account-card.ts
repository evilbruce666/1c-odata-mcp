import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { accountingChart, accountingRegister } from "../config/accounting.js";
import { requireEntity } from "./publication.js";
import { fetchAllForAggregation, type ScanMeta, emptyMeta, addMeta } from "./aggregate.js";
import { and, cmp, odataGuid, or } from "./query.js";
import { num } from "./accounting.js";
import { InputError } from "../errors.js";
import { entitySetOfType, refOf } from "./settlements-aging.js";

/**
 * Движения по счёту за период — основа карточки счёта и акта сверки. Источник — виртуальная
 * таблица «Движения с субконто» (RecordsWithExtDimensions): в отличие от записей регистра
 * (_RecordType) она отдаёт субконто Дт и Кт с их типами. Отбор по счёту — в $filter (параметр
 * Account 1С не поддерживает), по контрагенту — cast(guid, 'Catalog_Контрагенты') на любой из
 * позиций субконто (без cast 1С отвечает 500 на поле неограниченной длины). Конец периода
 * включительно: проводки закрытия месяца в 23:59:59 попадают в выборку (проверено на живой базе).
 */

export const RECORDS_WITH_DIMS = "RecordsWithExtDimensions";
/**
 * Сортировка, при которой листание этой таблицы не теряет строк. Без неё 1С между страницами
 * отдаёт одни строки повторно, а другие пропускает: на живой базе $top=200 по 76.09 дал 912 строк,
 * из них 123 повтора вместо пропущенных; только по Period — 4. Дороже примерно вдвое, поэтому
 * применяется лишь при обнаруженных повторах.
 */
export const STABLE_MOVEMENT_ORDER = "Period asc,Recorder asc,LineNumber asc";
/** Счетов в одном $filter: каждый даёт два условия (Дт и Кт), 1С ограничивает вложенность. */
export const MOVEMENT_ACCOUNT_BATCH = 10;

const chunks = <T>(xs: readonly T[], n: number): T[][] =>
  Array.from({ length: Math.ceil(xs.length / n) }, (_, i) => xs.slice(i * n, i * n + n));

export interface Dim {
  ref: string;
  set?: string;
}

/** Одна сторона движения, попавшая в карточку: наш счёт на Дт (debit) или на Кт (credit). */
export interface CardEntry {
  period: string;
  lineNumber?: number;
  recorder?: string;
  recorderSet?: string;
  side: "debit" | "credit";
  cents: number;
  account: string;
  dims: Dim[];
  corrAccount?: string;
  corrDims: Dim[];
  content?: string;
  org?: string;
}

function dimsOf(r: ODataEntity, side: "Dr" | "Cr"): Dim[] {
  const out: Dim[] = [];
  for (const i of [1, 2, 3]) {
    const ref = refOf(r[`ExtDimension${side}${i}`]);
    if (!ref) continue;
    const set = entitySetOfType(r[`ExtDimension${side}${i}_Type`]);
    out.push({ ref, ...(set ? { set } : {}) });
  }
  return out;
}

/**
 * Движения → строки карточки. Сторона попадает в карточку, если её счёт входит в accounts и (при
 * отборе) среди её субконто есть все refs. Проводка внутри счёта (62.01 ↔ 62.02) даёт две строки.
 * Неактивные записи пропускаются — в обороты 1С они не входят.
 */
export function cardEntries(
  rows: readonly ODataEntity[],
  accounts: ReadonlySet<string>,
  refs: readonly string[] = [],
): CardEntry[] {
  const out: CardEntry[] = [];
  for (const r of rows) {
    if (r["Active"] === false) continue;
    const cents = Math.round(num(r["Сумма"]) * 100);
    const line = Number(r["LineNumber"]);
    const base = {
      period: String(r["Period"] ?? ""),
      ...(Number.isFinite(line) ? { lineNumber: line } : {}),
      ...(refOf(r["Recorder"]) ? { recorder: refOf(r["Recorder"]) } : {}),
      ...(entitySetOfType(r["Recorder_Type"]) ? { recorderSet: entitySetOfType(r["Recorder_Type"]) } : {}),
      ...(typeof r["Содержание"] === "string" && r["Содержание"] ? { content: String(r["Содержание"]) } : {}),
      ...(refOf(r["Организация_Key"]) ? { org: refOf(r["Организация_Key"]) } : {}),
      cents,
    };
    for (const [side, mine, other] of [
      ["debit", "Dr", "Cr"],
      ["credit", "Cr", "Dr"],
    ] as const) {
      const account = refOf(r[`Account${mine}_Key`]);
      if (!account || !accounts.has(account)) continue;
      const dims = dimsOf(r, mine);
      if (refs.some((ref) => !dims.some((d) => d.ref === ref))) continue;
      const corrAccount = refOf(r[`Account${other}_Key`]);
      out.push({
        ...base,
        side,
        account,
        dims,
        ...(corrAccount ? { corrAccount } : {}),
        corrDims: dimsOf(r, other),
      });
    }
  }
  return out.sort(
    (a, b) =>
      a.period.localeCompare(b.period) ||
      (a.recorder ?? "").localeCompare(b.recorder ?? "") ||
      (a.lineNumber ?? 0) - (b.lineNumber ?? 0) ||
      (a.side === b.side ? 0 : a.side === "debit" ? -1 : 1),
  );
}

const movementKey = (r: ODataEntity): string =>
  `${String(r["Recorder"])}|${String(r["Recorder_Type"])}|${String(r["LineNumber"])}`;
const hasRepeats = (rows: readonly ODataEntity[]): boolean =>
  new Set(rows.map(movementKey)).size !== rows.length;

/**
 * Все движения по счетам за период (с отбором по организации и контрагенту — на стороне 1С).
 * Счета — пачками по ACCOUNT_FILTER_BATCH; движение, попавшее в две пачки (Дт и Кт из разных),
 * берётся один раз. Повтор внутри одной выборки значит, что листание сбилось и какие-то строки
 * пропущены: такая выборка перечитывается с устойчивой сортировкой, а не молча схлопывается.
 */
export async function fetchMovements(
  conn: Connection,
  p: { accountKeys: readonly string[]; from: string; to: string; orgKey?: string; counterpartyRef?: string },
): Promise<{ rows: ODataEntity[]; meta: ScanMeta }> {
  const reg = await requireEntity(conn, accountingRegister(conn.cfg), "Регистр бухгалтерии «Хозрасчётный»");
  const path = `${reg}/${RECORDS_WITH_DIMS}(StartPeriod=datetime'${p.from}T00:00:00',EndPeriod=datetime'${p.to}T23:59:59')`;
  const cp = p.counterpartyRef
    ? or(
        ...(["Dr", "Cr"] as const).flatMap((s) =>
          [1, 2, 3].map((i) =>
            cmp(
              `ExtDimension${s}${i}`,
              "eq",
              `cast(${odataGuid(p.counterpartyRef!)}, 'Catalog_Контрагенты')`,
            ),
          ),
        ),
      )
    : undefined;
  const org = p.orgKey ? cmp("Организация_Key", "eq", odataGuid(p.orgKey)) : undefined;
  // С контрагентом отбор по нему избирателен сам по себе: счета фильтруем у себя (cardEntries),
  // иначе длинный OR по счетам упирается в лимит вложенности условий 1С («Превышение допустимой
  // вложенности операций»). Без контрагента — счета пачками по MOVEMENT_ACCOUNT_BATCH.
  const filters = cp
    ? [and(org, cp)]
    : chunks(p.accountKeys, MOVEMENT_ACCOUNT_BATCH).map((batch) =>
        and(
          or(
            ...batch.flatMap((k) => [
              cmp("AccountDr_Key", "eq", odataGuid(k)),
              cmp("AccountCr_Key", "eq", odataGuid(k)),
            ]),
          ),
          org,
        ),
      );
  const seen = new Set<string>();
  const rows: ODataEntity[] = [];
  let meta = emptyMeta();
  for (const filter of filters) {
    const label = `движения ${p.from}..${p.to}`;
    let part = await fetchAllForAggregation(conn, path, { filter: filter || undefined }, label);
    meta = addMeta(meta, part.meta);
    if (hasRepeats(part.rows)) {
      part = await fetchAllForAggregation(
        conn,
        path,
        { filter: filter || undefined, orderby: STABLE_MOVEMENT_ORDER },
        label,
      );
      meta = addMeta(meta, part.meta);
      if (hasRepeats(part.rows)) {
        throw new InputError(
          `1С вернула движения за ${p.from}..${p.to} с повторами, точный итог посчитать нельзя. ` +
            "Сузьте период или повторите запрос.",
        );
      }
    }
    for (const r of part.rows) {
      const k = movementKey(r);
      if (seen.has(k)) continue;
      seen.add(k);
      rows.push(r);
    }
  }
  return { rows, meta };
}

/**
 * Отбор по субконто, которое на этом счёте ведётся только по оборотам (статья ДДС на 51, статья
 * затрат на 91.02): сальдо по нему в регистре нет, сверять движения не с чем. Признак — TurnoversOnly
 * в плане счетов на той позиции, где субконто стоит в движениях. Истина, если есть ref, у которого
 * все такие позиции только оборотные. План счетов не прочитался — обычное расхождение, а не отказ.
 */
export async function isTurnoversOnlyFilter(
  conn: Connection,
  rows: readonly ODataEntity[],
  accounts: ReadonlySet<string>,
  refs: readonly string[],
): Promise<boolean> {
  const positions = new Map<string, Set<string>>();
  for (const r of rows) {
    for (const side of ["Dr", "Cr"] as const) {
      const account = refOf(r[`Account${side}_Key`]);
      if (!account || !accounts.has(account)) continue;
      for (const i of [1, 2, 3]) {
        const ref = refOf(r[`ExtDimension${side}${i}`]);
        if (ref && refs.includes(ref)) {
          positions.set(ref, (positions.get(ref) ?? new Set()).add(`${account}|${i}`));
        }
      }
    }
  }
  if (!positions.size) return false;
  const chart = await requireEntity(conn, accountingChart(conn.cfg), "План счетов «Хозрасчётный»");
  const turnoversOnly = new Set<string>();
  try {
    const keys = new Set([...positions.values()].flatMap((s) => [...s].map((k) => k.split("|")[0]!)));
    for (const key of keys) {
      const a = await conn.client.getEntity<ODataEntity>(
        `${chart}(guid'${key}')?$select=Ref_Key,ExtDimensionTypes`,
      );
      for (const t of (a["ExtDimensionTypes"] ?? []) as ODataEntity[]) {
        if (t["TurnoversOnly"] === true) turnoversOnly.add(`${key}|${Number(t["LineNumber"])}`);
      }
    }
  } catch {
    return false;
  }
  return [...positions.values()].some((s) => [...s].every((k) => turnoversOnly.has(k)));
}

/** Чистое сальдо (Дт − Кт, копейки) строк Balance, где среди субконто есть все refs. */
export function netBalance(rows: readonly ODataEntity[], refs: readonly string[] = []): number {
  let cents = 0;
  for (const r of rows) {
    const dims = [1, 2, 3].map((i) => refOf(r[`ExtDimension${i}`]));
    if (refs.some((ref) => !dims.includes(ref))) continue;
    cents += Math.round(num(r["СуммаBalanceDr"]) * 100) - Math.round(num(r["СуммаBalanceCr"]) * 100);
  }
  return cents;
}
