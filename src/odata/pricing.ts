import type { Connection } from "../context.js";
import { InputError } from "../errors.js";
import type { ODataEntity } from "../types/odata.js";
import { GUID_RE } from "./accounting.js";
import { ODataError } from "./errors.js";
import { fetchAllForAggregation } from "./aggregate.js";
import { cached } from "./cache.js";
import { strictText } from "./inventory.js";
import { buildQuery, cmp, odataGuid, or } from "./query.js";
import { refOf } from "./settlements-aging.js";

/**
 * Цены номенклатуры Управления торговлей (УТ 11, КА 2, ERP 2): документ «Установка цен номенклатуры» и регистр цен.
 * Два механизма: прежний (табличная часть «Товары», регистр «Цены номенклатуры») и ценообразование 2.5 («Товары2_5»,
 * регистр «Цены номенклатуры 2.5»). Признаки — положительные: нужные поля есть в схеме. Бухгалтерия устроена иначе
 * (типы цен, а не виды) и здесь не поддерживается.
 */
export const PRICE_DOC = "Document_УстановкаЦенНоменклатуры";
export const PRICE_KINDS = "Catalog_ВидыЦен";
const PRICE_DOC_TYPE = `StandardODATA.${PRICE_DOC}`;

export type PriceMechanism = "legacy" | "2.5";

interface MechanismSchema {
  /** Табличная часть документа. */
  rows: "Товары" | "Товары2_5";
  /** Регистр: набор записей по регистратору. */
  register: string;
  /** Записи регистра: срез последних и отбор по периоду. */
  records: string;
  /** Поле характеристики в строке и в регистре. */
  characteristic: "Характеристика_Key" | "ХарактеристикаЦО_Key";
  /** Пустые ссылки, обязательные в ключе записи 2.5. */
  emptyKeys: readonly string[];
}

export const MECHANISMS: Record<PriceMechanism, MechanismSchema> = {
  legacy: {
    rows: "Товары",
    register: "InformationRegister_ЦеныНоменклатуры",
    records: "InformationRegister_ЦеныНоменклатуры_RecordType",
    characteristic: "Характеристика_Key",
    emptyKeys: ["Упаковка_Key"],
  },
  "2.5": {
    rows: "Товары2_5",
    register: "InformationRegister_ЦеныНоменклатуры25",
    records: "InformationRegister_ЦеныНоменклатуры25_RecordType",
    characteristic: "ХарактеристикаЦО_Key",
    emptyKeys: ["СерияЦО_Key", "УпаковкаЦО_Key", "Упаковка_Key"],
  },
};

const EMPTY = "00000000-0000-0000-0000-000000000000";
const HEADER = ["Date", "Статус", "Согласован", "ВидыЦен"];
const ROW = ["Номенклатура_Key", "ВидЦены_Key", "Цена", "ЦенаИзмененаВручную", "Валюта_Key"];
const KIND_FIELDS = ["Ref_Key", "Description", "ВалютаЦены_Key"];

/** Механизмы, которые поддерживает схема этой базы: документ, вид цен и регистр — со всеми нужными полями. */
export async function priceMechanisms(conn: Connection): Promise<PriceMechanism[]> {
  const meta = await conn.getMetadata();
  const props = (set: string) => new Set(meta.entities.get(set)?.properties.map((p) => p.name) ?? []);
  const doc = props(PRICE_DOC);
  const kinds = props(PRICE_KINDS);
  if (!HEADER.every((f) => doc.has(f)) || !KIND_FIELDS.every((f) => kinds.has(f))) return [];
  return (Object.keys(MECHANISMS) as PriceMechanism[]).filter((m) => {
    const s = MECHANISMS[m];
    const row = props(`${PRICE_DOC}_${s.rows}`);
    const rec = props(s.records);
    return (
      doc.has(s.rows) &&
      meta.entities.has(s.register) &&
      [...ROW, s.characteristic].every((f) => row.has(f)) &&
      ["Period", "Номенклатура_Key", "ВидЦены_Key", "Цена", s.characteristic].every((f) => rec.has(f))
    );
  });
}

/** Отказ, если цены этой базы не устроены как в Управлении торговлей. */
export async function requirePriceSchema(conn: Connection): Promise<PriceMechanism[]> {
  const found = await priceMechanisms(conn);
  if (found.length) return found;
  throw new InputError(
    "Цены поддерживаются для Управления торговлей 11 (КА 2, ERP 2 — без проверки): в этой базе не опубликованы " +
      `${PRICE_DOC}, ${PRICE_KINDS} и регистр цен с нужными полями. В Бухгалтерии цены задаются типами цен, ` +
      "этот вариант пока не поддерживается.",
  );
}

/**
 * Механизм цен базы. Первый признак — какую табличную часть заполняют последние проведённые документы: после
 * перехода на 2.5 прежний регистр остаётся непустым, и записи в обоих регистрах — обычное дело. Второй — в каком
 * регистре есть записи. Не ясно — отказ с просьбой указать mechanism: ошибка в выборе дала бы документ без движений.
 */
export async function detectMechanism(conn: Connection, explicit?: PriceMechanism): Promise<PriceMechanism> {
  const found = await requirePriceSchema(conn);
  if (explicit) {
    if (!found.includes(explicit))
      throw new InputError(
        `Механизм цен ${explicit} в этой базе не опубликован; доступно: ${found.join(", ")}.`,
      );
    return explicit;
  }
  if (found.length === 1) return found[0]!;
  return cached(conn, "price-mechanism", async () => {
    // Только ссылки последних документов, строки — по одной из набора табличной части: строк в документе бывают
    // тысячи (ревью: $select табличных частей 20 документов тянул их все).
    const recent = await conn.client.getCollection(
      `${PRICE_DOC}${buildQuery({ select: ["Ref_Key"], filter: cmp("Posted", "eq", "true"), orderby: "Date desc", top: 5 })}`,
    );
    const refs = (recent.value ?? []).map((d) => String(d["Ref_Key"]));
    const used: PriceMechanism[] = [];
    if (refs.length)
      for (const m of found) {
        const rows = await conn.client.getCollection(
          `${PRICE_DOC}_${MECHANISMS[m].rows}${buildQuery({
            select: ["Ref_Key"],
            filter: or(...refs.map((r) => cmp("Ref_Key", "eq", odataGuid(r)))),
            top: 1,
          })}`,
        );
        if (rows.value?.length) used.push(m);
      }
    if (used.length === 1) return used[0]!;
    const withRecords: PriceMechanism[] = [];
    for (const m of found) {
      const r = await conn.client.getCollection(
        `${MECHANISMS[m].records}${buildQuery({ select: ["Period"], top: 1 })}`,
      );
      if (r.value?.length) withRecords.push(m);
    }
    if (withRecords.length === 1) return withRecords[0]!;
    throw new InputError(
      "Не удалось определить механизм цен этой базы: используются оба или ни один. Укажите mechanism: legacy — " +
        "прежний (табличная часть «Товары»), 2.5 — ценообразование 2.5 («Товары2_5»).",
    );
  });
}

export interface PriceKind {
  ref: string;
  name: string;
  currencyRef?: string;
  method?: string;
  status?: string;
  includesVat?: boolean;
  deleted: boolean;
  folder: boolean;
  baseRef?: string;
  influencing: string[];
}

/** Все виды цен (их немного): название, валюта, способ задания и от каких видов цена зависит. */
export function readPriceKinds(conn: Connection): Promise<PriceKind[]> {
  return cached(conn, "price-kinds", async () => {
    const props = new Set(
      (await conn.getMetadata()).entities.get(PRICE_KINDS)?.properties.map((p) => p.name),
    );
    const optional = [
      "DeletionMark",
      "IsFolder",
      "СпособЗаданияЦены",
      "Статус",
      "ЦенаВключаетНДС",
      "БазовыйВидЦены_Key",
      "ВлияющиеВидыЦен",
    ].filter((f) => props.has(f));
    const { rows } = await fetchAllForAggregation(
      conn,
      PRICE_KINDS,
      { select: [...KIND_FIELDS, ...optional] },
      "виды цен",
      5000,
    );
    return rows.map((r) => {
      const currencyRef = refOf(r["ВалютаЦены_Key"]);
      const baseRef = refOf(r["БазовыйВидЦены_Key"]);
      return {
        ref: String(r["Ref_Key"]),
        name: String(r["Description"] ?? ""),
        ...(currencyRef ? { currencyRef } : {}),
        ...(typeof r["СпособЗаданияЦены"] === "string" && r["СпособЗаданияЦены"]
          ? { method: r["СпособЗаданияЦены"] }
          : {}),
        ...(typeof r["Статус"] === "string" && r["Статус"] ? { status: r["Статус"] } : {}),
        ...(typeof r["ЦенаВключаетНДС"] === "boolean" ? { includesVat: r["ЦенаВключаетНДС"] } : {}),
        deleted: r["DeletionMark"] === true,
        folder: r["IsFolder"] === true,
        ...(baseRef ? { baseRef } : {}),
        influencing: ((r["ВлияющиеВидыЦен"] as ODataEntity[] | undefined) ?? [])
          .map((x) => refOf(x["ВлияющийВидЦен_Key"]))
          .filter((x): x is string => !!x),
      };
    });
  });
}

/** Вид цен по Ref_Key или точному названию; несколько с одним названием — отказ со списком. */
export function resolvePriceKind(query: string, kinds: readonly PriceKind[]): PriceKind {
  const q = query.trim();
  const usable = kinds.filter((k) => !k.folder);
  const found = GUID_RE.test(q)
    ? usable.filter((k) => k.ref.toLowerCase() === q.replace(/[{}]/g, "").toLowerCase())
    : usable.filter((k) => strictText(k.name.trim()) === strictText(q));
  if (found.length === 1) return found[0]!;
  if (!found.length) {
    const similar = usable
      .filter((k) => strictText(k.name).includes(strictText(q)))
      .map((k) => `${k.name} (${k.ref})`)
      .slice(0, 10);
    throw new InputError(
      `Вид цен "${query}" не найден.${similar.length ? ` Похожие: ${similar.join("; ")}.` : ""} Укажите Ref_Key или точное название.`,
    );
  }
  throw new InputError(
    `Под "${query}" подходит несколько видов цен: ${found.map((k) => `${k.name} (${k.ref})`).join("; ")}. Укажите Ref_Key.`,
  );
}

/** Виды цен, которые считаются от этого: при установке цены через OData 1С их не пересчитает. */
export function dependentKinds(kind: PriceKind, kinds: readonly PriceKind[]): PriceKind[] {
  return kinds.filter(
    (k) => !k.deleted && k.ref !== kind.ref && (k.baseRef === kind.ref || k.influencing.includes(kind.ref)),
  );
}

export interface PriceRecord {
  itemRef: string;
  characteristicRef?: string;
  kindRef: string;
  price: number;
  period: string;
  recorderRef?: string;
}

const toRecord =
  (mech: PriceMechanism) =>
  (r: ODataEntity): PriceRecord | undefined => {
    const itemRef = refOf(r["Номенклатура_Key"]);
    const kindRef = refOf(r["ВидЦены_Key"]);
    if (!itemRef || !kindRef) return undefined;
    const characteristicRef = refOf(r[MECHANISMS[mech].characteristic]);
    const recorderRef = refOf(r["Recorder"]);
    return {
      itemRef,
      ...(characteristicRef ? { characteristicRef } : {}),
      kindRef,
      price: Number(r["Цена"] ?? 0),
      period: String(r["Period"] ?? ""),
      ...(recorderRef ? { recorderRef } : {}),
    };
  };

/** Ключ цены в регистре: позиция, характеристика, вид цен. */
export const priceKey = (r: { itemRef: string; characteristicRef?: string | undefined; kindRef: string }) =>
  `${r.itemRef.toLowerCase()}|${(r.characteristicRef ?? "").toLowerCase()}|${r.kindRef.toLowerCase()}`;

/** Пачка ссылок в одном условии: длинный or упирается в лимит вложенности условий 1С. */
const BATCH = 20;

const quoted = (text: string) => `'${text.replace(/'/g, "''")}'`;

/** Действующие цены на момент (срез последних); без itemRefs — все позиции этих видов цен. */
export async function slicePrices(
  conn: Connection,
  mech: PriceMechanism,
  { kindRefs, itemRefs, at }: { kindRefs: readonly string[]; itemRefs?: readonly string[]; at?: string },
): Promise<PriceRecord[]> {
  const { records } = MECHANISMS[mech];
  const kinds = or(...kindRefs.map((k) => cmp("ВидЦены_Key", "eq", odataGuid(k))));
  const groups: Array<string[] | undefined> = itemRefs ? chunk([...new Set(itemRefs)], BATCH) : [undefined];
  const out: PriceRecord[] = [];
  for (const items of groups) {
    const condition = items
      ? `(${kinds}) and (${or(...items.map((i) => cmp("Номенклатура_Key", "eq", odataGuid(i))))})`
      : kinds;
    const params = [...(at ? [`Period=datetime'${at}'`] : []), `Condition=${quoted(condition)}`];
    const { rows } = await fetchAllForAggregation(
      conn,
      `${records}/SliceLast(${params.join(",")})`,
      {},
      `цены на ${at ?? "сейчас"}`,
    );
    for (const r of rows) {
      const rec = toRecord(mech)(r);
      if (rec) out.push(rec);
    }
  }
  return out;
}

/** Цены этих позиций и вида цен с периодом позже момента: новая цена с ранней датой их не перекроет. */
export async function laterPrices(
  conn: Connection,
  mech: PriceMechanism,
  { kindRef, itemRefs, after }: { kindRef: string; itemRefs: readonly string[]; after: string },
): Promise<PriceRecord[]> {
  const { records } = MECHANISMS[mech];
  const out: PriceRecord[] = [];
  for (const items of chunk([...new Set(itemRefs)], BATCH)) {
    const filter = [
      cmp("ВидЦены_Key", "eq", odataGuid(kindRef)),
      cmp("Period", "ge", `datetime'${after}'`),
      `(${or(...items.map((i) => cmp("Номенклатура_Key", "eq", odataGuid(i))))})`,
    ].join(" and ");
    const { rows } = await fetchAllForAggregation(conn, records, { filter }, "более поздние цены");
    for (const r of rows) {
      const rec = toRecord(mech)(r);
      if (rec) out.push(rec);
    }
  }
  return out;
}

/** Движения документа по регистру цен: набор записей по регистратору. */
export async function priceMovements(
  conn: Connection,
  mech: PriceMechanism,
  docRef: string,
): Promise<PriceRecord[]> {
  const { register } = MECHANISMS[mech];
  const guid = docRef.replace(/[{}']/g, "");
  let set: ODataEntity;
  try {
    set = await conn.client.getEntity<ODataEntity>(
      `${register}(Recorder=${quoted(guid)}, Recorder_Type=${quoted(PRICE_DOC_TYPE)})?$format=json`,
    );
  } catch (e) {
    // Живая УТ 11.5: у документа без движений набора записей нет — 404, а не пустой набор.
    if (e instanceof ODataError && e.kind === "not_found") return [];
    throw e;
  }
  return ((set["RecordSet"] as ODataEntity[] | undefined) ?? [])
    .filter((r) => r["Active"] !== false)
    .map(toRecord(mech))
    .filter((r): r is PriceRecord => !!r);
}

export interface PriceLine {
  itemRef: string;
  characteristicRef?: string;
  kindRef: string;
  price: number;
}

/** Строки документа цен в его табличной части. */
export async function priceDocumentLines(
  conn: Connection,
  mech: PriceMechanism,
  docRef: string,
): Promise<{ posted: boolean; lines: PriceLine[] }> {
  const s = MECHANISMS[mech];
  const guid = docRef.replace(/[{}']/g, "");
  const doc = await conn.client.getEntity<ODataEntity>(
    `${PRICE_DOC}(${odataGuid(guid)})${buildQuery({ select: ["Posted", s.rows] })}`,
  );
  const lines = ((doc[s.rows] as ODataEntity[] | undefined) ?? []).flatMap((r) => {
    const itemRef = refOf(r["Номенклатура_Key"]);
    const kindRef = refOf(r["ВидЦены_Key"]);
    if (!itemRef || !kindRef) return [];
    const characteristicRef = refOf(r[s.characteristic]);
    return [
      {
        itemRef,
        ...(characteristicRef ? { characteristicRef } : {}),
        kindRef,
        price: Number(r["Цена"] ?? 0),
      },
    ];
  });
  return { posted: doc["Posted"] === true, lines };
}

const cents = (n: number) => Math.round(n * 100);

/**
 * Сверка движений с документом: каждая строка дала запись с той же ценой до копейки, лишних записей нет. После отмены
 * проведения записей быть не должно.
 */
export function compareMovements(
  lines: readonly PriceLine[],
  movements: readonly PriceRecord[],
  posted: boolean,
): { matched: boolean; expected: number; found: number; mismatches: string[] } {
  const mismatches: string[] = [];
  if (!posted) {
    if (movements.length)
      mismatches.push(`после отмены проведения осталось записей в регистре: ${movements.length}`);
    return { matched: !mismatches.length, expected: 0, found: movements.length, mismatches };
  }
  const byKey = new Map(movements.map((m) => [priceKey(m), m]));
  for (const l of lines) {
    const m = byKey.get(priceKey(l));
    if (!m)
      mismatches.push(`нет записи для ${l.itemRef}${l.characteristicRef ? ` / ${l.characteristicRef}` : ""}`);
    else if (cents(m.price) !== cents(l.price))
      mismatches.push(`${l.itemRef}: в регистре ${m.price}, в документе ${l.price}`);
    byKey.delete(priceKey(l));
  }
  if (byKey.size) mismatches.push(`лишних записей в регистре: ${byKey.size}`);
  return { matched: !mismatches.length, expected: lines.length, found: movements.length, mismatches };
}

export { EMPTY as EMPTY_REF };

function chunk<T>(items: readonly T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}

/**
 * Проверка проведения документа цен: движения регистра совпадают со строками документа до копейки, после отмены
 * проведения записей нет. undefined — это не документ цен Управления торговлей (проверяют другие). Проведённый
 * документ с расхождением — не успех и не «ничего не сделано»: verified=false и следующий шаг.
 */
export async function verifyPriceMovements(
  conn: Connection,
  entitySet: string,
  ref: string,
  post: boolean,
): Promise<Record<string, unknown> | undefined> {
  if (entitySet !== PRICE_DOC) return undefined;
  const found = await priceMechanisms(conn);
  if (!found.length) return undefined;
  const guid = ref.replace(/[{}']/g, "");
  try {
    const doc = await conn.client.getEntity<ODataEntity>(
      `${PRICE_DOC}(${odataGuid(guid)})${buildQuery({ select: found.map((m) => MECHANISMS[m].rows) })}`,
    );
    const mech = found.find((m) => (doc[MECHANISMS[m].rows] as unknown[] | undefined)?.length) ?? found[0]!;
    const { lines } = await priceDocumentLines(conn, mech, guid);
    const movements = await priceMovements(conn, mech, guid);
    const check = compareMovements(lines, movements, post);
    return {
      movementsVerified: check.matched,
      movements: {
        mechanism: mech,
        expected: check.expected,
        found: check.found,
        ...(check.mismatches.length ? { mismatches: check.mismatches.slice(0, 20) } : {}),
      },
      ...(check.matched
        ? {}
        : {
            verified: false,
            note: post
              ? "Документ проведён, но движения регистра цен не совпали со строками. Проверьте документ в 1С; " +
                "отменить проведение — write.document.post_document с post=false."
              : "Проведение отменено, но в регистре цен остались записи документа. Проверьте его в 1С.",
          }),
    };
  } catch (e) {
    return {
      movementsVerified: false,
      verified: false,
      note:
        `Действие выполнено, но движения регистра цен прочитать не удалось (${e instanceof Error ? e.message : String(e)}). ` +
        "Проверьте документ в 1С, прежде чем повторять.",
    };
  }
}
