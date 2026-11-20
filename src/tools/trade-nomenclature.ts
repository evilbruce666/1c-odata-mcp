import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { EMPTY_REF } from "../odata/pricing.js";
import { and, buildQuery, cmp, odataGuid, odataString, or } from "../odata/query.js";
import { fail } from "./_shared.js";
import { createCore } from "./write.js";
import { previewDay } from "./trade-counterparty.js";

/**
 * Номенклатура Управления торговлей. В 1С карточка берёт свойства из вида номенклатуры (тип, характеристики, серии,
 * упаковки, группы учёта): на копии рабочей УТ 11.5 у всех 4000 просмотренных позиций эти поля совпадают с видом.
 * Через OData 1С их не переносит — коннектор переносит сам. Единица по умолчанию — «шт» (код ОКЕИ 796), ставка НДС —
 * основная действующая (решение владельца, 20.11.2026).
 */

export const KINDS = "Catalog_ВидыНоменклатуры";
const UNITS = "Catalog_УпаковкиЕдиницыИзмерения";
const RATES = "Catalog_СтавкиНДС";

/** Поля, которые 1С копирует из вида номенклатуры в позицию (проверено на живой базе: совпадают у всех позиций). */
const FROM_KIND = [
  "ТипНоменклатуры",
  "ИспользованиеХарактеристик",
  "ВариантОформленияПродажи",
  "ВладелецСерий_Key",
  "ВладелецХарактеристик_Key",
  "ВладелецТоварныхКатегорий_Key",
  "ГруппаАналитическогоУчета_Key",
  "ГруппаФинансовогоУчета_Key",
  "ГруппаДоступа_Key",
  "ИспользоватьУпаковки",
  "НаборУпаковок_Key",
  "СкладскаяГруппа_Key",
  "ВестиУчетПоГТД",
  "ВестиУчетСертификатовНоменклатуры",
  "АлкогольнаяПродукция",
  "ИмпортнаяАлкогольнаяПродукция",
  "ВидАлкогольнойПродукции_Key",
  "ПодконтрольнаяПродукцияВЕТИС",
  "ПродукцияМаркируемаяДляГИСМ",
  "КиЗГИСМ",
  "ОблагаетсяНДСУПокупателя",
  "ПоставляетсяВМногооборотнойТаре",
  "НоменклатураМногооборотнаяТара_Key",
  "ХарактеристикаМногооборотнаяТара_Key",
  "ЕдиницаИзмеренияСрокаГодности",
  "КонтролироватьКачествоПоступившегоТовара",
  "СрокКонтроляКачества",
  "ОбособленнаяЗакупкаПродажа",
  "ЦеноваяГруппа_Key",
  "СезоннаяГруппа_Key",
  "ШаблонЦенника_Key",
  "ШаблонЭтикетки_Key",
] as const;

const isEmpty = (v: unknown) =>
  v === undefined || v === null || v === "" || v === false || v === 0 || v === "0" || v === EMPTY_REF;

/** Схема номенклатуры УТ: у позиции есть вид номенклатуры и единица из «Упаковки и единицы измерения». */
export async function tradeNomenclatureSchema(conn: Connection, set: string): Promise<boolean> {
  const available = await conn.available();
  if (!available.has(KINDS) || !available.has(UNITS)) return false;
  const props = (await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? [];
  return props.includes("ВидНоменклатуры_Key") && props.includes("ЕдиницаИзмерения_Key");
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Один элемент по Ref_Key или точному наименованию; несколько или ни одного — текст отказа. */
async function one(
  conn: Connection,
  set: string,
  query: string,
  label: string,
  extra: string[] = [],
): Promise<ODataEntity | string> {
  const q = query.trim();
  const filter = GUID.test(q)
    ? cmp("Ref_Key", "eq", odataGuid(q))
    : and("DeletionMark eq false", cmp("Description", "eq", odataString(q)));
  const { value } = await conn.client.getCollection<ODataEntity>(
    `${set}${buildQuery({ filter, select: ["Ref_Key", "Description", ...extra], top: 5 })}`,
  );
  if (value.length === 1) return value[0]!;
  if (value.length === 0) return `${label} «${q}» не найден(а). Ничего не создано.`;
  return (
    `${label}: под «${q}» подходит несколько — ${value.map((r) => `${String(r["Description"])} (${String(r["Ref_Key"])})`).join("; ")}. ` +
    "Укажите Ref_Key. Ничего не создано."
  );
}

/**
 * Единица по названию или коду. Одинаковых бывает несколько (на живой УТ — две «шт»: «0796» у 14998 позиций и
 * заведённая позже «796 » у двух): тогда берётся та, у которой позиций хотя бы вдвое больше, чем у следующей; иначе —
 * отказ со ссылками и числами.
 */
async function resolveUnit(
  conn: Connection,
  set: string,
  query: string,
  notes: string[],
): Promise<ODataEntity | string> {
  const q = query.trim();
  if (GUID.test(q)) return one(conn, UNITS, q, "Единица измерения");
  const { value } = await conn.client.getCollection<ODataEntity>(
    `${UNITS}${buildQuery({
      filter: and(
        "DeletionMark eq false",
        or(
          cmp("Description", "eq", odataString(q)),
          cmp("Code", "eq", odataString(q)),
          cmp("Code", "eq", odataString(q.padStart(4, "0"))),
        ),
      ),
      select: ["Ref_Key", "Description", "Code"],
      top: 5,
    })}`,
  );
  if (value.length === 1) return value[0]!;
  if (value.length === 0) return `Единица измерения «${q}» не найдена. Ничего не создано.`;
  const counts: number[] = [];
  for (const u of value) {
    const r = await conn.client.getCollection<ODataEntity>(
      `${set}${buildQuery({
        filter: cmp("ЕдиницаИзмерения_Key", "eq", odataGuid(String(u["Ref_Key"]))),
        select: ["Ref_Key"],
        top: 1,
        count: true,
      })}`,
    );
    counts.push(Number(r["odata.count"] ?? r.value.length));
  }
  const order = counts.map((n, i) => [n, i] as const).sort((a, b) => b[0] - a[0]);
  const [first, second] = order;
  const list = value
    .map(
      (r, i) =>
        `${String(r["Description"])}, код ${String(r["Code"]).trim()} (${String(r["Ref_Key"])}) — позиций ${counts[i]}`,
    )
    .join("; ");
  // Явный перевес — хотя бы вдвое больше позиций, чем у следующей: иначе угадывать нельзя.
  if (first && first[0] > 0 && (!second || first[0] >= 2 * second[0])) {
    notes.push(
      `Единиц «${q}» в базе ${value.length}: ${list}. Взята самая используемая; другая — unit с её Ref_Key.`,
    );
    return value[first[1]]!;
  }
  return `Единица измерения: под «${q}» подходит несколько — ${list}. Укажите Ref_Key. Ничего не создано.`;
}

/** Основная действующая ставка НДС на день: признак «Основная», не расчётная, день внутри периода действия. */
async function mainVatRate(conn: Connection, day: string): Promise<ODataEntity | string> {
  const { value } = await conn.client.getCollection<ODataEntity>(
    `${RATES}${buildQuery({
      filter: and("DeletionMark eq false", "Основная eq true", "РасчетнаяСтавка eq false"),
      select: ["Ref_Key", "Description", "НачалоПериода", "КонецПериода"],
      top: 20,
    })}`,
  );
  const date = day.slice(0, 10);
  const open = (v: unknown) => !v || String(v).startsWith("0001");
  const current = value.filter(
    (r) =>
      (open(r["НачалоПериода"]) || String(r["НачалоПериода"]).slice(0, 10) <= date) &&
      (open(r["КонецПериода"]) || String(r["КонецПериода"]).slice(0, 10) >= date),
  );
  if (current.length === 1) return current[0]!;
  return current.length === 0
    ? "Основная действующая ставка НДС не найдена в справочнике «Ставки НДС»: укажите vatRate. Ничего не создано."
    : "Основных действующих ставок НДС несколько: укажите vatRate. Ничего не создано.";
}

export interface TradeNomenclatureInput {
  name: string;
  fullName?: string | undefined;
  article?: string | undefined;
  parentRef?: string | undefined;
  isService: boolean;
  kind?: string | undefined;
  unit?: string | undefined;
  vatRate?: string | undefined;
}

export async function createTradeNomenclature(
  conn: Connection,
  set: string,
  input: TradeNomenclatureInput,
  confirm: boolean,
): Promise<CallToolResult> {
  if (!input.kind)
    return fail(
      "В Управлении торговлей у номенклатуры обязателен вид (kind — название или Ref_Key из справочника «Виды " +
        "номенклатуры»): от него зависят тип, серии, характеристики и учёт. Ничего не создано.",
    );
  // Только поля, опубликованные и у вида, и у позиции: в другой версии УТ части может не быть (иначе 400 на $select).
  const meta = await conn.getMetadata();
  const propsOf = (e: string) => new Set(meta.entities.get(e)?.properties.map((p) => p.name));
  const kindProps = propsOf(KINDS);
  const itemProps = propsOf(set);
  const copied = FROM_KIND.filter((f) => kindProps.has(f) && itemProps.has(f));
  const kindRow = await one(conn, KINDS, input.kind, "Вид номенклатуры", [...copied, "IsFolder"]);
  if (typeof kindRow === "string") return fail(kindRow);
  if (kindRow["IsFolder"] === true)
    return fail(
      `«${String(kindRow["Description"])}» — группа видов номенклатуры, а не вид. Ничего не создано.`,
    );
  if (input.isService && kindRow["ТипНоменклатуры"] !== "Услуга")
    return fail(
      `isService=true, а вид «${String(kindRow["Description"])}» — ${String(kindRow["ТипНоменклатуры"])}: выберите вид ` +
        "услуг. Ничего не создано.",
    );
  const unitNotes: string[] = [];
  const unit = await resolveUnit(conn, set, input.unit ?? "796", unitNotes);
  if (typeof unit === "string") return fail(unit);
  if (!(await conn.available()).has(RATES))
    return fail(
      "Справочник «Ставки НДС» (Catalog_СтавкиНДС) не опубликован в OData: ставку для номенклатуры не выбрать. " +
        "Опубликуйте его. Ничего не создано.",
    );
  const day = await previewDay(conn);
  const rate = input.vatRate
    ? await one(conn, RATES, input.vatRate, "Ставка НДС")
    : await mainVatRate(conn, day);
  if (typeof rate === "string") return fail(rate);

  const fromKind = Object.fromEntries(copied.map((f) => [f, kindRow[f]]).filter(([, v]) => !isEmpty(v)));
  const payload = {
    Description: input.name,
    НаименованиеПолное: input.fullName ?? input.name,
    ...(input.article ? { Артикул: input.article } : {}),
    ...(input.parentRef ? { Parent_Key: input.parentRef } : {}),
    ВидНоменклатуры_Key: String(kindRow["Ref_Key"]),
    ...fromKind,
    ЕдиницаИзмерения_Key: String(unit["Ref_Key"]),
    ЕдиницаДляОтчетов_Key: String(unit["Ref_Key"]),
    КоэффициентЕдиницыДляОтчетов: 1,
    СтавкаНДС_Key: String(rate["Ref_Key"]),
  };
  const notes = [
    ...unitNotes,
    `Вид «${String(kindRow["Description"])}», единица «${String(unit["Description"])}», ставка НДС «${String(rate["Description"])}»` +
      (input.vatRate ? "." : " — основная действующая."),
  ];
  return createCore(conn, set, payload, confirm, notes, () => undefined);
}
