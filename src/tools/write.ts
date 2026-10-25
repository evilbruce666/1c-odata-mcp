import {
  ALIGN_HINT,
  bankPostingBlocker,
  expectedDocumentVersion,
  guardedDocumentAction,
} from "../odata/existing-bank-settlement.js";
import { operationKind as configuredOperationKind, type AccountDefault } from "../config/accounting.js";
import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, organizationField } from "./_shared.js";
import { CATALOGS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { resolveOrgOrDefault } from "../odata/orgs.js";
import { ensurePublished, requireEntity } from "../odata/publication.js";
import {
  accountsByCode,
  accountsByRef,
  nomenclatureAccounts,
  resolveDefaultAccounts,
  type NomAccounts,
} from "../odata/accounting.js";
import { settlementAccountsFor, type SettlementSide } from "../odata/settlement-accounts.js";
import {
  contactKindsForCounterparties,
  resolveBankByBik,
  resolveAdditionalProperty,
  type ContactKinds,
} from "../odata/refdata.js";
import { fetchAll } from "../odata/pagination.js";
import { and, buildQuery, cmp, contains, odataGuid, odataString } from "../odata/query.js";
import type { ODataEntity } from "../types/odata.js";
import {
  createResultSchema,
  patchResultSchema,
  markForDeletionResultSchema,
  postDocumentResultSchema,
  operationStatusSchema,
} from "../schemas/output.js";
import { currentWriteOperationId, fingerprintRecordFields } from "../odata/write-operation-context.js";
import { InputError } from "../errors.js";
import { ODataError } from "../odata/errors.js";
import { getDocumentPostings } from "./registers.js";
import { tryPreviewView } from "./preview-view.js";

/** Тип ссылки на номенклатуру в табличной части (полиморфная ссылка 1С). */
const NOMENCLATURE_TYPE = "StandardODATA.Catalog_Номенклатура";
const COUNTERPARTY_TYPE = "StandardODATA.Catalog_Контрагенты";

/** Строит строки табличной части «КонтактнаяИнформация» из телефона/email/адреса. */
function buildContactRows(
  kinds: ContactKinds,
  data: { phone?: string | undefined; email?: string | undefined; address?: string | undefined },
): Array<Record<string, unknown>> {
  const rows: Array<Record<string, unknown>> = [];
  if (data.phone && kinds.phone)
    rows.push({
      Тип: kinds.phone.тип,
      Вид_Key: kinds.phone.key,
      Представление: data.phone,
      НомерТелефона: data.phone,
    });
  if (data.email && kinds.email)
    rows.push({
      Тип: kinds.email.тип,
      Вид_Key: kinds.email.key,
      Представление: data.email,
      АдресЭП: data.email,
    });
  if (data.address && kinds.address)
    rows.push({ Тип: kinds.address.тип, Вид_Key: kinds.address.key, Представление: data.address });
  return rows.map((r, i) => ({ LineNumber: i + 1, ...r }));
}

/** Дополнительные поля карточки контрагента: контактная информация + ОГРН (доп.реквизит). */
async function counterpartyExtras(
  conn: Connection,
  data: {
    phone?: string | undefined;
    email?: string | undefined;
    address?: string | undefined;
    ogrn?: string | undefined;
  },
): Promise<{ fields: Record<string, unknown>; notes: string[] }> {
  const fields: Record<string, unknown> = {};
  const notes: string[] = [];
  if (data.phone || data.email || data.address) {
    const kinds = await contactKindsForCounterparties(conn);
    const rows = buildContactRows(kinds, data);
    if (rows.length) fields["КонтактнаяИнформация"] = rows;
    if (data.phone && !kinds.phone) notes.push("Вид «Телефон» не найден — телефон не записан.");
    if (data.email && !kinds.email) notes.push("Вид «Email» не найден — email не записан.");
    if (data.address && !kinds.address) notes.push("Вид «Адрес» не найден — адрес не записан.");
  }
  if (data.ogrn) {
    const prop = await resolveAdditionalProperty(conn, "ОГРН");
    if (prop)
      fields["ДополнительныеРеквизиты"] = [{ LineNumber: 1, Свойство_Key: prop, Значение: data.ogrn }];
    else
      notes.push(
        "ОГРН не записан: в этой базе нет дополнительного реквизита «ОГРН» для контрагентов. Заведите его в 1С (Администрирование → Дополнительные реквизиты), либо вносите ОГРН вручную.",
      );
  }
  return { fields, notes };
}

/** Виды договоров (Enum_ВидыДоговоровКонтрагентов). */
const CONTRACT_KINDS = ["СПокупателем", "СПоставщиком", "Прочее", "СКомиссионером", "СКомитентом"] as const;

/** Ставки НДС (Enum_СтавкиНДС), подмножество ходовых. */
const VAT_RATES = ["БезНДС", "НДС0", "НДС5", "НДС7", "НДС10", "НДС20", "НДС22"] as const;

/**
 * Дата → формат 1С Edm.DateTime ('YYYY-MM-DDTHH:mm:ss', без зоны).
 * Берём ЛОКАЛЬНЫЕ компоненты (не toISOString/UTC): 1С хранит «настенную» дату без
 * зоны, а UTC-сдвиг в плюсовых поясах уводит полночь на предыдущий день.
 */
export function odataDate(d: Date): string {
  const p = (n: number): string => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/** Резолвит организацию: по названию, либо авто, если в базе ровно одна. */
async function resolveOrg(
  conn: Connection,
  organization: string | undefined,
): Promise<{ key: string; name: string }> {
  const o = await resolveOrgOrDefault(conn, organization);
  return { key: o.ref, name: o.name };
}

/** Резолвит склад по названию; если не задан и склад один — берёт его; иначе undefined. */
async function resolveWarehouse(conn: Connection, name: string | undefined): Promise<string | undefined> {
  const set = resolveEntity(CATALOGS.warehouses, await conn.available());
  if (!set) return undefined;
  if (name) {
    const { rows } = await fetchAll(
      conn.client,
      set,
      { filter: contains("Description", name), select: ["Ref_Key", "Description"] },
      5,
      5,
    );
    const first = rows[0];
    if (!first) throw new InputError(`Склад "${name}" не найден (см. list_entities / справочник Склады).`);
    return String(first["Ref_Key"]);
  }
  const { rows } = await fetchAll(conn.client, set, { select: ["Ref_Key"] }, 2, 2);
  return rows.length === 1 ? String((rows[0] as ODataEntity)["Ref_Key"]) : undefined;
}

/**
 * Резолвит элемент справочника по точному коду или части наименования.
 * Публикацию проверяет (requireEntity). Бросает Error, если не найдено.
 */
async function resolveCatalogItem(
  conn: Connection,
  candidates: readonly string[],
  label: string,
  query: string,
): Promise<{ ref: string; code?: string; name?: string }> {
  const set = await requireEntity(conn, candidates, label);
  const byCode = await fetchAll(
    conn.client,
    set,
    { filter: cmp("Code", "eq", odataString(query)), select: ["Ref_Key", "Code", "Description"] },
    3,
    3,
  );
  const rows = byCode.rows.length
    ? byCode.rows
    : (
        await fetchAll(
          conn.client,
          set,
          { filter: contains("Description", query), select: ["Ref_Key", "Code", "Description"] },
          5,
          5,
        )
      ).rows;
  const first = rows[0];
  if (!first) throw new InputError(`${label}: «${query}» не найдено (по коду или наименованию).`);
  return {
    ref: String(first["Ref_Key"]),
    code: first["Code"] ? String(first["Code"]) : undefined,
    name: first["Description"] ? String(first["Description"]) : undefined,
  };
}

const GUID_RE = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;

/** Находит папку (группу, IsFolder=true) в иерархическом справочнике по части имени. */
async function resolveFolder(
  conn: Connection,
  entitySet: string,
  query: string,
): Promise<{ ref: string; name: string }> {
  const { rows } = await fetchAll(
    conn.client,
    entitySet,
    {
      filter: and(contains("Description", query), cmp("IsFolder", "eq", "true")),
      select: ["Ref_Key", "Description"],
    },
    10,
    10,
  );
  if (rows.length === 0) throw new InputError(`Папка «${query}» не найдена в ${entitySet}.`);
  if (rows.length > 1) {
    const names = rows.map((r) => String(r["Description"])).join(", ");
    throw new InputError(`Под «${query}» несколько папок: ${names}. Уточните название.`);
  }
  const r = rows[0] as ODataEntity;
  return { ref: String(r["Ref_Key"]), name: String(r["Description"] ?? "") };
}

/** Папка по ref (GUID) или по имени (резолв). undefined → без родителя. */
async function folderRefOf(
  conn: Connection,
  entitySet: string,
  folder: string | undefined,
): Promise<string | undefined> {
  if (!folder) return undefined;
  return GUID_RE.test(folder.trim())
    ? folder.replace(/[{}]/g, "")
    : (await resolveFolder(conn, entitySet, folder)).ref;
}

export interface GoodsLine {
  nomenclatureRef: string;
  quantity: number;
  price: number;
  vatRate: string;
  /** Содержание строки — печатается в счёте/УПД (для услуг обязательно по смыслу). */
  content?: string | undefined;
  /** Счёт доходов: код плана счетов («90.01.2») или готовый Ref_Key. */
  incomeAccount?: string | undefined;
  /** Счёт расходов/себестоимости: код («90.02.2») или Ref_Key. */
  expenseAccount?: string | undefined;
  /** Исходная строка документа (при правке существующего) — чтобы не терять её поля. */
  source?: Record<string, unknown> | undefined;
}
type LineAccountsFor = (line: GoodsLine) => Record<string, string>;

/** Счета в строке не заполняются (счёт поставщика, счёт покупателю — проводок нет). */
const noAccounts: LineAccountsFor = () => ({});

/** Округление денег до копеек — суммы складываем только через него. */
const roundMoney = (n: number): number => Math.round(n * 100) / 100;
const lineSum = (l: GoodsLine): number => roundMoney(l.quantity * l.price);
const rowsTotal = (rows: Array<Record<string, unknown>>): number =>
  Math.round(rows.reduce((s, r) => s + (r["Сумма"] as number), 0) * 100) / 100;

/**
 * Поля исходной строки, которые НЕ переносим: считаем заново (иначе после смены
 * количества/цены останется старая сумма НДС) либо это служебная ссылка.
 */
const NOT_CARRIED_ROW_FIELDS = ["LineNumber", "Сумма", "СуммаНДС", "СуммаСНДС", "Ref_Key"] as const;

/**
 * Слагаемые суммы НДС строки не изменились — значит прежняя сумма налога верна.
 * Пересчитать её мы не можем: формула зависит от СуммаВключаетНДС документа и
 * процента ставки, и угадывать тут нельзя. Но и терять нельзя — иначе у документа
 * с НДС строка уезжает в 1С с нулевым налогом.
 */
function vatInputsUnchanged(l: GoodsLine): boolean {
  const s = l.source;
  if (!s) return false;
  return (
    Number(s["Количество"]) === l.quantity &&
    Number(s["Цена"]) === l.price &&
    String(s["СтавкаНДС"] ?? "БезНДС") === l.vatRate &&
    // Сумму строки мы считаем как количество×цену. Если в исходной строке она была
    // другой (ручная скидка в 1С), то и прежний налог посчитан не от нашей суммы.
    Number(s["Сумма"]) === lineSum(l)
  );
}

/**
 * Прочие поля исходной строки (номенклатурная группа, ГТД, страна происхождения,
 * субконто) — переносим как есть: мы ими не управляем, но терять их нельзя.
 * СуммаНДС переносится только при неизменных количестве/цене/ставке.
 */
export function carriedRowFields(l: GoodsLine): Record<string, unknown> {
  if (!l.source) return {};
  const out: Record<string, unknown> = { ...l.source };
  for (const f of NOT_CARRIED_ROW_FIELDS) delete out[f];
  if (vatInputsUnchanged(l)) out["СуммаНДС"] = l.source["СуммаНДС"];
  return out;
}

/** Строки табличной части «Товары»/«Услуги» для поступления/реализации (Номенклатура_Key + счета). */
export function buildGoodsRows(
  lines: GoodsLine[],
  lineAccountsFor: LineAccountsFor,
): Array<Record<string, unknown>> {
  return lines.map((l, i) =>
    clean({
      ...carriedRowFields(l),
      LineNumber: i + 1,
      Номенклатура_Key: l.nomenclatureRef,
      // Только когда задано: иначе затёрли бы содержание, перенесённое из исходной строки.
      ...(l.content === undefined ? {} : { Содержание: l.content }),
      Количество: l.quantity,
      Цена: l.price,
      Сумма: lineSum(l),
      СтавкаНДС: l.vatRate,
      ...lineAccountsFor(l),
    }),
  );
}

/** Строки «Товары» для счёта покупателю (полиморфная ссылка Номенклатура+_Type). */
export function buildInvoiceRows(lines: GoodsLine[]): Array<Record<string, unknown>> {
  return lines.map((l, i) =>
    clean({
      ...carriedRowFields(l),
      LineNumber: i + 1,
      Номенклатура: l.nomenclatureRef,
      Номенклатура_Type: NOMENCLATURE_TYPE,
      // Только когда задано: иначе затёрли бы содержание, перенесённое из исходной строки.
      ...(l.content === undefined ? {} : { Содержание: l.content }),
      Количество: l.quantity,
      Цена: l.price,
      Сумма: lineSum(l),
      СтавкаНДС: l.vatRate,
    }),
  );
}

/** Строит и создаёт/предпросматривает товарный документ (поступление/реализация). */
async function createGoodsDoc(
  conn: Connection,
  entitySet: string,
  p: {
    orgKey: string;
    counterpartyRef: string;
    contractRef?: string | undefined;
    warehouseKey?: string | undefined;
    date?: string | undefined;
    sumIncludesVat: boolean;
    lines: GoodsLine[];
    settlement?: string | undefined; // СчетУчетаРасчетовСКонтрагентом_Key (шапка)
    bankAccountKey?: string | undefined; // БанковскийСчетОрганизации_Key (есть у реализации)
    lineAccountsFor: LineAccountsFor; // счета строки
  },
  confirm: boolean,
) {
  const rows = buildGoodsRows(p.lines, p.lineAccountsFor);
  const payload = clean({
    Date: odataDate(p.date ? new Date(`${p.date}T00:00:00`) : new Date()),
    Posted: false,
    Организация_Key: p.orgKey,
    Контрагент_Key: p.counterpartyRef,
    ДоговорКонтрагента_Key: p.contractRef,
    Склад_Key: p.warehouseKey,
    СчетУчетаРасчетовСКонтрагентом_Key: p.settlement,
    БанковскийСчетОрганизации_Key: p.bankAccountKey,
    СуммаВключаетНДС: p.sumIncludesVat,
    СуммаДокумента: rowsTotal(rows),
    Товары: rows,
  });
  return createOrPreview(conn, entitySet, payload, confirm);
}

/**
 * Резолвит счета учёта по каждой номенклатуре: сперва из регистра «Счета учёта
 * номенклатуры», недостающее добивает дефолтами по кодам плана счетов.
 */
async function resolveLineAccounts(
  conn: Connection,
  orgKey: string,
  nomRefs: string[],
  defaults: NomAccounts,
): Promise<Map<string, NomAccounts>> {
  const map = new Map<string, NomAccounts>();
  for (const ref of new Set(nomRefs)) {
    const reg = await nomenclatureAccounts(conn, orgKey, ref);
    map.set(ref, {
      goods: reg?.goods ?? defaults.goods,
      incomingVat: reg?.incomingVat ?? defaults.incomingVat,
      outgoingVat: reg?.outgoingVat ?? defaults.outgoingVat,
      income: reg?.income ?? defaults.income,
      expense: reg?.expense ?? defaults.expense,
    });
  }
  return map;
}

/**
 * Делит заданные пользователем счета на готовые ссылки (GUID — берём как есть,
 * нормализуя фигурные скобки) и коды плана счетов, которые надо найти в базе.
 * Ключ карты и элементы codes — исходное значение после trim.
 */
export function splitAccountRefs(values: Array<string | undefined>): {
  refs: Map<string, string>;
  codes: string[];
} {
  const refs = new Map<string, string>();
  const codes = new Set<string>();
  for (const v of values) {
    const t = v?.trim();
    if (!t) continue;
    if (GUID_RE.test(t)) refs.set(t, t.replace(/[{}]/g, ""));
    else codes.add(t);
  }
  return { refs, codes: [...codes] };
}

/**
 * Резолвит заданные пользователем счета: код плана счетов («90.01.2») → Ref_Key,
 * готовый GUID — как есть. Все коды одним запросом; ключ карты — исходная строка
 * после trim. Неизвестный код — ошибка, чтобы не записать документ «мимо» счёта.
 */
async function resolveAccountOverrides(
  conn: Connection,
  values: Array<string | undefined>,
): Promise<Map<string, string>> {
  const { refs, codes } = splitAccountRefs(values);
  if (codes.length === 0) return refs;
  const found = await accountsByCode(conn, codes);
  for (const c of codes) {
    const ref = found.get(c);
    if (!ref)
      throw new InputError(
        `Счёт «${c}» не найден в плане счетов «Хозрасчётный». Укажите код точно как в 1С (напр. 90.01.2) или Ref_Key счёта.`,
      );
    refs.set(c, ref);
  }
  return refs;
}

/** Счета учёта, заданные на уровне документа (применяются ко всем строкам). */
interface DocAccounts {
  income?: string | undefined;
  expense?: string | undefined;
}

/**
 * Готовит счета учёта для товарного документа: счёт расчётов (шапка) и функцию
 * счетов строки по номенклатуре. kind различает поступление (Дт 41 Кт 60) и
 * реализацию (Дт 62 Кт 90, Дт 90 Кт 41). Приоритет счетов доходов/расходов:
 * строка → документ → регистр «Счета учёта номенклатуры» → коды по умолчанию.
 */
async function goodsAccounts(
  conn: Connection,
  orgKey: string,
  lines: GoodsLine[],
  kind: "purchase" | "shipment" | "service",
  docAccounts?: DocAccounts,
): Promise<{ settlement?: string | undefined; lineAccountsFor: LineAccountsFor }> {
  const accMap = await resolveLineAccounts(
    conn,
    orgKey,
    lines.map((l) => l.nomenclatureRef),
    {},
  );
  const overrides =
    kind === "purchase"
      ? new Map<string, string>()
      : await resolveAccountOverrides(conn, [
          docAccounts?.income,
          docAccounts?.expense,
          ...lines.flatMap((l) => [l.incomeAccount, l.expenseAccount]),
        ]);
  const ref = (v: string | undefined): string | undefined => overrides.get(v?.trim() ?? "");
  const needed = new Set<AccountDefault>(
    conn.cfg?.accounting?.strictDefaults ? [] : [kind === "purchase" ? "payable" : "receivable"],
  );
  for (const l of lines) {
    const a = accMap.get(l.nomenclatureRef);
    if (kind !== "service" && !a?.goods) needed.add("goods");
    if (kind === "purchase") {
      if (l.vatRate !== "БезНДС" && !a?.incomingVat) needed.add("incomingVat");
    } else {
      if (!(ref(l.incomeAccount) ?? ref(docAccounts?.income) ?? a?.income)) needed.add("income");
      if (!(ref(l.expenseAccount) ?? ref(docAccounts?.expense) ?? a?.expense)) needed.add("expense");
      if (l.vatRate !== "БезНДС" && !a?.outgoingVat) needed.add("outgoingVat");
    }
  }
  const defaults = await resolveDefaultAccounts(conn, [...needed]);
  return {
    settlement: defaults.get(kind === "purchase" ? "payable" : "receivable"),
    lineAccountsFor: (l) => {
      const a = accMap.get(l.nomenclatureRef);
      const goods = kind === "service" ? {} : { СчетУчета_Key: a?.goods ?? defaults.get("goods") };
      if (kind === "purchase")
        return clean({
          ...goods,
          ...(l.vatRate !== "БезНДС"
            ? { СчетУчетаНДС_Key: a?.incomingVat ?? defaults.get("incomingVat") }
            : {}),
        }) as Record<string, string>;
      return clean({
        ...goods,
        СчетДоходов_Key:
          ref(l.incomeAccount) ?? ref(docAccounts?.income) ?? a?.income ?? defaults.get("income"),
        СчетРасходов_Key:
          ref(l.expenseAccount) ?? ref(docAccounts?.expense) ?? a?.expense ?? defaults.get("expense"),
        ...(l.vatRate !== "БезНДС"
          ? { СчетУчетаНДСПоРеализации_Key: a?.outgoingVat ?? defaults.get("outgoingVat") }
          : {}),
      }) as Record<string, string>;
    },
  };
}

/**
 * Счёт учёта товаров (41.01) на строку — для складских документов без расчётов
 * с контрагентом (перемещение/оприходование/списание/инвентаризация).
 */
async function goodsOnlyAccounts(
  conn: Connection,
  orgKey: string,
  lines: GoodsLine[],
): Promise<(nomRef: string) => Record<string, string>> {
  const accMap = await resolveLineAccounts(
    conn,
    orgKey,
    lines.map((l) => l.nomenclatureRef),
    {},
  );
  const needsGoods = lines.some((l) => !accMap.get(l.nomenclatureRef)?.goods);
  const defaults = await resolveDefaultAccounts(conn, needsGoods ? ["goods"] : []);
  return (nomRef) =>
    clean({ СчетУчета_Key: accMap.get(nomRef)?.goods ?? defaults.get("goods") }) as Record<string, string>;
}

const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";
/** Непустая ссылка строкой; пустой GUID и пустая строка → undefined. */
const refOrUndefined = (v: unknown): string | undefined => {
  const s = typeof v === "string" ? v.replace(/[{}']/g, "") : "";
  return s && s !== EMPTY_GUID ? s : undefined;
};

/**
 * Разбирает строку ТЧ документа в GoodsLine, сохраняя исходную строку в source:
 * при пересборке ТЧ (add/remove/update_document_lines) содержание, счета и прочие
 * поля строки должны пережить правку, а не обнулиться.
 */
export function lineFromRow(r: Record<string, unknown>): GoodsLine {
  return {
    nomenclatureRef: String(r["Номенклатура_Key"] ?? r["Номенклатура"] ?? ""),
    quantity: Number(r["Количество"] ?? 0),
    price: Number(r["Цена"] ?? 0),
    vatRate: String(r["СтавкаНДС"] ?? "БезНДС"),
    content: (r["Содержание"] as string) || undefined,
    incomeAccount: refOrUndefined(r["СчетДоходов_Key"]),
    expenseAccount: refOrUndefined(r["СчетРасходов_Key"]),
    source: r,
  };
}

/** Табличная часть документа, в которой лежат позиции. */
export type Section = "Товары" | "Услуги";

/**
 * Определяет табличную часть документа. У реализации их две: «Товары» и «Услуги»,
 * и заполнена ровно одна — акт услуг (ВидОперации=Услуги) живёт в «Услугах».
 * Ориентируемся сперва на факт (где есть строки), затем на тип и вид операции —
 * у пустого документа строк ещё нет.
 */
export function pickSection(
  doc: { goods?: unknown[] | undefined; services?: unknown[] | undefined },
  hints: { servicesAct: boolean; operationKind?: unknown },
): Section {
  if (doc.services?.length) return "Услуги";
  if (doc.goods?.length) return "Товары";
  if (hints.servicesAct) return "Услуги";
  return hints.operationKind === "Услуги" ? "Услуги" : "Товары";
}

/** Строки табличной части документа (пустой массив, если ТЧ нет или она пуста). */
export function sectionRows(doc: Record<string, unknown>, section: Section): Array<Record<string, unknown>> {
  return (doc[section] as Array<Record<string, unknown>> | undefined) ?? [];
}

/** Вторая табличная часть — та, которую сейчас не правим. */
const otherSection = (s: Section): Section => (s === "Товары" ? "Услуги" : "Товары");

/**
 * Сумма «Сумма» по строкам. Документ может нести позиции в ОБЕИХ ТЧ (реализация
 * с товарами и услугами), поэтому итог документа = правимая ТЧ + вторая как есть:
 * считать только по одной значило бы занизить СуммаДокумента.
 */
export function rowsSum(rows: Array<Record<string, unknown>>): number {
  return roundMoney(rows.reduce((s, r) => s + Number(r["Сумма"] ?? 0), 0));
}

/**
 * Читает документ: организация, проведён ли, заполненная ТЧ и её строки.
 * otherTotal — сумма ВТОРОЙ табличной части: её мы не трогаем, но в
 * СуммаДокумента она входит (у реализации бывают заполнены и «Товары», и «Услуги»).
 */
async function getDocInfo(
  conn: Connection,
  entitySet: string,
  guid: string,
): Promise<{
  orgKey: string;
  posted: boolean;
  lines: GoodsLine[];
  section: Section;
  otherTotal: number;
}> {
  const doc = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
  const section = pickSection(
    {
      goods: doc["Товары"] as unknown[] | undefined,
      services: doc["Услуги"] as unknown[] | undefined,
    },
    {
      servicesAct: (DOCUMENTS.servicesAct as readonly string[]).includes(entitySet),
      operationKind: doc["ВидОперации"],
    },
  );
  return {
    orgKey: String(doc["Организация_Key"] ?? ""),
    posted: doc["Posted"] === true,
    lines: sectionRows(doc, section).map(lineFromRow),
    section,
    otherTotal: rowsSum(sectionRows(doc, otherSection(section))),
  };
}

/**
 * Реквизиты шапки, которые в копию не переносим: они принадлежат исходному
 * документу. Number отсутствует намеренно — 1С автонумерует новый документ,
 * только если Number не передан вовсе.
 */
const NOT_COPIED_DOC_FIELDS = [
  "Ref_Key",
  "Number",
  "Posted",
  "DeletionMark",
  "DataVersion",
  "СсылочныйИдентификатор",
] as const;

/**
 * Поля строк ТЧ, которые не копируем: идентификаторы строк связывают строку с
 * исходным документом и его «потомками» — в копии они должны быть своими.
 */
const NOT_COPIED_ROW_FIELDS = ["Ref_Key", "ИдентификаторСтроки", "ИдентификаторРодительскойСтроки"] as const;

/**
 * Готовит прочитанный документ к записи как нового: убирает служебные поля 1С,
 * технические ключи OData (odata.metadata, *@navigationLinkUrl) и идентификаторы
 * строк табличных частей. Остальное переносится как есть — в этом и смысл копии.
 */
export function stripForCopy(entity: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(entity)) {
    if (key.startsWith("odata.") || key.includes("@")) continue;
    if ((NOT_COPIED_DOC_FIELDS as readonly string[]).includes(key)) continue;
    if (Array.isArray(value)) {
      out[key] = value.map((row) => {
        if (!row || typeof row !== "object") return row;
        const copy: Record<string, unknown> = { ...(row as Record<string, unknown>) };
        for (const f of NOT_COPIED_ROW_FIELDS) delete copy[f];
        return copy;
      });
      continue;
    }
    out[key] = value;
  }
  return out;
}

/** Переопределение полей одной строки табличной части по её номеру. */
export interface LineOverride {
  lineNumber: number;
  fields: Record<string, unknown>;
}

/**
 * Применяет переопределения к строкам ТЧ по номеру строки. Номер, которого нет
 * в документе, — ошибка ввода: молча проигнорировать правку суммы нельзя.
 */
export function applyLineOverrides(
  rows: Array<Record<string, unknown>>,
  overrides: LineOverride[],
): Array<Record<string, unknown>> {
  const byNumber = new Map(overrides.map((o) => [o.lineNumber, o.fields]));
  const missing = overrides.filter((o) => o.lineNumber < 1 || o.lineNumber > rows.length);
  if (missing.length) {
    throw new InputError(
      `В документе ${rows.length} строк(и), а правка задана для строки ${missing
        .map((m) => m.lineNumber)
        .join(", ")}.`,
    );
  }
  return rows.map((row, i) => ({ ...row, ...(byNumber.get(i + 1) ?? {}) }));
}

/**
 * Сумма документа по строкам ТЧ. undefined, если хотя бы у одной строки нет
 * числовой «Суммы» — тогда пересчитывать нельзя и сумму оставляем исходную.
 */
export function totalFromRows(rows: Array<Record<string, unknown>>): number | undefined {
  let sum = 0;
  for (const row of rows) {
    const v = row["Сумма"];
    const n = typeof v === "number" ? v : typeof v === "string" ? Number(v) : NaN;
    if (!Number.isFinite(n)) return undefined;
    sum += n;
  }
  return Math.round(sum * 100) / 100;
}

/** Читает документ-основание счёта-фактуры: организация, контрагент, договор, суммы, НДС. */
async function invoiceBasis(
  conn: Connection,
  entitySet: string,
  ref: string,
): Promise<{ org: string; counterparty: string; contract?: string; total: number; vat: number }> {
  const guid = ref.replace(/[{}']/g, "");
  const doc = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
  // НДС собираем по ОБЕИМ табличным частям: акт услуг (create_act) живёт в
  // «Услугах», и чтение одних «Товаров» давало счёт-фактуру с нулевым НДС.
  const rows = [...sectionRows(doc, "Товары"), ...sectionRows(doc, "Услуги")];
  const vat = Math.round(rows.reduce((s, r) => s + Number(r["СуммаНДС"] ?? 0), 0) * 100) / 100;
  return {
    org: String(doc["Организация_Key"] ?? ""),
    counterparty: String(doc["Контрагент_Key"] ?? doc["Контрагент"] ?? ""),
    contract: (doc["ДоговорКонтрагента_Key"] as string) || undefined,
    total: Number(doc["СуммаДокумента"] ?? 0),
    vat,
  };
}

/**
 * Собирает строки табличной части под тип документа (счёт/поступление/реализация/акт).
 * section — куда пойдут строки: для «Услуг» счёт учёта 41 не ставится (склада нет).
 */
async function buildSectionRows(
  conn: Connection,
  entitySet: string,
  orgKey: string,
  lines: GoodsLine[],
  section: Section,
): Promise<{ rows: Array<Record<string, unknown>>; total: number } | undefined> {
  const inList = (arr: readonly string[]): boolean => arr.includes(entitySet);
  if (inList(DOCUMENTS.customerInvoice)) {
    const rows = buildInvoiceRows(lines);
    return { rows, total: rowsTotal(rows) };
  }
  if (inList(DOCUMENTS.servicesAct)) {
    const { lineAccountsFor } = await goodsAccounts(conn, orgKey, lines, "service");
    const rows = buildGoodsRows(lines, lineAccountsFor);
    return { rows, total: rowsTotal(rows) };
  }
  // Реализация/поступление и возвраты: возврат покупателя считается как реализация
  // (Дт 90/62), возврат поставщику — как поступление (Дт 41/19).
  if (
    inList(DOCUMENTS.purchases) ||
    inList(DOCUMENTS.sales) ||
    inList(DOCUMENTS.returnFromCustomer) ||
    inList(DOCUMENTS.returnToSupplier)
  ) {
    const sale = inList(DOCUMENTS.sales) || inList(DOCUMENTS.returnFromCustomer);
    const kind = sale ? (section === "Услуги" ? "service" : "shipment") : "purchase";
    const { lineAccountsFor } = await goodsAccounts(conn, orgKey, lines, kind);
    const rows = buildGoodsRows(lines, lineAccountsFor);
    return { rows, total: rowsTotal(rows) };
  }
  if (inList(DOCUMENTS.surplus) || inList(DOCUMENTS.writeoff) || inList(DOCUMENTS.transfer)) {
    const lineAcc = await goodsOnlyAccounts(conn, orgKey, lines);
    const withSum = !inList(DOCUMENTS.transfer); // у перемещения в ТЧ нет суммы
    const rows = lines.map((l, i) =>
      clean({
        ...carriedRowFields(l),
        LineNumber: i + 1,
        Номенклатура_Key: l.nomenclatureRef,
        ...(l.content === undefined ? {} : { Содержание: l.content }),
        Количество: l.quantity,
        Цена: l.price,
        ...(withSum ? { Сумма: lineSum(l) } : {}),
        ...lineAcc(l.nomenclatureRef),
      }),
    );
    return { rows, total: rowsTotal(rows) };
  }
  return undefined;
}

/**
 * Банковский счёт организации. Приоритет: явный Ref_Key → поиск по названию/номеру →
 * ОСНОВНОЙ счёт организации → первый по алфавиту. Раньше без названия брался
 * произвольный счёт (rows[0] без сортировки) — на счёте и в акте печатались
 * реквизиты не того банка.
 */
export async function resolveOrgBankAccount(
  conn: Connection,
  orgKey: string,
  query: string | undefined,
): Promise<string | undefined> {
  const q = query?.trim();
  if (q && GUID_RE.test(q)) return q.replace(/[{}]/g, "");
  if (!q) {
    const main = await mainOrgBankAccount(conn, orgKey);
    if (main) return main;
    // Подбор не удался — не критично: 1С подставит счёт организации по умолчанию. Из неполного
    // списка «первый по алфавиту» был бы случайным счётом, поэтому тогда не подставляем ничего.
    const list = await orgBankAccounts(conn, orgKey).catch(() => undefined);
    return list && !list.truncated ? list.accounts[0]?.ref : undefined;
  }
  // Ошибка чтения — не «счёта нет»: пусть дойдёт до пользователя как есть (#54).
  const { accounts, truncated } = await orgBankAccounts(conn, orgKey);
  const needle = q.toLowerCase();
  const hits = accounts.filter((a) => a.name.toLowerCase().includes(needle) || a.number.includes(q));
  // Список счетов обрезан: нужного может не быть в прочитанной части, а единственное совпадение —
  // не гарантия единственности. Без явного Ref_Key принимаем только точный номер счёта.
  if (truncated) {
    const exact = hits.filter((a) => a.number === q);
    if (exact.length === 1) return exact[0]!.ref;
    throw new InputError(
      `Банковский счёт организации «${q}»: в базе больше ${ORG_BANK_ACCOUNTS_MAX} банковских счетов, ` +
        "прочитана только часть, поэтому по названию счёт не подобрать. Укажите полный номер счёта или его Ref_Key.",
    );
  }
  if (hits.length === 1) return hits[0]?.ref;
  if (hits.length === 0) {
    const known = accounts.map((a) => a.name).join("; ");
    throw new InputError(
      `Банковский счёт организации «${q}» не найден.` +
        (known ? ` Есть: ${known}.` : " У организации нет счетов."),
    );
  }
  // Несколько совпадений: основной среди них — берём его, иначе просим уточнить.
  const main = await mainOrgBankAccount(conn, orgKey);
  const preferred = hits.find((a) => a.ref === main);
  if (preferred) return preferred.ref;
  throw new InputError(
    `Под «${q}» подходит несколько счетов организации: ${hits.map((a) => a.name).join("; ")}. Уточните номер счёта.`,
  );
}

/**
 * Банковские счета организации списком.
 *
 * Владелец счёта — полиморфная ссылка Owner (+ Owner_Type). Сравнивать саму Owner
 * в $filter 1С не даёт: «Нельзя сравнивать поля неограниченной длины и поля
 * несовместимых типов» (HTTP 500). Зато Owner_Type — обычная строка, ею и сужаем
 * выборку до счетов организаций, а конкретного владельца сверяем на клиенте.
 * Отбор по названию тоже клиентский: substringof по Description 1С отклоняет
 * по той же причине.
 */
/** Сколько банковских счетов (всех организаций) читаем для подбора по названию. */
const ORG_BANK_ACCOUNTS_MAX = 1000;

async function orgBankAccounts(
  conn: Connection,
  orgKey: string,
): Promise<{ accounts: Array<{ ref: string; name: string; number: string }>; truncated: boolean }> {
  const available = await conn.available();
  const set = resolveEntity(CATALOGS.bankAccounts, available);
  if (!set) return { accounts: [], truncated: false };
  const orgSet = resolveEntity(CATALOGS.organizations, available);
  const filter = and(
    orgSet ? cmp("Owner_Type", "eq", odataString(`StandardODATA.${orgSet}`)) : undefined,
    cmp("DeletionMark", "eq", "false"),
  );
  const { rows, truncated } = await fetchAll(
    conn.client,
    set,
    { filter, select: ["Ref_Key", "Description", "НомерСчета", "Owner"] },
    conn.behavior.pageSize,
    ORG_BANK_ACCOUNTS_MAX,
  );
  const accounts = rows
    .filter((r) => String(r["Owner"] ?? "") === orgKey)
    .map((r) => ({
      ref: String(r["Ref_Key"] ?? ""),
      name: String(r["Description"] ?? ""),
      number: String(r["НомерСчета"] ?? ""),
    }))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
  return { accounts, truncated };
}

/** Основной банковский счёт организации (реквизит ОсновнойБанковскийСчет). */
async function mainOrgBankAccount(conn: Connection, orgKey: string): Promise<string | undefined> {
  const orgSet = resolveEntity(CATALOGS.organizations, await conn.available());
  if (!orgSet) return undefined;
  try {
    const org = await conn.client.getEntity(
      `${orgSet}(guid'${orgKey.replace(/[{}']/g, "")}')${buildQuery({ select: ["ОсновнойБанковскийСчет_Key"] })}`,
    );
    return refOrUndefined(org["ОсновнойБанковскийСчет_Key"]);
  } catch {
    return undefined;
  }
}

/** Валюта документа по умолчанию: код 643 (рубль), иначе — единственная в справочнике. */
async function resolveDefaultCurrency(conn: Connection): Promise<string | undefined> {
  const set = resolveEntity(CATALOGS.currencies, await conn.available());
  if (!set) return undefined;
  try {
    const byCode = await fetchAll(
      conn.client,
      set,
      { filter: cmp("Code", "eq", odataString("643")), select: ["Ref_Key"] },
      2,
      2,
    );
    if (byCode.rows[0]) return String(byCode.rows[0]["Ref_Key"]);
    const all = await fetchAll(conn.client, set, { select: ["Ref_Key"] }, 2, 2);
    return all.rows.length === 1 ? String((all.rows[0] as ODataEntity)["Ref_Key"]) : undefined;
  } catch {
    return undefined;
  }
}

const confirmField = z
  .boolean()
  .default(false)
  .describe(
    "false (по умолчанию) — только предпросмотр (dry-run), запись НЕ выполняется. " +
      "true — выполнить создание. Сначала всегда показывайте dry-run и получайте согласие пользователя; " +
      "при подтверждении передавайте operationId из результата предпросмотра и те же аргументы. После таймаута " +
      "повторяйте подтверждение с тем же id, не создавая новый предпросмотр.",
  );

const contentField = z
  .string()
  .optional()
  .describe(
    "Содержание строки (реквизит «Содержание») — именно этот текст печатается в счёте и УПД. " +
      "Для услуг заполняйте обязательно, напр. «Услуги согласно приложению № 4 от 01.10.2025 г. " +
      "к договору № 87 от 01.12.2023 г. за июль 2026 г.». 1С сама это поле НЕ заполняет.",
  );

const incomeAccountField = z
  .string()
  .optional()
  .describe(
    "Счёт доходов — код плана счетов как в 1С (напр. «90.01.1» или «90.01.2») либо Ref_Key счёта. " +
      "Без указания берётся из регистра «Счета учёта номенклатуры», иначе 90.01.1.",
  );

const expenseAccountField = z
  .string()
  .optional()
  .describe(
    "Счёт расходов (себестоимости) — код как в 1С (напр. «90.02.1» или «90.02.2») либо Ref_Key. " +
      "Без указания берётся из регистра «Счета учёта номенклатуры», иначе 90.02.1.",
  );

const orgBankAccountField = z
  .string()
  .optional()
  .describe(
    "Банковский счёт организации — печатается в документе как реквизиты для оплаты. " +
      "Название/номер счёта (напр. «Сбербанк» или «40802…») либо Ref_Key. " +
      "Без указания берётся ОСНОВНОЙ счёт организации.",
  );

/** Базовые поля позиции документа (без счетов учёта). */
const baseLineShape = {
  nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
  quantity: z.number().positive().describe("Количество"),
  price: z.number().nonnegative().describe("Цена за единицу"),
  vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
  content: contentField,
};

/** Позиция документа продажи: со счетами доходов/расходов на строку. */
const saleLine = z.object({
  ...baseLineShape,
  incomeAccount: incomeAccountField,
  expenseAccount: expenseAccountField,
});

/** Позиция документа закупки: счета доходов/расходов неприменимы. */
const purchaseLine = z.object(baseLineShape);

/** Убирает undefined-поля, чтобы не слать их в 1С. */
function clean(obj: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined && v !== ""));
}

const resolveSet = (conn: Connection, candidates: readonly string[], human: string): Promise<string> =>
  requireEntity(conn, candidates, human);

const OPERATION_MARKER_RE = /\s*\[op:[0-9a-f-]{36}\]/gi;
const operationMarker = (operationId: string) => `[op:${operationId}]`;

/**
 * Дописывает в «Комментарий» метку операции — по ней write.operation.status находит
 * созданный объект после таймаута. Только для сущностей, у которых есть «Комментарий».
 * Метки чужих операций (напр. из копируемого документа) вырезаются.
 */
async function withOperationMarker(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const operationId = currentWriteOperationId();
  if (!operationId || !conn.behavior.writeOperationMarker) return payload;
  if (!(await entityHasComment(conn, entitySet))) return payload;
  const prev = typeof payload["Комментарий"] === "string" ? payload["Комментарий"] : "";
  const base = prev.replace(OPERATION_MARKER_RE, "").trim();
  return { ...payload, Комментарий: `${base}${base ? " " : ""}${operationMarker(operationId)}` };
}

async function entityHasComment(conn: Connection, entitySet: string): Promise<boolean> {
  const meta = (await conn.getMetadata()).entities.get(entitySet);
  return !!meta?.properties.some((p) => p.name === "Комментарий");
}

/**
 * Вид договора, который 1С подставляет в документ при выборе контрагента — по направлению
 * документа. Денежные документы сюда не входят: там вид зависит от вида операции.
 */
const CONTRACT_KIND_BY_DOC: ReadonlyMap<string, "СПокупателем" | "СПоставщиком"> = new Map([
  ...[
    ...DOCUMENTS.sales,
    ...DOCUMENTS.customerInvoice,
    ...DOCUMENTS.servicesAct,
    ...DOCUMENTS.returnFromCustomer,
  ].map((set) => [set, "СПокупателем"] as const),
  ...[...DOCUMENTS.purchases, ...DOCUMENTS.supplierInvoice, ...DOCUMENTS.returnToSupplier].map(
    (set) => [set, "СПоставщиком"] as const,
  ),
]);

const EMPTY_REF = "00000000-0000-0000-0000-000000000000";

/**
 * Договор не указан — поступаем как форма 1С при выборе контрагента: если у контрагента ровно один
 * договор нужного вида с этой организацией, подставляем его. Несколько или ни одного — не угадываем,
 * а предупреждаем: без договора 1С документ не проведёт (в БП 3.0 падает невнятной ошибкой модуля
 * прослеживаемости «Преобразование значения к типу Булево»).
 */
export async function withDefaultContract(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
): Promise<{ payload: Record<string, unknown>; notes: string[] }> {
  const kind = CONTRACT_KIND_BY_DOC.get(entitySet);
  const counterparty = payload["Контрагент_Key"];
  if (!kind || typeof counterparty !== "string" || !counterparty || payload["ДоговорКонтрагента_Key"]) {
    return { payload, notes: [] };
  }
  const available = await conn.available();
  const set = CATALOGS.contracts.find((c) => available.has(c));
  if (!set) return { payload, notes: [] };
  const org = typeof payload["Организация_Key"] === "string" ? payload["Организация_Key"] : undefined;
  const found = await conn.client.getCollection<ODataEntity>(
    `${set}${buildQuery({
      select: ["Ref_Key", "Description"],
      filter: and(
        cmp("Owner_Key", "eq", odataGuid(counterparty)),
        cmp("ВидДоговора", "eq", odataString(kind)),
        cmp("DeletionMark", "eq", "false"),
        org ? cmp("Организация_Key", "eq", odataGuid(org)) : undefined,
      ),
      top: 3,
    })}`,
  );
  const kindText = kind === "СПокупателем" ? "с покупателем" : "с поставщиком";
  if (found.value.length === 1) {
    const ctr = found.value[0]!;
    return {
      payload: { ...payload, ДоговорКонтрагента_Key: ctr["Ref_Key"] },
      notes: [
        `Договор не указан — подставлен единственный договор ${kindText}: «${String(ctr["Description"] ?? "")}».`,
      ],
    };
  }
  return {
    payload,
    notes: [
      found.value.length === 0
        ? `У контрагента нет договора ${kindText} — документ создастся, но 1С его не проведёт. ` +
          "Создайте договор (write.catalog.create_contract; вид уточните у пользователя) и передайте contractRef."
        : `У контрагента несколько договоров ${kindText} — укажите нужный в contractRef, ` +
          "иначе 1С документ не проведёт.",
    ],
  };
}

const SETTLE_FIELD = "СчетУчетаРасчетовСКонтрагентом_Key";
const ADVANCE_FIELD = "СчетУчетаРасчетовПоАвансам_Key";
const PAYMENT_ROWS = "РасшифровкаПлатежа";
const MONEY_DOCUMENTS: ReadonlySet<string> = new Set([
  ...DOCUMENTS.bankIn,
  ...DOCUMENTS.bankOut,
  ...DOCUMENTS.cashIn,
  ...DOCUMENTS.cashOut,
]);
const LEVEL_TEXT = {
  договор: "настройки договора",
  контрагент: "настройки контрагента",
  группа: "настройки группы",
  организация: "настройки организации",
  общая: "общей настройки",
} as const;

/** Сторона расчётов документа: по виду документа, для денежных — по виду операции. */
function documentSettlementSide(
  entitySet: string,
  payload: Record<string, unknown>,
): SettlementSide | undefined {
  const kind = CONTRACT_KIND_BY_DOC.get(entitySet);
  if (kind) return kind === "СПокупателем" ? "customer" : "supplier";
  const op = typeof payload["ВидОперации"] === "string" ? payload["ВидОперации"] : "";
  if (/Покупател/i.test(op)) return "customer";
  if (/Поставщик/i.test(op)) return "supplier";
  return undefined;
}

/** В строгом режиме defaults нужны лишь после более приоритетной настройки расчётов. */
async function strictSettlementDefaults(
  conn: Connection,
  payload: Record<string, unknown>,
  side: SettlementSide | undefined,
  hasHeader: boolean,
  choice?: { settle?: string | undefined; advance?: string | undefined },
): Promise<{ payload: Record<string, unknown>; used: boolean }> {
  if (!conn.cfg?.accounting?.strictDefaults || !side) return { payload, used: false };
  const settleGroup = side === "customer" ? "receivable" : "payable";
  const advanceGroup = side === "customer" ? "receivableAdvance" : "payableAdvance";
  const rows = Array.isArray(payload[PAYMENT_ROWS])
    ? (payload[PAYMENT_ROWS] as Record<string, unknown>[])
    : [];
  const matching = rows.filter((r) => r["ДоговорКонтрагента_Key"] === payload["ДоговорКонтрагента_Key"]);
  const missing = new Set<AccountDefault>();
  if (
    !choice?.settle &&
    ((hasHeader && !refOrUndefined(payload[SETTLE_FIELD])) ||
      matching.some((r) => !refOrUndefined(r[SETTLE_FIELD])))
  )
    missing.add(settleGroup);
  if (!choice?.advance && matching.some((r) => !refOrUndefined(r[ADVANCE_FIELD]))) missing.add(advanceGroup);
  const defaults = await resolveDefaultAccounts(conn, [...missing]);
  const fill = (target: Record<string, unknown>, advance: boolean) =>
    clean({
      ...target,
      ...(!choice?.settle && !refOrUndefined(target[SETTLE_FIELD])
        ? { [SETTLE_FIELD]: defaults.get(settleGroup) }
        : {}),
      ...(advance && !choice?.advance && !refOrUndefined(target[ADVANCE_FIELD])
        ? { [ADVANCE_FIELD]: defaults.get(advanceGroup) }
        : {}),
    });
  return {
    payload: {
      ...(hasHeader ? fill(payload, false) : payload),
      ...(rows.length ? { [PAYMENT_ROWS]: rows.map((r) => (matching.includes(r) ? fill(r, true) : r)) } : {}),
    },
    used: missing.size > 0,
  };
}

/**
 * Счета расчётов и авансов из регистра «Счета учёта расчётов с контрагентами» (#38). 1С через OData
 * настройку договора не применяет, поэтому подставляем сами: самая частная запись по договору →
 * контрагенту → организации с учётом вида расчётов. Регистр не опубликован или подходящей записи нет —
 * остаются счета по умолчанию, и предпросмотр говорит об этом прямо.
 */
export async function withSettlementAccounts(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
  side: SettlementSide | undefined = documentSettlementSide(entitySet, payload),
): Promise<{ payload: Record<string, unknown>; notes: string[] }> {
  // Для оплаты по договору счета нужны в расшифровке, а не в шапке (#39).
  // Явную шапку сохраняем; отсутствие поля в payload не заменяем наличием в metadata.
  const rowSettlement =
    MONEY_DOCUMENTS.has(entitySet) &&
    side !== undefined &&
    Array.isArray(payload[PAYMENT_ROWS]) &&
    payload[PAYMENT_ROWS].length > 0;
  const contractRef = payload["ДоговорКонтрагента_Key"];
  if (typeof contractRef !== "string" || !contractRef || contractRef === EMPTY_REF) {
    if (!conn.cfg?.accounting?.strictDefaults) return { payload, notes: [] };
    const hasHeader =
      !rowSettlement &&
      ((await conn.getMetadata()).entities.get(entitySet)?.properties.some((p) => p.name === SETTLE_FIELD) ??
        false);
    const fallback = await strictSettlementDefaults(conn, payload, side, hasHeader);
    return {
      payload: fallback.payload,
      notes: fallback.used
        ? ["Счёт расчётов — по настроенным значениям по умолчанию; договор не выбран."]
        : [],
    };
  }
  const props = new Set(
    (await conn.getMetadata()).entities.get(entitySet)?.properties.map((p) => p.name) ?? [],
  );
  let rows = Array.isArray(payload[PAYMENT_ROWS]) ? (payload[PAYMENT_ROWS] as Record<string, unknown>[]) : [];
  const hasHeader = !rowSettlement && props.has(SETTLE_FIELD);
  if (!hasHeader && rows.length === 0) return { payload, notes: [] };

  const codesOf = async (...refs: Array<unknown>) => {
    const keys = refs.filter((r): r is string => typeof r === "string" && !!r);
    const accounts = await accountsByRef(conn, keys);
    return keys.map((k) => accounts.get(k)?.code ?? "?").join(" / ");
  };

  const orgKey = typeof payload["Организация_Key"] === "string" ? payload["Организация_Key"] : undefined;
  const found = await settlementAccountsFor(conn, {
    orgKey,
    contractRef,
    side,
  });
  const choice = found.published && found.side ? found.choice : undefined;
  const fallback = await strictSettlementDefaults(conn, payload, found.side, hasHeader, choice);
  payload = fallback.payload;
  if (Array.isArray(payload[PAYMENT_ROWS])) rows = payload[PAYMENT_ROWS] as Record<string, unknown>[];
  const current = [
    payload[SETTLE_FIELD] ?? rows[0]?.[SETTLE_FIELD],
    payload[ADVANCE_FIELD] ?? rows[0]?.[ADVANCE_FIELD],
  ];
  const notes: string[] = [];
  if (choice) {
    notes.push(
      `Счета расчётов ${await codesOf(choice.settle, choice.advance)} — из ${LEVEL_TEXT[choice.level]} ` +
        "(регистр «Счета учёта расчётов с контрагентами»).",
      `Источник каждого счёта (правило MCP): ${JSON.stringify(choice.provenance)}`,
    );
  } else {
    notes.push(
      `Счета расчётов — по умолчанию (${await codesOf(...current)}): ` +
        (!found.published
          ? "регистр «Счета учёта расчётов с контрагентами» не опубликован в OData, поэтому особая настройка договора не учтена."
          : found.side
            ? "в регистре «Счета учёта расчётов с контрагентами» нет подходящей записи."
            : "вид договора не «с покупателем» и не «с поставщиком» — настройка не подбирается."),
    );
  }
  if (choice && fallback.used)
    notes.push("Недостающие в настройке счета расчётов/авансов — по настроенным значениям по умолчанию.");
  const apply = (target: Record<string, unknown>, settleKey: boolean, advanceKey: boolean) =>
    choice
      ? {
          ...target,
          ...(settleKey && choice.settle ? { [SETTLE_FIELD]: choice.settle } : {}),
          ...(advanceKey && choice.advance ? { [ADVANCE_FIELD]: choice.advance } : {}),
        }
      : target;

  // Сумма взаиморасчётов в расшифровке платежа: без неё 1С проводит оплату молча и без проводок
  // (проверено на живой базе). Для договора в рублях она равна сумме платежа при курсе 1; для валюты
  // и у.е. курс не угадываем — говорим прямо.
  const rowProps = new Set(
    (await conn.getMetadata()).entities.get(`${entitySet}_${PAYMENT_ROWS}`)?.properties.map((p) => p.name) ??
      [],
  );
  const regulated = found.kind === "РасчетыВВалютеРегламентированногоУчета";
  const withSettlementSum = (r: Record<string, unknown>) =>
    regulated && rowProps.has("СуммаВзаиморасчетов") && r["СуммаВзаиморасчетов"] === undefined
      ? {
          ...r,
          СуммаВзаиморасчетов: r["СуммаПлатежа"],
          ...(rowProps.has("КурсВзаиморасчетов") ? { КурсВзаиморасчетов: 1 } : {}),
          ...(rowProps.has("КратностьВзаиморасчетов") ? { КратностьВзаиморасчетов: 1 } : {}),
        }
      : r;

  let next: Record<string, unknown> = apply(payload, hasHeader, props.has(ADVANCE_FIELD) && !rows.length);
  if (rows.length) {
    next = {
      ...next,
      [PAYMENT_ROWS]: rows.map((r) =>
        r["ДоговорКонтрагента_Key"] === contractRef ? withSettlementSum(apply(r, true, true)) : r,
      ),
    };
    if (found.kind && !regulated) {
      notes.push(
        "Договор в иностранной валюте или у.е.: сумму взаиморасчётов и курс MCP не заполняет — укажите их " +
          "в 1С до проведения, иначе 1С проведёт оплату без проводок по расчётам.",
      );
    }
  }
  return { payload: next, notes };
}

/** После записи: счёт расчётов пуст — 1С проведёт документ без проводок по расчётам, молча. */
export function emptySettlementWarning(
  created: ODataEntity,
  payload: Record<string, unknown>,
): string | undefined {
  // Без договора расчётов с контрагентом нет (налог, зарплата, прочее списание) — пустой счёт законен.
  const contract = payload["ДоговорКонтрагента_Key"];
  if (typeof contract !== "string" || !contract || contract === EMPTY_REF) return undefined;
  const empty = (v: unknown) => v === EMPTY_REF || v === "";
  const rows = Array.isArray(created[PAYMENT_ROWS]) ? (created[PAYMENT_ROWS] as ODataEntity[]) : [];
  const headerEmpty =
    Object.hasOwn(created, SETTLE_FIELD) && empty(created[SETTLE_FIELD]) && rows.length === 0;
  const rowEmpty = rows.some((r) => Object.hasOwn(r, SETTLE_FIELD) && empty(r[SETTLE_FIELD]));
  const sumEmpty = rows.some(
    (r) =>
      r["ДоговорКонтрагента_Key"] === contract &&
      Object.hasOwn(r, "СуммаВзаиморасчетов") &&
      !r["СуммаВзаиморасчетов"],
  );
  if (!headerEmpty && !rowEmpty && !sumEmpty) return undefined;
  const what = [
    headerEmpty || rowEmpty ? "пуст счёт расчётов с контрагентом" : "",
    sumEmpty ? "не заполнена сумма взаиморасчётов в расшифровке платежа" : "",
  ]
    .filter(Boolean)
    .join(" и ");
  // Подсказка — только к инструменту, который это действительно исправляет (#61): update_entity
  // табличную часть не меняет.
  const how = [
    headerEmpty || rowEmpty ? "счёт расчётов — write.document.apply_bank_settlement_settings" : "",
    sumEmpty ? `сумму взаиморасчётов — ${ALIGN_HINT}` : "",
  ]
    .filter(Boolean)
    .join("; ");
  return (
    `ВНИМАНИЕ: в созданном документе ${what} — при проведении 1С не сформирует проводки по расчётам. ` +
    `Заполните до проведения: ${how}.`
  );
}

/**
 * Перед проведением: документ с контрагентом, но без договора 1С (БП 3.0) не проводит, причём падает
 * невнятной ошибкой модуля прослеживаемости. Говорим причину заранее, понятным текстом.
 */
async function missingContract(
  conn: Connection,
  entitySet: string,
  guid: string,
): Promise<string | undefined> {
  const em = (await conn.getMetadata()).entities.get(entitySet);
  const props = new Set(em?.properties.map((p) => p.name) ?? []);
  if (!props.has("ДоговорКонтрагента_Key") || !props.has("Контрагент_Key")) return undefined;
  const doc = await conn.client.getEntity(
    `${entitySet}(guid'${guid}')${buildQuery({ select: ["Контрагент_Key", "ДоговорКонтрагента_Key"] })}`,
  );
  const empty = (v: unknown) => !v || v === EMPTY_REF;
  if (empty(doc["Контрагент_Key"]) || !empty(doc["ДоговорКонтрагента_Key"])) return undefined;
  return (
    "В документе не заполнен договор с контрагентом — 1С его не проведёт. Укажите договор " +
    '(write.entity.update_entity с {"ДоговорКонтрагента_Key": "<Ref_Key договора>"}) и проведите снова; ' +
    "если договора нет — создайте его (write.catalog.create_contract, вид уточните у пользователя)."
  );
}

/**
 * 1С отвергла значение перечисления («Перечисление 'X' не содержит элемент 'Y'») — дописываем к
 * ошибке допустимые значения из $metadata, чтобы вызывающий мог выбрать верное, а не угадывать.
 */
export async function withEnumValues(conn: Connection, e: unknown): Promise<unknown> {
  if (!(e instanceof ODataError)) return e;
  const m = /Перечисление '([^']+)' не содержит элемент/.exec(e.message);
  if (!m) return e;
  try {
    const xml = await conn.client.getText("$metadata");
    const name = m[1]!.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const body = new RegExp(`<EnumType Name="${name}"[^>]*>([\\s\\S]*?)</EnumType>`).exec(xml)?.[1];
    const values = body ? [...body.matchAll(/Member Name="([^"]+)"/g)].map((x) => x[1]) : [];
    if (!values.length) return e;
    return new ODataError({
      kind: e.kind,
      message: `${e.message} Допустимые значения ${m[1]}: ${values.join(", ")}.`,
      cause: e,
    });
  } catch {
    return e;
  }
}

/**
 * Общий путь создания: при confirm=false возвращает предпросмотр (не пишет в 1С),
 * при confirm=true выполняет POST. Гард записи (READ_ONLY + WRITABLE) — в клиенте.
 */
async function createOrPreview(
  conn: Connection,
  entitySet: string,
  payload: Record<string, unknown>,
  confirm: boolean,
  notes?: string[],
  /** keepSettlement — счета расчётов уже заданы (копия документа): настройку договора не применять. */
  opts?: { keepSettlement?: boolean; settlementSide?: SettlementSide },
) {
  const contract = await withDefaultContract(conn, entitySet, payload);
  const settlement = opts?.keepSettlement
    ? { payload: contract.payload, notes: [] }
    : await withSettlementAccounts(conn, entitySet, contract.payload, opts?.settlementSide);
  payload = await withOperationMarker(conn, entitySet, settlement.payload);
  const allNotes = [...(notes ?? []), ...contract.notes, ...settlement.notes];
  const extra = allNotes.length ? { notes: allNotes } : {};
  if (!confirm) {
    await conn.client.prepareCreate(entitySet, payload);
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      writableBase: conn.cfg.writable,
      willCreate: entitySet,
      ...(await tryPreviewView(conn, entitySet, payload)),
      payload,
      ...extra,
      note: conn.cfg.writable
        ? "Предпросмотр. Чтобы создать запись, повторите вызов с confirm=true."
        : `Предпросмотр. ВНИМАНИЕ: запись в базу "${conn.cfg.name}" сейчас запрещена — включите ODATA_DB_${conn.cfg.name.toUpperCase()}_WRITABLE=true (и READ_ONLY=false).`,
    });
  }
  let created: ODataEntity;
  try {
    created = await conn.client.create<ODataEntity>(entitySet, payload);
  } catch (e) {
    throw await withEnumValues(conn, e);
  }
  const warning = emptySettlementWarning(created, payload);
  if (warning) allNotes.push(warning);
  const createdExtra = allNotes.length ? { notes: allNotes } : {};
  return ok({
    created: true,
    ...(created["_operation_replayed"] === true
      ? {
          replayed: true,
          note: "Эта запись уже была создана с данным operationId; повторный POST не выполнялся.",
        }
      : {}),
    database: conn.cfg.name,
    entitySet,
    ref: created["Ref_Key"],
    code: created["Code"],
    description: created["Description"],
    ...createdExtra,
  });
}

/** Скалярные поля шапки из fields, опубликованные у объекта, с их типами из $metadata. */
async function scalarFields(
  conn: Connection,
  entitySet: string,
  fields: Record<string, unknown>,
): Promise<Map<string, string>> {
  const types = new Map(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => [p.name, p.type]),
  );
  return new Map(
    Object.keys(fields)
      .filter((k) => !Array.isArray(fields[k]) && types.has(k))
      .map((k) => [k, types.get(k)!]),
  );
}

/** Читает из 1С текущие значения полей (и *_Type составных ссылок). Бросает при сбое чтения. */
async function readFields(
  conn: Connection,
  entitySet: string,
  guid: string,
  keys: string[],
): Promise<Record<string, unknown>> {
  const props = new Set(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => p.name),
  );
  const select = [...keys, ...keys.map((k) => `${k.replace(/_Key$/, "")}_Type`).filter((k) => props.has(k))];
  return conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({ select })}`);
}

/**
 * Текущие значения изменяемых полей шапки — для «было → станет» в предпросмотре. Не прочитались —
 * предпросмотр без изменений, а не отказ.
 */
async function currentValues(
  conn: Connection,
  entitySet: string,
  guid: string,
  fields: Record<string, unknown>,
): Promise<Record<string, unknown> | undefined> {
  const keys = [...(await scalarFields(conn, entitySet, fields)).keys()];
  if (keys.length === 0) return undefined;
  return readFields(conn, entitySet, guid, keys).catch(() => undefined);
}

/**
 * Совпадает ли сохранённое значение с отправленным — с учётом того, как 1С возвращает типы:
 * GUID в любом регистре, дата без времени — как полночь, пустая ссылка и строка — как null.
 */
export function sameStoredValue(sent: unknown, stored: unknown, type: string): boolean {
  const empty = (v: unknown) => v === null || v === undefined || v === "" || v === EMPTY_REF;
  if (empty(sent) || empty(stored)) return empty(sent) && empty(stored);
  if (type === "Edm.Guid") return String(sent).toLowerCase() === String(stored).toLowerCase();
  if (type === "Edm.Boolean") return sent === stored;
  if (/^Edm\.(Double|Decimal|Int16|Int32|Int64|Single|Byte)$/.test(type))
    return Math.abs(Number(sent) - Number(stored)) < 1e-9;
  if (type === "Edm.DateTime") {
    const d = (v: unknown) => (/^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? `${String(v)}T00:00:00` : String(v));
    return d(sent) === d(stored);
  }
  return String(sent).trimEnd() === String(stored).trimEnd();
}

/**
 * Расчётные реквизиты банковского документа: договор и счета расчётов/авансов живут и в шапке, и в
 * строках «Расшифровки платежа», а 1С при записи берёт шапку из строк. Общий PATCH шапки даёт
 * частичную правку без согласования со строками и проводками (#60), поэтому он запрещён — и для
 * проведённого, и для непроведённого документа. Поддержанный путь — apply_bank_settlement_settings.
 */
const BANK_SETTLEMENT_FIELDS = /^(ДоговорКонтрагента_Key|СчетУчетаРасчетов[\wА-Яа-яЁё]*_Key)$/;
const BANK_DOCUMENTS = new Set<string>([...DOCUMENTS.bankIn, ...DOCUMENTS.bankOut]);

function bankSettlementPatchRefusal(entitySet: string, fields: Record<string, unknown>): string | undefined {
  if (!BANK_DOCUMENTS.has(entitySet)) return undefined;
  const blocked = Object.keys(fields).filter((k) => BANK_SETTLEMENT_FIELDS.test(k));
  if (!blocked.length) return undefined;
  return (
    `Изменение не отправлено: ${blocked.join(", ")} банковского документа общим инструментом не меняются — ` +
    "1С берёт их из строк «Расшифровки платежа», и правка шапки была бы частичной. Порядок: если документ " +
    "проведён — отмените проведение (write.document.post_document с post=false); затем договор и счета — " +
    "write.document.apply_bank_settlement_settings (шапка и строка согласованно, с проверкой результата); " +
    "затем проведите отдельным вызовом."
  );
}

/** Общий путь изменения (PATCH): dry-run при confirm=false, иначе применяет. */
async function patchOrPreview(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
  notes?: string[],
) {
  const guid = ref.replace(/[{}']/g, "");
  const extra = notes?.length ? { notes } : {};
  // Проверка не зависит от состояния документа — одинаково в предпросмотре и при подтверждении.
  const refusal = bankSettlementPatchRefusal(entitySet, fields);
  if (refusal) return fail(refusal);
  if (!confirm) {
    return ok({
      dryRun: true,
      database: conn.cfg.name,
      willPatch: `${entitySet}(${guid})`,
      ...(await tryPreviewView(conn, entitySet, fields, await currentValues(conn, entitySet, guid, fields))),
      fields,
      ...extra,
      note: conn.cfg.writable
        ? "Предпросмотр изменения. Чтобы применить, повторите с confirm=true."
        : `Предпросмотр. ВНИМАНИЕ: запись в базу "${conn.cfg.name}" запрещена — включите ODATA_DB_${conn.cfg.name.toUpperCase()}_WRITABLE=true (и READ_ONLY=false).`,
    });
  }
  const updated = await conn.client.patch<ODataEntity>(`${entitySet}(guid'${guid}')?$format=json`, fields);
  // Успех — только по фактически сохранённым значениям (#60): 1С может принять PATCH и при записи
  // объекта перезаполнить часть полей сама.
  const types = await scalarFields(conn, entitySet, fields);
  let stored: Record<string, unknown>;
  try {
    stored = types.size ? await readFields(conn, entitySet, guid, [...types.keys()]) : {};
  } catch (e) {
    return fail(
      `1С приняла изменение, но перечитать объект не удалось (${e instanceof Error ? e.message : String(e)}). ` +
        "Сверьте поля в 1С, прежде чем повторять.",
    );
  }
  const unconfirmed = [...types]
    .filter(([k, t]) => !sameStoredValue(fields[k], stored[k], t))
    .map(([k]) => k);
  if (unconfirmed.length) {
    return fail(
      `1С приняла изменение, но сохранила не всё: не совпали поля ${unconfirmed.join(", ")}` +
        `${unconfirmed.length < types.size ? " (остальные сохранены)" : ""}. ` +
        "Скорее всего, 1С перезаполнила их при записи объекта; повтор того же изменения даст тот же результат.",
    );
  }
  // verified — только когда сверено всё отправленное; табличные части этим путём не сверяются.
  const skipped = Object.keys(fields).filter((k) => !types.has(k));
  return ok({
    updated: true,
    verified: skipped.length === 0,
    database: conn.cfg.name,
    entitySet,
    ref: updated["Ref_Key"] ?? guid,
    description: updated["Description"],
    ...extra,
    ...(skipped.length ? { readbackNote: `Не сверялись (табличные части): ${skipped.join(", ")}.` } : {}),
  });
}

/**
 * Правка строк документа, которую нельзя слепо повторять (добавление/удаление строки:
 * инструмент заново читает документ, и повтор добавил бы ещё одну строку или удалил соседнюю).
 * Подтверждение идёт через журнал операций; для сверки после таймаута в журнал пишется число
 * строк до записи и ожидаемое после.
 */
async function patchLinesOrPreview(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
  check: { section: string; before: number; expected: number },
) {
  const guid = ref.replace(/[{}']/g, "");
  const key = { ref: guid };
  if (!confirm) {
    await conn.client.prepareCreate(entitySet, key, { withRefKey: false });
    return patchOrPreview(conn, entitySet, ref, fields, false);
  }
  const updated = await conn.client.patchOnce<ODataEntity>(
    `${entitySet}(guid'${guid}')?$format=json`,
    fields,
    entitySet,
    key,
    { kind: "lineCount", ref: guid, ...check },
  );
  return ok({
    updated: true,
    ...(updated["_operation_replayed"] === true
      ? {
          replayed: true,
          note: "Эта правка уже была применена с данным operationId; повторный PATCH не выполнялся.",
        }
      : {}),
    database: conn.cfg.name,
    entitySet,
    ref: updated["Ref_Key"] ?? guid,
    description: updated["Description"],
  });
}

/**
 * Создаёт подчинённый объект (банковский счёт / контактное лицо) и, при makeMain,
 * проставляет его основным у владельца. dry-run при confirm=false.
 */
async function createSubordinate(
  conn: Connection,
  set: string,
  payload: Record<string, unknown>,
  confirm: boolean,
  main?: { ownerSet: string; ownerRef: string; field: string },
) {
  if (!confirm) return createOrPreview(conn, set, payload, false);
  const created = await conn.client.create<ODataEntity>(set, payload);
  const ref = String(created["Ref_Key"] ?? "");
  if (main && ref) {
    const g = main.ownerRef.replace(/[{}']/g, "");
    await conn.client.patch(`${main.ownerSet}(guid'${g}')?$format=json`, { [main.field]: ref });
  }
  return ok({
    created: true,
    ...(created["_operation_replayed"] === true
      ? {
          replayed: true,
          note: "Эта запись уже была создана с данным operationId; повторный POST не выполнялся.",
        }
      : {}),
    database: conn.cfg.name,
    entitySet: set,
    ref,
    description: created["Description"],
    ...(main ? { setAsMain: true } : {}),
  });
}

/** Раньше этого срока после попытки операцию нельзя закрыть как неприменённую даже вручную. */
const NOT_APPLIED_MIN_AGE_SEC = 120;

export function registerWriteTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "write.operation.status",
    {
      title: "Статус операции записи",
      description:
        "Состояние операции записи по operationId из предпросмотра. После потери ответа созданный объект " +
        "сверяется по назначенной ссылке (старые операции — по метке), строки документа — по числу строк, " +
        "настройка регистра — по точному составному ключу и отпечатку проверяемых полей. Совпадение фиксирует " +
        "успех без повторной отправки. Для изменения группы, Post/Unpost и неподтверждённого результата " +
        "регистра ответ unverifiable удерживает повтор и требует ручной сверки. Отсутствие объекта само по " +
        "себе не разрешает новый запрос: 1С могла ещё выполнять его, сколько бы времени ни прошло, — ответ " +
        "not_found. Закрыть операцию как not_applied можно только с markNotApplied=true после того, как " +
        "человек сам сверил базу. Новый предпросмотр допустим после not_applied либо rejected с устранённой " +
        "причиной. Ничего не записывает в 1С.",
      inputSchema: z.object({
        database: databaseField,
        operationId: z.string().uuid().describe("operationId из предпросмотра создания"),
        markNotApplied: z
          .boolean()
          .optional()
          .describe(
            "Только по явному решению человека, который сам сверил базу 1С и убедился, что записи нет. " +
              "Сверка всё равно выполняется: найденный объект фиксируется как успех. Не раньше чем через " +
              `${NOT_APPLIED_MIN_AGE_SEC} с после попытки.`,
          ),
      }),
      outputSchema: operationStatusSchema,
    },
    ({ database, operationId, markNotApplied }) =>
      guard("write.operation.status", async () => {
        const conn = ctx.db(database);
        const base = { operationId, database: conn.cfg.name };
        const entry = await conn.client.operationEntry(operationId);
        if (!entry) {
          return ok({
            ...base,
            status: "not_in_journal",
            note: "Операция не найдена в локальном журнале этой базы: подтверждение не выполнялось.",
          });
        }
        const common = {
          ...base,
          entitySet: entry.entitySet,
          // Причина неизвестного исхода или отказа — текст 1С из журнала (#55).
          ...(entry.errorMessage && entry.state !== "succeeded" ? { lastError: entry.errorMessage } : {}),
        };
        if (entry.state === "succeeded") {
          return ok({
            ...common,
            status: "succeeded",
            ref: entry.result?.["Ref_Key"],
            number: entry.result?.["Number"],
            ...(entry.result?.["RecordPath"] ? { recordPath: entry.result["RecordPath"] } : {}),
          });
        }
        if (entry.state === "prepared") {
          return ok({
            ...common,
            status: "prepared",
            note: "Предпросмотр был, подтверждение не отправлялось.",
          });
        }
        if (entry.state === "rejected") {
          return ok({
            ...common,
            status: "rejected",
            note: "1С отклонила запись. Нужен новый предпросмотр после устранения причины отказа.",
          });
        }
        if (entry.check?.kind === "manual")
          return ok({
            ...common,
            status: "unverifiable",
            note: "Исход изменения группы контрагента неизвестен. Сверьте Parent_Key и версию в 1С; повторная отправка удерживается.",
          });
        if (entry.check?.kind === "registerValues") {
          const check = entry.check;
          try {
            const record = await conn.client.getEntity(`${check.recordPath}?$format=json`);
            if (fingerprintRecordFields(record, check.fields) === check.expectedHash) {
              await conn.client.reconcileOperation(operationId, { RecordPath: check.recordPath });
              return ok({
                ...common,
                status: "found_reconciled",
                recordPath: check.recordPath,
                note: "Точная запись регистра и проверяемые ресурсы совпали с ожидаемыми; результат зафиксирован без повторной отправки.",
              });
            }
          } catch {
            // Missing records and failed reads cannot prove non-application of a delayed request.
          }
          return ok({
            ...common,
            status: "unverifiable",
            recordPath: check.recordPath,
            note: "Точная запись регистра отсутствует, отличается или не прочитана. Требуется ручная сверка; повторная отправка удерживается.",
          });
        }
        if (entry.check?.kind === "createdRef") {
          try {
            const actual = await conn.client.getEntity(
              `${entry.entitySet}(${odataGuid(entry.check.ref)})?$format=json`,
            );
            if (String(actual.Ref_Key).toLowerCase() !== entry.check.ref.toLowerCase())
              return ok({
                ...common,
                status: "unverifiable",
                note: "Прочитана другая ссылка; повтор удерживается.",
              });
            await conn.client.reconcileOperation(operationId, { Ref_Key: entry.check.ref });
            return ok({ ...common, status: "found_reconciled", ref: entry.check.ref });
          } catch (error) {
            if (!(error instanceof ODataError && error.kind === "not_found")) throw error;
            return ok({
              ...common,
              status: "unverifiable",
              note: "Группа по назначенной ссылке пока не найдена; запрос мог ещё выполняться. Повторная запись удерживается.",
            });
          }
        }
        if (entry.check?.kind === "documentState")
          return ok({
            ...common,
            status: "unverifiable",
            note: "Исход изменения документа неизвестен. Нужна ручная сверка версии, полей и состояния; повторная запись удерживается.",
          });
        if (entry.entitySet.startsWith("InformationRegister_"))
          return ok({
            ...common,
            status: "unverifiable",
            note: "Исход записи настройки неизвестен. Сверьте точный составной ключ и ресурсы вручную; повторная отправка заблокирована.",
          });
        const ageSec = Math.round((Date.now() - Date.parse(entry.updatedAt)) / 1000);
        // Отсутствие следа записи не доказывает, что запрос не применился: клиент обрывает ожидание по
        // таймауту, а 1С может закончить запись позже. Поэтому операция закрывается как not_applied только
        // по явному решению человека после ручной сверки (issue #52), и не сразу после попытки.
        const absent = async (what: string, ref?: string) => {
          if (markNotApplied && ageSec >= NOT_APPLIED_MIN_AGE_SEC) {
            await conn.client.markOperationNotApplied(operationId);
            return ok({
              ...common,
              status: "not_applied",
              ...(ref ? { ref } : {}),
              note: `${what} Операция закрыта как неприменённая по ручному подтверждению; для новой записи выполните новый предпросмотр.`,
            });
          }
          return ok({
            ...common,
            status: "not_found",
            ...(ref ? { ref } : {}),
            note:
              `${what} С последней попытки ${ageSec} с. ` +
              (markNotApplied
                ? `Закрыть как неприменённую можно не раньше чем через ${NOT_APPLIED_MIN_AGE_SEC} с — повторите позже. `
                : "Это не доказывает, что запись не дошла: 1С могла ещё выполнять запрос. Не создавайте запись заново. ") +
              "Сверьте базу 1С вручную; если записи точно нет, вызовите write.operation.status с markNotApplied=true.",
          });
        };
        // Правка строк документа — сверяем по числу строк до/после записи.
        if (entry.check?.kind === "lineCount") {
          const { ref, before, expected } = entry.check;
          const count = (await getDocInfo(conn, entry.entitySet, ref)).lines.length;
          if (count === expected && count !== before) {
            await conn.client.reconcileOperation(operationId, { Ref_Key: ref });
            return ok({
              ...common,
              status: "found_reconciled",
              ref,
              note: `В документе ${count} строк(и), как и ожидалось после правки: операция зафиксирована как успешная.`,
            });
          }
          if (count === before)
            return absent(`Правка не видна: в документе по-прежнему ${count} строк(и).`, ref);
          return ok({
            ...common,
            status: "unverifiable",
            ref,
            note: `Число строк ${count} не совпадает ни с исходным (${before}), ни с ожидаемым (${expected}): документ меняли. Проверьте его в 1С вручную.`,
          });
        }
        // Создание с назначенным Ref_Key — сверка простым GET по ссылке (справочники тоже).
        if (entry.refKey) {
          let exists: boolean;
          try {
            await conn.client.getEntity(
              `${entry.entitySet}(guid'${entry.refKey}')?$format=json&$select=Ref_Key`,
            );
            exists = true;
          } catch (e) {
            if (!(e instanceof ODataError && e.kind === "not_found")) throw e;
            exists = false;
          }
          if (exists) {
            await conn.client.reconcileOperation(operationId, { Ref_Key: entry.refKey });
            return ok({
              ...common,
              status: "found_reconciled",
              ref: entry.refKey,
              note: "Объект с назначенной ссылкой есть в 1С; операция зафиксирована как успешная. Повторное подтверждение вернёт эту ссылку.",
            });
          }
          return absent("Объекта с назначенной ссылкой в 1С нет.");
        }
        // executing | outcome_unknown, старые операции без Ref_Key — сверяем с 1С по метке
        if (!(await entityHasComment(conn, entry.entitySet))) {
          return ok({
            ...common,
            status: "unverifiable",
            note: "У этого объекта нет «Комментария», метку не поставить. Проверьте в 1С вручную (по наименованию/номеру).",
          });
        }
        // 1С превращает substringof в регулярное выражение — скобки метки «[op:…]» ломают запрос (500),
        // поэтому ищем по «op:<id>» без них (UUID — только hex и «-»).
        // Number есть только у документов; у справочников — Code (иначе 1С отвечает 400).
        const numberField = entry.entitySet.startsWith("Document_") ? "Number" : "Code";
        const found = await conn.client.getCollection<ODataEntity>(
          `${entry.entitySet}?$format=json&$top=2&$select=Ref_Key,${numberField}&$filter=${encodeURIComponent(
            `substringof('op:${operationId}',Комментарий)`,
          )}`,
        );
        const hit = found.value[0];
        if (hit) {
          await conn.client.reconcileOperation(operationId, hit);
          return ok({
            ...common,
            status: "found_reconciled",
            ref: hit["Ref_Key"],
            number: hit[numberField],
            note: "Объект найден по метке; операция зафиксирована как успешная. Повторное подтверждение вернёт эту ссылку.",
          });
        }
        return absent("По метке объект не найден.");
      }),
  );

  server.registerTool(
    "write.counterparty.create_counterparty",
    {
      title: "Создать контрагента",
      description:
        "Создаёт нового контрагента в справочнике. ПО УМОЛЧАНИЮ это предпросмотр (dry-run) — " +
        "сначала покажите пользователю, что будет создано, и только после явного согласия " +
        "вызовите повторно с confirm=true. Запись идёт в боевую базу, поэтому без подтверждения не создавайте.",
      inputSchema: z.object({
        database: databaseField,
        name: z.string().min(1).describe("Наименование контрагента (Description)"),
        inn: z.string().optional().describe("ИНН"),
        kpp: z.string().optional().describe("КПП"),
        fullName: z.string().optional().describe("Полное наименование"),
        legalType: z
          .enum(["ЮридическоеЛицо", "ФизическоеЛицо"])
          .optional()
          .describe("Юридическое или физическое лицо"),
        phone: z.string().optional().describe("Телефон (в контактную информацию)"),
        email: z.string().optional().describe("Email (в контактную информацию)"),
        address: z.string().optional().describe("Адрес текстом (в контактную информацию, юр. адрес)"),
        ogrn: z.string().optional().describe("ОГРН (пишется, если в базе настроен доп.реквизит «ОГРН»)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, name, inn, kpp, fullName, legalType, phone, email, address, ogrn, confirm }) =>
      guard("write.counterparty.create_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Контрагенты");
        const { fields, notes } = await counterpartyExtras(conn, { phone, email, address, ogrn });
        const payload = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
          ЮридическоеФизическоеЛицо: legalType,
          ...fields,
        });
        return createOrPreview(conn, set, payload, confirm, notes);
      }),
  );

  server.registerTool(
    "write.entity.mark_for_deletion",
    {
      title: "Пометить на удаление",
      description:
        "Ставит/снимает пометку на удаление у объекта (как в 1С) — мягкое удаление, " +
        "физически данные не стираются. По умолчанию предпросмотр (dry-run); для применения " +
        "повторите с confirm=true. Полезно, чтобы убрать ошибочно созданную запись.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя объекта, напр. Catalog_Контрагенты"),
        ref: z.string().describe("Ref_Key объекта (GUID)"),
        mark: z.boolean().default(true).describe("true — пометить на удаление, false — снять пометку"),
        confirm: confirmField,
      }),
      outputSchema: markForDeletionResultSchema,
    },
    ({ database, entitySet, ref, mark, confirm }) =>
      guard("write.entity.mark_for_deletion", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const guid = ref.replace(/[{}']/g, "");
        const path = `${entitySet}(guid'${guid}')?$format=json`;
        if (!confirm) {
          return ok({
            dryRun: true,
            database: conn.cfg.name,
            willPatch: `${entitySet}(${guid})`,
            payload: { DeletionMark: mark },
            note: "Предпросмотр. Чтобы применить пометку, повторите с confirm=true.",
          });
        }
        const updated = await conn.client.patch<ODataEntity>(path, { DeletionMark: mark });
        return ok({
          updated: true,
          database: conn.cfg.name,
          ref: updated["Ref_Key"] ?? guid,
          deletionMark: updated["DeletionMark"] ?? mark,
        });
      }),
  );

  server.registerTool(
    "write.catalog.create_nomenclature",
    {
      title: "Создать номенклатуру",
      description:
        "Создаёт новую позицию номенклатуры (товар/услуга). ПО УМОЛЧАНИЮ предпросмотр (dry-run): " +
        "покажите пользователю payload и только после согласия повторите с confirm=true. " +
        "Можно указать артикул, папку (folder — имя группы или Ref_Key) и пометить услугой (isService). " +
        "Запись идёт в боевую базу.",
      inputSchema: z.object({
        database: databaseField,
        name: z.string().min(1).describe("Наименование номенклатуры (Description)"),
        fullName: z.string().optional().describe("Полное наименование"),
        article: z.string().optional().describe("Артикул"),
        folder: z.string().optional().describe("Папка/группа — имя или Ref_Key (для размещения в группе)"),
        isService: z
          .boolean()
          .default(false)
          .describe("true — услуга (Услуга=true, вид «Услуга»), false — товар"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, name, fullName, article, folder, isService, confirm }) =>
      guard("write.catalog.create_nomenclature", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.nomenclature, "Номенклатура");
        const parentRef = await folderRefOf(conn, set, folder);
        const payload = clean({
          Description: name,
          НаименованиеПолное: fullName,
          Артикул: article,
          Parent_Key: parentRef,
          ...(isService ? { Услуга: true } : {}),
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.catalog.create_contract",
    {
      title: "Создать договор",
      description:
        "Заводит договор контрагента (подчинённый справочник): владелец — контрагент, " +
        "организация, ВИД договора, номер/дата, валюта/тип цен и (опц.) руководитель/подписант. " +
        "ВАЖНО: вид договора (kind) обязателен и его НЕЛЬЗЯ выбирать самостоятельно — " +
        "уточните у пользователя, т.к. один контрагент может быть и покупателем, и поставщиком. " +
        "Номер по умолчанию «б/н», дата — сегодня. По умолчанию предпросмотр (dry-run); создание — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента (владелец договора)"),
        kind: z
          .enum(CONTRACT_KINDS)
          .describe("Вид договора — ОБЯЗАТЕЛЬНО спросить у пользователя, не выбирать самому"),
        number: z.string().optional().describe("Номер договора (по умолчанию «б/н»)"),
        date: z.string().optional().describe("Дата договора YYYY-MM-DD (по умолчанию сегодня)"),
        name: z.string().optional().describe("Наименование договора (по умолчанию «Договор <номер>»)"),
        currency: z
          .string()
          .optional()
          .describe("Валюта взаиморасчётов — код (напр. 643/840) или название (RUB/USD)"),
        priceType: z
          .string()
          .optional()
          .describe("Тип цен — название или код (из справочника Типы цен номенклатуры)"),
        headName: z
          .string()
          .optional()
          .describe("ФИО руководителя/подписанта контрагента (реквизит договора)"),
        headPosition: z
          .string()
          .optional()
          .describe("Должность руководителя контрагента (напр. «Генеральный директор»)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      kind,
      number,
      date,
      name,
      currency,
      priceType,
      headName,
      headPosition,
      confirm,
    }) =>
      guard("write.catalog.create_contract", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.contracts, "Договоры контрагентов");
        const org = await resolveOrg(conn, organization);

        let cur: { ref: string; code?: string } | undefined;
        if (currency)
          cur = await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", currency);
        const pt = priceType
          ? await resolveCatalogItem(
              conn,
              CATALOGS.priceTypes,
              "Справочник «Типы цен номенклатуры»",
              priceType,
            )
          : undefined;
        if (cur && !cur.code?.trim())
          throw new InputError("Код выбранной валюты неизвестен: нельзя определить валютность договора.");
        const foreign = cur ? cur.code!.trim() !== (conn.behavior.regulatedCurrencyCode ?? "643") : undefined;
        const num = number ?? "б/н";

        const payload = clean({
          Description: name ?? `Договор ${num}`,
          Номер: num,
          Дата: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Owner_Key: counterpartyRef,
          ВидДоговора: kind,
          Организация_Key: org.key,
          ВалютаВзаиморасчетов_Key: cur?.ref,
          ТипЦен_Key: pt?.ref,
          РуководительКонтрагента: headName,
          ДолжностьРуководителяКонтрагента: headPosition,
          ...(foreign !== undefined ? { Валютный: foreign } : {}),
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_invoice",
    {
      title: "Создать счёт покупателю",
      description:
        "Создаёт счёт на оплату покупателю (документ) с табличной частью «Товары». " +
        "Документ создаётся НЕПРОВЕДЁННЫМ (черновик) — провести можно вручную в 1С или инструментом post_document. " +
        "Для услуг задавайте content в строке — это «Содержание», именно оно печатается в счёте. " +
        "По умолчанию предпросмотр (dry-run); создание — при confirm=true. " +
        "Контрагент и (опц.) договор — по Ref_Key; позиции — по Ref_Key номенклатуры.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора (необязательно)"),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        orgBankAccount: orgBankAccountField,
        // Счёт на оплату проводок не делает — счета учёта в его строках не нужны.
        lines: z.array(purchaseLine).min(1).describe("Позиции счёта"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      date,
      sumIncludesVat,
      orgBankAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.customerInvoice,
          "Документ «Счёт на оплату покупателю»",
        );
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const rows = buildInvoiceRows(lines);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          // У счёта на оплату банковский счёт организации зовётся «Структурная единица».
          СтруктурнаяЕдиница_Key: bank,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.document.post_document",
    {
      title: "Провести / отменить проведение документа",
      description:
        "Проводит документ (влияет на учёт) или отменяет проведение. По умолчанию предпросмотр (dry-run); " +
        "Для повторного проведения Posted=true задайте repost=true, post=true и expectedDataVersion. " +
        "Версионный Post использует PostingModeOperational=false без Unpost; проверяет состояние документа, не проводки. " +
        "применение — при confirm=true. ВНИМАНИЕ: проведение меняет учётные данные — используйте осознанно.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_СчетНаОплатуПокупателю"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        expectedDataVersion: expectedDocumentVersion
          .optional()
          .describe(
            "Точная текущая версия включает durable preview/action/readback. Post/Unpost выполняются отдельно; атомарная CAS действия не подтверждена.",
          ),
        post: z.boolean().default(true).describe("true — провести, false — отменить проведение"),
        repost: z
          .boolean()
          .optional()
          .describe(
            "Явно разрешить повторный Post уже проведённого документа; требует post=true и expectedDataVersion",
          ),
        confirm: confirmField,
      }),
      outputSchema: postDocumentResultSchema,
    },
    ({ database, entitySet, ref, post, repost, confirm, expectedDataVersion }) =>
      guard("write.document.post_document", async () => {
        if (repost && (!post || expectedDataVersion === undefined))
          throw new InputError("repost=true требует post=true и точную expectedDataVersion.");
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        // Суммы расшифровки банковского документа — и в предпросмотре: ошибку видно до подтверждения (#61).
        if (post) {
          const blocker = await bankPostingBlocker(conn, entitySet, ref.replace(/[{}']/g, ""));
          if (blocker) return fail(blocker);
        }
        if (expectedDataVersion !== undefined) {
          if (post && confirm) {
            const missing = await missingContract(conn, entitySet, ref);
            if (missing) return fail(missing);
          }
          return ok(
            await guardedDocumentAction(conn, { entitySet, ref, post, repost, expectedDataVersion }, confirm),
          );
        }
        const guid = ref.replace(/[{}']/g, "");
        const action = post ? "Post" : "Unpost";
        const path = `${entitySet}(guid'${guid}')/${action}?$format=json`;
        if (!confirm) {
          return ok({
            dryRun: true,
            database: conn.cfg.name,
            willCall: `${entitySet}(${guid})/${action}`,
            note: "Предпросмотр. Чтобы применить, повторите с confirm=true.",
          });
        }
        if (post) {
          const missing = await missingContract(conn, entitySet, guid);
          if (missing) return fail(missing);
        }
        // Успех — только по состоянию документа: 1С может ответить 200 и не провести (ошибка в
        // обработке проведения). Повторный Post проведённого документа подтверждает смена DataVersion.
        const statePath = `${entitySet}(guid'${guid}')?$format=json&$select=Posted,DataVersion`;
        const state = () => conn.client.getEntity<{ Posted?: boolean; DataVersion?: string }>(statePath);
        const before = await state();
        const response = await conn.client.action(path);
        let after: Awaited<ReturnType<typeof state>>;
        try {
          after = await state();
        } catch (e) {
          return fail(
            `1С приняла ${action}, но перечитать документ не удалось (${e instanceof Error ? e.message : String(e)}). ` +
              "Проверьте в 1С, проведён ли документ, прежде чем повторять.",
          );
        }
        const applied = post
          ? after.Posted === true && (before.Posted !== true || after.DataVersion !== before.DataVersion)
          : after.Posted === false;
        if (!applied) {
          const body = response === undefined ? "" : JSON.stringify(response).slice(0, 500);
          return fail(
            `1С ответила на ${action} без ошибки, но документ ${post ? "не проведён" : "остался проведённым"} ` +
              `(Posted=${String(after.Posted)}${post && before.Posted === true ? ", версия не изменилась" : ""}). ` +
              "Причину покажет проведение документа в самой 1С." +
              (body ? ` Ответ 1С: ${body}` : ""),
          );
        }
        // Что провёл документ — сразу в ответе: проводки по регистру Хозрасчетный.
        // Сбой чтения проводок не отменяет проведения, только поясняется.
        let postings: Record<string, unknown> = {};
        if (post) {
          try {
            const p = await getDocumentPostings(conn, entitySet, guid);
            postings = {
              postings: {
                count: p.postingsCount,
                debitTotal: p.debitTotal,
                creditTotal: p.creditTotal,
                byCorrespondence: p.byCorrespondence.slice(0, 20),
              },
              ...(p.postingsCount === 0
                ? { postingsNote: "Документ проведён, но проводок по бухгалтерскому учёту не сформировал." }
                : {}),
            };
          } catch (e) {
            postings = {
              postingsNote: `Проведено; проводки прочитать не удалось: ${e instanceof Error ? e.message : String(e)}`,
            };
          }
        }
        return ok({
          done: true,
          verified: true,
          verificationScope: "document_state",
          database: conn.cfg.name,
          ref: guid,
          action,
          ...postings,
        });
      }),
  );

  server.registerTool(
    "write.entity.update_entity",
    {
      title: "Изменить объект (общий)",
      description:
        "Изменяет произвольные поля объекта 1С через PATCH по Ref_Key. Имена полей — технические " +
        '(узнать через describe_entity), напр. {"ИНН":"7701234567"}. По умолчанию предпросмотр (dry-run); ' +
        "применение — при confirm=true. Мощный инструмент: меняйте только те поля, что указали. " +
        "После записи поля перечитываются: успех — только если 1С сохранила всё. Договор и счета расчётов " +
        "банковских документов этим инструментом не меняются — write.document.apply_bank_settlement_settings.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя объекта, напр. Catalog_Контрагенты"),
        ref: z.string().describe("Ref_Key объекта (GUID)"),
        fields: z
          .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
          .describe("Поля для изменения: { техническоеИмя: значение }"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, fields, confirm }) =>
      guard("write.entity.update_entity", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        if (Object.keys(fields).length === 0) return fail("Не заданы поля для изменения.");
        return patchOrPreview(conn, entitySet, ref, fields, confirm);
      }),
  );

  server.registerTool(
    "write.counterparty.update_counterparty",
    {
      title: "Изменить контрагента",
      description:
        "Изменяет реквизиты контрагента по Ref_Key (наименование/ИНН/КПП/полное имя). " +
        "По умолчанию предпросмотр (dry-run); применение — при confirm=true. Указывайте только меняемые поля.",
      inputSchema: z.object({
        database: databaseField,
        ref: z.string().describe("Ref_Key контрагента"),
        name: z.string().optional().describe("Новое наименование (Description)"),
        inn: z.string().optional().describe("ИНН"),
        kpp: z.string().optional().describe("КПП"),
        fullName: z.string().optional().describe("Полное наименование"),
        phone: z.string().optional().describe("Телефон (контактная информация)"),
        email: z.string().optional().describe("Email (контактная информация)"),
        address: z.string().optional().describe("Адрес текстом (юр. адрес)"),
        ogrn: z.string().optional().describe("ОГРН (если в базе настроен доп.реквизит)"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, ref, name, inn, kpp, fullName, phone, email, address, ogrn, confirm }) =>
      guard("write.counterparty.update_counterparty", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.counterparties, "Справочник «Контрагенты»");
        const fields: Record<string, unknown> = clean({
          Description: name,
          ИНН: inn,
          КПП: kpp,
          НаименованиеПолное: fullName,
        });
        const notes: string[] = [];
        // Контакты — со слиянием: заменяем только заданные виды, прочие сохраняем.
        if (phone || email || address) {
          const kinds = await contactKindsForCounterparties(conn);
          const guid = ref.replace(/[{}']/g, "");
          const existing =
            ((await conn.client.getEntity(`${set}(guid'${guid}')${buildQuery({})}`))[
              "КонтактнаяИнформация"
            ] as Array<Record<string, unknown>>) ?? [];
          const replacedTypes = new Set(
            [phone && kinds.phone?.тип, email && kinds.email?.тип, address && kinds.address?.тип].filter(
              Boolean,
            ),
          );
          const kept = existing.filter((r) => !replacedTypes.has(String(r["Тип"])));
          const merged = [...kept, ...buildContactRows(kinds, { phone, email, address })].map((r, i) => ({
            ...r,
            LineNumber: i + 1,
          }));
          fields["КонтактнаяИнформация"] = merged;
        }
        if (ogrn) {
          const prop = await resolveAdditionalProperty(conn, "ОГРН");
          if (prop)
            fields["ДополнительныеРеквизиты"] = [{ LineNumber: 1, Свойство_Key: prop, Значение: ogrn }];
          else notes.push("ОГРН не записан: в базе нет доп.реквизита «ОГРН» для контрагентов.");
        }
        if (Object.keys(fields).length === 0) return fail("Не задано ни одного поля для изменения.");
        return patchOrPreview(conn, set, ref, fields, confirm, notes);
      }),
  );

  server.registerTool(
    "write.catalog.update_nomenclature",
    {
      title: "Изменить номенклатуру",
      description:
        "Изменяет реквизиты номенклатуры по Ref_Key (наименование/полное имя/артикул). " +
        "По умолчанию предпросмотр (dry-run); применение — при confirm=true. Указывайте только меняемые поля.",
      inputSchema: z.object({
        database: databaseField,
        ref: z.string().describe("Ref_Key номенклатуры"),
        name: z.string().optional().describe("Новое наименование (Description)"),
        fullName: z.string().optional().describe("Полное наименование"),
        article: z.string().optional().describe("Артикул"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, ref, name, fullName, article, confirm }) =>
      guard("write.catalog.update_nomenclature", async () => {
        const conn = ctx.db(database);
        const set = await resolveSet(conn, CATALOGS.nomenclature, "Справочник «Номенклатура»");
        const fields = clean({ Description: name, НаименованиеПолное: fullName, Артикул: article });
        if (Object.keys(fields).length === 0) return fail("Не задано ни одного поля для изменения.");
        return patchOrPreview(conn, set, ref, fields, confirm);
      }),
  );

  // Общая схема входов для товарных документов (поступление/реализация).
  const goodsDocCommon = {
    database: databaseField,
    organization: organizationField,
    counterpartyRef: z.string().describe("Ref_Key контрагента"),
    contractRef: z.string().optional().describe("Ref_Key договора (необязательно)"),
    warehouse: z.string().optional().describe("Название склада (если в базе несколько)"),
    date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
    sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
  };

  /** Схема закупки: счета доходов/расходов не применяются (Дт 41/19 Кт 60). */
  const purchaseDocInput = {
    ...goodsDocCommon,
    lines: z.array(purchaseLine).min(1).describe("Позиции документа"),
    confirm: confirmField,
  };

  /** Схема продажи: счета доходов/расходов — на документ и (приоритетнее) на строку. */
  const salesDocInput = {
    ...goodsDocCommon,
    incomeAccount: incomeAccountField.describe(
      "Счёт доходов для ВСЕХ строк документа — код как в 1С (напр. «90.01.2») или Ref_Key. " +
        "Счёт, заданный в строке, важнее. Без указания — из регистра «Счета учёта номенклатуры», иначе 90.01.1.",
    ),
    expenseAccount: expenseAccountField.describe(
      "Счёт расходов для ВСЕХ строк документа — код как в 1С (напр. «90.02.2») или Ref_Key. " +
        "Счёт, заданный в строке, важнее. Без указания — из регистра «Счета учёта номенклатуры», иначе 90.02.1.",
    ),
    lines: z.array(saleLine).min(1).describe("Позиции документа"),
    confirm: confirmField,
  };

  /** Реализация: та же схема плюс банковский счёт организации (у возврата его нет). */
  const shipmentDocInput = { ...salesDocInput, orgBankAccount: orgBankAccountField };

  server.registerTool(
    "write.purchase.create_purchase",
    {
      title: "Создать поступление от поставщика",
      description:
        "Создаёт документ «Поступление товаров и услуг» (закупка у поставщика) с табличной частью «Товары». " +
        "Документ НЕПРОВЕДЁННЫЙ; провести — вручную в 1С или post_document (тогда 1С сформирует проводки Дт 41/19 Кт 60; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "Позиции могут нести content (Содержание строки) — для услуг вроде доставки. " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — поставщик, договор — вида «СПоставщиком».",
      inputSchema: z.object(purchaseDocInput),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      sumIncludesVat,
      lines,
      confirm,
    }) =>
      guard("write.purchase.create_purchase", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.purchases, "Документ «Поступление товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "purchase");
        return createGoodsDoc(
          conn,
          set,
          {
            orgKey: org.key,
            counterpartyRef,
            contractRef,
            warehouseKey,
            date,
            sumIncludesVat,
            lines,
            settlement,
            lineAccountsFor,
          },
          confirm,
        );
      }),
  );

  server.registerTool(
    "write.purchase.create_supplier_invoice",
    {
      title: "Зарегистрировать счёт от поставщика",
      description:
        "Создаёт документ «Счёт на оплату поставщика» — регистрирует входящий счёт-основание для оплаты " +
        "(подходит для предоплаты, когда товар ещё не пришёл). Проводок НЕ делает (счёт не проводится), " +
        "счета учёта не нужны. Табличная часть «Товары» вмещает и товары, и услуги (напр. доставку — задайте " +
        "content/Содержание). Реквизиты входящего счёта: incomingNumber/incomingDate. " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — поставщик, договор — вида «СПоставщиком».",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-поставщика"),
        contractRef: z.string().optional().describe("Ref_Key договора (вид «СПоставщиком»)"),
        supplierBankAccountRef: z
          .string()
          .optional()
          .describe("Ref_Key банковского счёта поставщика (БанковскийСчетКонтрагента)"),
        incomingNumber: z.string().optional().describe("Номер входящего счёта поставщика"),
        incomingDate: z.string().optional().describe("Дата входящего счёта YYYY-MM-DD"),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        comment: z.string().optional().describe("Комментарий к документу"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        lines: z.array(purchaseLine).min(1).describe("Позиции счёта (товары и услуги в одной таблице)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      supplierBankAccountRef,
      incomingNumber,
      incomingDate,
      date,
      comment,
      sumIncludesVat,
      lines,
      confirm,
    }) =>
      guard("write.purchase.create_supplier_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.supplierInvoice,
          "Документ «Счёт на оплату поставщика»",
        );
        const org = await resolveOrg(conn, organization);
        const rows = buildGoodsRows(lines, noAccounts);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          БанковскийСчетКонтрагента_Key: supplierBankAccountRef,
          НомерВходящегоДокумента: incomingNumber,
          ДатаВходящегоДокумента: incomingDate ? odataDate(new Date(`${incomingDate}T00:00:00`)) : undefined,
          Комментарий: comment,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.money.create_payout_order",
    {
      title: "Создать платёжное поручение (исходящее) контрагенту",
      description:
        "Создаёт документ «Платёжное поручение» — исходящую выплату контрагенту (напр. оплата " +
        "по агентскому договору / услугам, зеркало платёжки, уже ушедшей в банк). Документ ВСЕГДА " +
        "создаётся НЕПРОВЕДЁННЫМ (Posted=false) — эта операция намеренно не проводит документ, " +
        "проведение делайте отдельно через post_document, если вообще нужно. Номер (number) — " +
        "задаётся явно вызывающим (обычно чтобы совпасть с номером уже отправленной банковской " +
        "платёжки), не генерируется автоматически. По умолчанию dry-run (confirm=false); создание — " +
        "при confirm=true. Счёт организации и валюта — резолвятся автоматически, если не заданы.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key контрагента-получателя (из find_counterparty)"),
        counterpartyAccountRef: z
          .string()
          .optional()
          .describe("Ref_Key банковского счёта контрагента (СчетКонтрагента_Key)"),
        contractRef: z
          .string()
          .optional()
          .describe("Ref_Key договора с контрагентом (ДоговорКонтрагента_Key)"),
        amount: z.number().positive().describe("Сумма документа"),
        number: z
          .string()
          .describe(
            "Номер документа — задайте явно (напр. чтобы совпасть с номером банковской платёжки). " +
              "1С автонумерует новые документы только если Number не передан вовсе — но здесь он обязателен.",
          ),
        date: z.string().optional().describe("Дата документа YYYY-MM-DD (по умолчанию сегодня)"),
        purposeText: z.string().describe("Назначение платежа (НазначениеПлатежа)"),
        recipientText: z.string().describe("Текст получателя — как печатается в платёжке (ТекстПолучателя)"),
        recipientInn: z.string().optional().describe("ИНН получателя"),
        recipientKpp: z.string().optional().describe("КПП получателя"),
        payerText: z.string().optional().describe("Текст плательщика (по умолчанию — имя организации)"),
        payerInn: z.string().optional().describe("ИНН плательщика (по умолчанию — ИНН организации)"),
        orgBankAccount: orgBankAccountField,
        responsibleRef: z.string().optional().describe("Ref_Key ответственного (Ответственный_Key)"),
        cashflowItemRef: z
          .string()
          .optional()
          .describe("Ref_Key статьи движения денежных средств (СтатьяДвиженияДенежныхСредств_Key)"),
        currencyRef: z.string().optional().describe("Ref_Key валюты документа; без указания — рубль (643)"),
        vatRate: z.enum(VAT_RATES).default("БезНДС").describe("Ставка НДС"),
        operationKind: z
          .string()
          .default("ОплатаПоставщику")
          .describe("Вид операции (ВидОперации), напр. ОплатаПоставщику"),
        priority: z.number().int().default(5).describe("Очередность платежа (ОчередностьПлатежа)"),
        comment: z.string().optional().describe("Комментарий к документу"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      counterpartyAccountRef,
      contractRef,
      amount,
      number,
      date,
      purposeText,
      recipientText,
      recipientInn,
      recipientKpp,
      payerText,
      payerInn,
      orgBankAccount,
      responsibleRef,
      cashflowItemRef,
      currencyRef,
      vatRate,
      operationKind,
      priority,
      comment,
      confirm,
    }) =>
      guard("write.money.create_payout_order", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.paymentOrder, "Документ «Платёжное поручение»");
        const org = await resolveOrgOrDefault(conn, organization);
        const [resolvedOrgAccount, resolvedCurrency] = await Promise.all([
          // Через общий резолвер: принимает название/номер счёта, а не только Ref_Key.
          resolveOrgBankAccount(conn, org.ref, orgBankAccount),
          currencyRef ? Promise.resolve(currencyRef) : resolveDefaultCurrency(conn),
        ]);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Number: number,
          Организация_Key: org.ref,
          СчетОрганизации_Key: resolvedOrgAccount,
          Контрагент: counterpartyRef,
          Контрагент_Type: COUNTERPARTY_TYPE,
          СчетКонтрагента_Key: counterpartyAccountRef,
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          СтавкаНДС: vatRate,
          ВидОперации: operationKind,
          ОчередностьПлатежа: priority,
          НазначениеПлатежа: purposeText,
          ТекстПолучателя: recipientText,
          ИННПолучателя: recipientInn,
          КПППолучателя: recipientKpp,
          ТекстПлательщика: payerText ?? org.name,
          ИННПлательщика: payerInn ?? org.inn,
          Ответственный_Key: responsibleRef,
          СтатьяДвиженияДенежныхСредств_Key: cashflowItemRef,
          ВалютаДокумента_Key: resolvedCurrency,
          Комментарий: comment,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_shipment",
    {
      title: "Создать реализацию покупателю",
      description:
        "Создаёт документ «Реализация товаров и услуг» (отгрузка покупателю) с табличной частью «Товары». " +
        "Документ НЕПРОВЕДЁННЫЙ; провести — вручную в 1С или post_document (тогда 1С сформирует проводки Дт 62 Кт 90, Дт 90 Кт 41 и др.; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "Позиции могут нести content (Содержание строки — печатается в УПД) и свои счета доходов/расходов. " +
        "По умолчанию dry-run; создание — при confirm=true. Контрагент — покупатель, договор — вида «СПокупателем».",
      inputSchema: z.object(shipmentDocInput),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      sumIncludesVat,
      incomeAccount,
      expenseAccount,
      orgBankAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_shipment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "shipment", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        return createGoodsDoc(
          conn,
          set,
          {
            orgKey: org.key,
            counterpartyRef,
            contractRef,
            warehouseKey,
            date,
            sumIncludesVat,
            lines,
            settlement,
            bankAccountKey: bank,
            lineAccountsFor,
          },
          confirm,
        );
      }),
  );

  server.registerTool(
    "write.document.update_document_lines",
    {
      title: "Изменить строки документа",
      description:
        "Заменяет табличную часть «Товары» существующего документа новым набором строк и пересчитывает сумму. " +
        "Поддержаны: счёт покупателю, поступление, реализация (и товарами, и услугами), акт об оказании услуг. " +
        "Строки пишутся в ту табличную часть, которая у документа заполнена («Товары» или «Услуги»). " +
        "Документ должен быть НЕПРОВЕДЁННЫМ " +
        "(если проведён — сначала отмените проведение через post_document). По умолчанию dry-run; применение — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа"),
        lines: z
          .array(saleLine)
          .min(1)
          .describe(
            "Новый полный набор строк (заменяет прежние — реквизиты старых строк не сохраняются, " +
              "задавайте content и счета заново)",
          ),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, lines, confirm }) =>
      guard("write.document.update_document_lines", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) {
          return fail(
            "Документ проведён. Сначала отмените проведение (post_document с post=false), затем меняйте строки.",
          );
        }
        const built = await buildSectionRows(conn, entitySet, info.orgKey, lines, info.section);
        if (!built) {
          return fail(
            "Поддерживаются: счёт покупателю, поступление/реализация товаров и услуг, акт об оказании " +
              "услуг, возвраты, перемещение, оприходование, списание товаров.",
          );
        }
        return patchOrPreview(
          conn,
          entitySet,
          ref,
          // Вторая ТЧ остаётся как есть, но в сумму документа входит.
          { [info.section]: built.rows, СуммаДокумента: roundMoney(built.total + info.otherTotal) },
          confirm,
        );
      }),
  );

  // === Фаза 1: товарные складские документы ===

  const whLines = z
    .array(
      z.object({
        nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
        quantity: z.number().positive().describe("Количество"),
        price: z.number().nonnegative().default(0).describe("Цена за единицу (для оценки)"),
      }),
    )
    .min(1)
    .describe("Позиции документа");

  const toGoodsLines = (
    lines: Array<{ nomenclatureRef: string; quantity: number; price: number }>,
  ): GoodsLine[] => lines.map((l) => ({ ...l, vatRate: "БезНДС" }));

  server.registerTool(
    "write.warehouse.create_return_from_customer",
    {
      title: "Создать возврат товаров от покупателя",
      description:
        "Создаёт документ «Возврат товаров от покупателя» с табличной частью «Товары». НЕПРОВЕДЁННЫЙ; " +
        "провести — post_document (проводки сторнируют реализацию: Дт 90.02 Кт 41, Дт 62 Кт 90.01 со знаком минус; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "dry-run/confirm. Контрагент — покупатель.",
      inputSchema: z.object(salesDocInput),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      warehouse,
      date,
      incomeAccount,
      expenseAccount,
      lines,
      confirm,
    }) =>
      guard("write.warehouse.create_return_from_customer", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.returnFromCustomer,
          "Документ «Возврат товаров от покупателя»",
        );
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "shipment", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_return_to_supplier",
    {
      title: "Создать возврат товаров поставщику",
      description:
        "Создаёт документ «Возврат товаров поставщику» с табличной частью «Товары». НЕПРОВЕДЁННЫЙ; " +
        "провести — post_document (проводки сторнируют поступление: Дт 60 Кт 41/19; счёт расчётов — по настройке договора, источник виден в предпросмотре). dry-run/confirm. " +
        "Контрагент — поставщик.",
      inputSchema: z.object(purchaseDocInput),
      outputSchema: createResultSchema,
    },
    ({ database, organization, counterpartyRef, contractRef, warehouse, date, lines, confirm }) =>
      guard("write.warehouse.create_return_to_supplier", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.returnToSupplier,
          "Документ «Возврат товаров поставщику»",
        );
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "purchase");
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_transfer",
    {
      title: "Создать перемещение товаров",
      description:
        "Создаёт документ «Перемещение товаров» между складами. НЕПРОВЕДЁННЫЙ; провести — post_document " +
        "(проводки Дт 41(склад-получатель) Кт 41(склад-отправитель)). dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        fromWarehouse: z.string().describe("Название склада-отправителя"),
        toWarehouse: z.string().describe("Название склада-получателя"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: whLines,
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, fromWarehouse, toWarehouse, date, lines, confirm }) =>
      guard("write.warehouse.create_transfer", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.transfer, "Документ «Перемещение товаров»");
        const org = await resolveOrg(conn, organization);
        const [fromKey, toKey] = await Promise.all([
          resolveWarehouse(conn, fromWarehouse),
          resolveWarehouse(conn, toWarehouse),
        ]);
        const lineAcc = await goodsOnlyAccounts(conn, org.key, toGoodsLines(lines));
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.quantity,
            Цена: l.price,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          СкладОтправитель_Key: fromKey,
          СкладПолучатель_Key: toKey,
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_surplus",
    {
      title: "Создать оприходование товаров",
      description:
        "Создаёт документ «Оприходование товаров» (излишки). НЕПРОВЕДЁННЫЙ; провести — post_document " +
        "(проводки Дт 41 Кт 91.01). Можно указать документ-основание — инвентаризацию. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        warehouse: z.string().optional().describe("Название склада (если несколько)"),
        inventoryRef: z.string().optional().describe("Ref_Key инвентаризации-основания (необязательно)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: whLines,
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, warehouse, inventoryRef, date, lines, confirm }) =>
      guard("write.warehouse.create_surplus", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.surplus, "Документ «Оприходование товаров»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const lineAcc = await goodsOnlyAccounts(conn, org.key, toGoodsLines(lines));
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: Math.round(l.quantity * l.price * 100) / 100,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          ИнвентаризацияТоваровНаСкладе_Key: inventoryRef,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_writeoff",
    {
      title: "Создать списание товаров",
      description:
        "Создаёт документ «Списание товаров» (недостача/порча). НЕПРОВЕДЁННЫЙ; провести — post_document " +
        "(проводки Дт 94 Кт 41). Можно указать документ-основание — инвентаризацию. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        warehouse: z.string().optional().describe("Название склада (если несколько)"),
        inventoryRef: z.string().optional().describe("Ref_Key инвентаризации-основания (необязательно)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: whLines,
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, warehouse, inventoryRef, date, lines, confirm }) =>
      guard("write.warehouse.create_writeoff", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.writeoff, "Документ «Списание товаров»");
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const lineAcc = await goodsOnlyAccounts(conn, org.key, toGoodsLines(lines));
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: Math.round(l.quantity * l.price * 100) / 100,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          ИнвентаризацияТоваровНаСкладе_Key: inventoryRef,
          СуммаДокумента: rowsTotal(rows),
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.warehouse.create_inventory",
    {
      title: "Создать инвентаризацию товаров",
      description:
        "Создаёт документ «Инвентаризация товаров на складе» (сверка факт/учёт). Документ НЕ проводится " +
        "(движений не делает) — служит основанием для оприходования/списания. dry-run/confirm. " +
        "Для каждой позиции: фактическое и учётное количество.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        warehouse: z.string().optional().describe("Название склада (если несколько)"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        lines: z
          .array(
            z.object({
              nomenclatureRef: z.string().describe("Ref_Key номенклатуры"),
              factQuantity: z.number().nonnegative().describe("Фактическое количество"),
              accountingQuantity: z.number().nonnegative().describe("Количество по учёту"),
              price: z.number().nonnegative().default(0).describe("Цена (для оценки)"),
            }),
          )
          .min(1)
          .describe("Позиции инвентаризации"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, organization, warehouse, date, lines, confirm }) =>
      guard("write.warehouse.create_inventory", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.inventory,
          "Документ «Инвентаризация товаров на складе»",
        );
        const org = await resolveOrg(conn, organization);
        const warehouseKey = await resolveWarehouse(conn, warehouse);
        const lineAcc = await goodsOnlyAccounts(
          conn,
          org.key,
          lines.map((l) => ({
            nomenclatureRef: l.nomenclatureRef,
            quantity: l.factQuantity,
            price: l.price,
            vatRate: "БезНДС",
          })),
        );
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Количество: l.factQuantity,
            КоличествоУчет: l.accountingQuantity,
            Цена: l.price,
            Сумма: Math.round(l.factQuantity * l.price * 100) / 100,
            СуммаУчет: Math.round(l.accountingQuantity * l.price * 100) / 100,
            ...lineAcc(l.nomenclatureRef),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Склад_Key: warehouseKey,
          Товары: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  const lineObject = saleLine;

  server.registerTool(
    "write.document.add_document_line",
    {
      title: "Добавить строку в документ",
      description:
        "Добавляет одну позицию в табличную часть «Товары» существующего НЕПРОВЕДЁННОГО документа " +
        "(счёт/поступление/реализация), сохраняя прежние строки со всеми их реквизитами " +
        "(содержание, счета учёта, номенклатурная группа). dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        line: lineObject.describe("Добавляемая позиция"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, line, confirm }) =>
      guard("write.document.add_document_line", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) return fail("Документ проведён. Сначала отмените проведение, затем меняйте строки.");
        const built = await buildSectionRows(
          conn,
          entitySet,
          info.orgKey,
          [...info.lines, line],
          info.section,
        );
        if (!built)
          return fail(
            "Поддерживаются: счёт, поступление/реализация, акт услуг, возвраты, перемещение, оприходование, списание.",
          );
        return patchLinesOrPreview(
          conn,
          entitySet,
          ref,
          // Вторая ТЧ остаётся как есть, но в сумму документа входит.
          { [info.section]: built.rows, СуммаДокумента: roundMoney(built.total + info.otherTotal) },
          confirm,
          { section: info.section, before: info.lines.length, expected: info.lines.length + 1 },
        );
      }),
  );

  server.registerTool(
    "write.document.copy_document",
    {
      title: "Скопировать документ",
      description:
        "Создаёт новый документ по образцу существующего: читает его целиком и записывает копию " +
        "со ВСЕМИ реквизитами — содержание строк, счета учёта, банковский счёт, ответственный, " +
        "адрес доставки, доп. условия. Это главный способ выставить ежемесячно повторяющийся " +
        "документ: инструменты create_* собирают документ из полей схемы и всё остальное теряют. " +
        "Копия создаётся НЕПРОВЕДЁННОЙ, с новым номером (1С нумерует сама) и датой (по умолчанию сегодня). " +
        "Правки: date — дата; fields — реквизиты шапки техническими именами (как в update_entity); " +
        "lines — правки строк по номеру (напр. количество и содержание). Сумма документа " +
        "пересчитывается по строкам, если не задана явно в fields. Счета расчётов — как в исходном документе; если в fields сменён договор — по настройке нового договора. По умолчанию dry-run; запись — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_СчетНаОплатуПокупателю"),
        ref: z.string().describe("Ref_Key документа-образца (GUID)"),
        date: z.string().optional().describe("Дата копии YYYY-MM-DD (по умолчанию сегодня)"),
        fields: z
          .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
          .optional()
          .describe(
            'Реквизиты шапки, которые надо изменить: { техническоеИмя: значение }, напр. {"Комментарий":"..."}',
          ),
        lines: z
          .array(
            z.object({
              lineNumber: z.number().int().positive().describe("Номер строки (с 1)"),
              fields: z
                .record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
                .describe(
                  'Поля строки: { техническоеИмя: значение }, напр. {"Количество":4.3,"Сумма":10750,"Содержание":"..."}',
                ),
            }),
          )
          .optional()
          .describe("Правки строк табличной части по номеру (остальные строки копируются как есть)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, entitySet, ref, date, fields, lines, confirm }) =>
      guard("write.document.copy_document", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const guid = ref.replace(/[{}']/g, "");
        const source = await conn.client.getEntity(`${entitySet}(guid'${guid}')${buildQuery({})}`);
        const payload = stripForCopy(source);
        payload["Date"] = odataDate(date ? new Date(`${date}T00:00:00`) : new Date());
        payload["Posted"] = false;

        if (lines?.length) {
          const section = pickSection(
            {
              goods: payload["Товары"] as unknown[] | undefined,
              services: payload["Услуги"] as unknown[] | undefined,
            },
            {
              servicesAct: (DOCUMENTS.servicesAct as readonly string[]).includes(entitySet),
              operationKind: payload["ВидОперации"],
            },
          );
          const rows = sectionRows(payload, section);
          if (rows.length === 0) return fail(`В документе нет строк табличной части «${section}».`);
          const patched = applyLineOverrides(rows, lines);
          payload[section] = patched;
          const total = totalFromRows(patched);
          // Вторая ТЧ копируется как есть, но в сумму документа входит.
          if (total !== undefined)
            payload["СуммаДокумента"] = roundMoney(
              total + rowsSum(sectionRows(payload, otherSection(section))),
            );
        }
        // Реквизиты шапки применяем последними — явное указание важнее пересчёта.
        Object.assign(payload, fields ?? {});
        // Копия сохраняет счета расчётов исходного документа (бухгалтер мог поставить их вручную).
        // Настройку договора применяем, только если в fields сменили договор и счета не заданы явно.
        const has = (k: string) => !!fields && Object.hasOwn(fields, k);
        const contractChanged =
          has("ДоговорКонтрагента_Key") &&
          fields?.["ДоговорКонтрагента_Key"] !== source["ДоговорКонтрагента_Key"];
        const keepSettlement = !contractChanged || has(SETTLE_FIELD) || has(ADVANCE_FIELD);
        return createOrPreview(
          conn,
          entitySet,
          payload,
          confirm,
          keepSettlement && Object.hasOwn(payload, SETTLE_FIELD)
            ? [
                "Счета расчётов — как в исходном документе: при копировании настройка договора не применяется.",
              ]
            : undefined,
          { keepSettlement },
        );
      }),
  );

  server.registerTool(
    "write.document.remove_document_line",
    {
      title: "Удалить строку документа",
      description:
        "Удаляет строку (по номеру LineNumber, начиная с 1) из табличной части «Товары» существующего " +
        "НЕПРОВЕДЁННОГО документа. Нельзя удалить последнюю строку. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Имя документа, напр. Document_РеализацияТоваровУслуг"),
        ref: z.string().describe("Ref_Key документа (GUID)"),
        lineNumber: z.number().int().positive().describe("Номер строки (с 1)"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, lineNumber, confirm }) =>
      guard("write.document.remove_document_line", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const info = await getDocInfo(conn, entitySet, ref.replace(/[{}']/g, ""));
        if (info.posted) return fail("Документ проведён. Сначала отмените проведение, затем меняйте строки.");
        if (lineNumber > info.lines.length) return fail(`В документе всего ${info.lines.length} строк(и).`);
        const kept = info.lines.filter((_, i) => i + 1 !== lineNumber);
        if (kept.length === 0)
          return fail("Нельзя удалить последнюю строку — в документе должна остаться хотя бы одна позиция.");
        const built = await buildSectionRows(conn, entitySet, info.orgKey, kept, info.section);
        if (!built)
          return fail(
            "Поддерживаются: счёт, поступление/реализация, акт услуг, возвраты, перемещение, оприходование, списание.",
          );
        return patchLinesOrPreview(
          conn,
          entitySet,
          ref,
          // Вторая ТЧ остаётся как есть, но в сумму документа входит.
          { [info.section]: built.rows, СуммаДокумента: roundMoney(built.total + info.otherTotal) },
          confirm,
          { section: info.section, before: info.lines.length, expected: kept.length },
        );
      }),
  );

  server.registerTool(
    "write.sales.create_act",
    {
      title: "Создать акт (реализация услуг)",
      description:
        "Создаёт «Реализация (акт, накладная)» с табличной частью УСЛУГИ (без склада/остатков). " +
        "Документ НЕПРОВЕДЁННЫЙ; проводки при проведении Дт 62 Кт 90.01 (без 41; счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "ВАЖНО для услуг: задавайте content в строке — это «Содержание», текст, который печатается " +
        "в акте и УПД (номер приложения, номер договора, период оказания). 1С его не заполнит. " +
        "Счета доходов/расходов при необходимости задаются явно (напр. 90.01.2 / 90.02.2). " +
        "dry-run/confirm. Контрагент — покупатель; позиции — услуги-номенклатура.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора"),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        operationKind: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe(
            "Вид операции (ВидОперации). Обычно НЕ задавайте: по умолчанию «Услуги» или значение из настройки базы " +
              "(ODATA_ACCOUNTING_CONFIG, operationKinds.salesAct). Нужен, только если в конфигурации другое значение — не угадывайте.",
          ),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        incomeAccount: incomeAccountField.describe(
          "Счёт доходов для ВСЕХ строк (напр. «90.01.2»); счёт в строке важнее",
        ),
        expenseAccount: expenseAccountField.describe(
          "Счёт расходов для ВСЕХ строк (напр. «90.02.2»); счёт в строке важнее",
        ),
        orgBankAccount: orgBankAccountField,
        lines: z.array(lineObject).min(1).describe("Позиции-услуги"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      date,
      operationKind,
      sumIncludesVat,
      incomeAccount,
      expenseAccount,
      orgBankAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_act", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.sales, "Документ «Реализация товаров и услуг»");
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "service", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        const rows = buildGoodsRows(lines, lineAccountsFor);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: configuredOperationKind(conn.cfg, "salesAct", operationKind),
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          БанковскийСчетОрганизации_Key: bank,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Услуги: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_services_act",
    {
      title: "Создать акт об оказании услуг",
      description:
        "Создаёт «Акт об оказании производственных услуг» (услуги с учётом доходов/расходов по " +
        "номенклатурной группе). Отличается от create_act тем, что требует номенклатурную группу " +
        "(субконто счёта 90.01) — она НЕ угадывается. Проводки Дт 62 Кт 90.01, НДС 90.03 (счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "НЕПРОВЕДЁННЫЙ; провести — post_document. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().optional().describe("Ref_Key договора"),
        nomenclatureGroupRef: z
          .string()
          .describe(
            "Ref_Key номенклатурной группы (ОБЯЗАТЕЛЬНО, субконто счёта доходов 90.01 — не угадывается)",
          ),
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        sumIncludesVat: z.boolean().default(true).describe("Сумма включает НДС"),
        incomeAccount: incomeAccountField.describe(
          "Счёт доходов для ВСЕХ строк (напр. «90.01.2»); счёт в строке важнее",
        ),
        expenseAccount: expenseAccountField.describe(
          "Счёт расходов для ВСЕХ строк (напр. «90.02.2»); счёт в строке важнее",
        ),
        lines: z.array(lineObject).min(1).describe("Позиции-услуги"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      nomenclatureGroupRef,
      date,
      sumIncludesVat,
      incomeAccount,
      expenseAccount,
      lines,
      confirm,
    }) =>
      guard("write.sales.create_services_act", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.servicesAct,
          "Документ «Акт об оказании производственных услуг»",
        );
        const org = await resolveOrg(conn, organization);
        const { settlement, lineAccountsFor } = await goodsAccounts(conn, org.key, lines, "service", {
          income: incomeAccount,
          expense: expenseAccount,
        });
        const grpSet = resolveEntity(CATALOGS.nomenclatureGroups, await conn.available());
        const grpSubconto = grpSet
          ? { Субконто: nomenclatureGroupRef, Субконто_Type: `StandardODATA.${grpSet}` }
          : {};
        const rows = lines.map((l, i) =>
          clean({
            LineNumber: i + 1,
            Номенклатура_Key: l.nomenclatureRef,
            Содержание: l.content,
            Количество: l.quantity,
            Цена: l.price,
            Сумма: lineSum(l),
            СтавкаНДС: l.vatRate,
            НоменклатурнаяГруппа_Key: nomenclatureGroupRef,
            ...grpSubconto,
            ...lineAccountsFor(l),
          }),
        );
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          Организация_Key: org.key,
          Контрагент_Key: counterpartyRef,
          ДоговорКонтрагента_Key: contractRef,
          НоменклатурнаяГруппа_Key: nomenclatureGroupRef,
          СчетУчетаРасчетовСКонтрагентом_Key: settlement,
          СуммаВключаетНДС: sumIncludesVat,
          СуммаДокумента: rowsTotal(rows),
          Услуги: rows,
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.money.create_payment",
    {
      title: "Создать оплату от покупателя",
      description:
        "Создаёт «Поступление на расчётный счёт» (оплата от покупателя) на сумму, по контрагенту и договору. " +
        "Документ НЕПРОВЕДЁННЫЙ; при проведении проводка Дт 51 Кт 62 (счёт расчётов — по настройке договора, источник виден в предпросмотре). " +
        "Сумма взаиморасчётов для рублёвого договора заполняется сама; для договора в валюте или у.е. " +
        "курс укажите в 1С, иначе 1С проведёт оплату без проводок. " +
        "Банковский счёт организации берётся по названию или первый. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        counterpartyRef: z.string().describe("Ref_Key покупателя"),
        contractRef: z.string().describe("Ref_Key договора"),
        operationKind: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe(
            "Вид операции (ВидОперации). Обычно НЕ задавайте: по умолчанию «ОплатаПокупателя» или значение из настройки " +
              "базы (ODATA_ACCOUNTING_CONFIG, operationKinds.customerPayment). Нужен, только если в конфигурации другое значение — не угадывайте.",
          ),
        amount: z.number().positive().describe("Сумма оплаты"),
        orgBankAccount: orgBankAccountField,
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      counterpartyRef,
      contractRef,
      operationKind,
      amount,
      orgBankAccount,
      date,
      confirm,
    }) =>
      guard("write.money.create_payment", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.bankIn, "Документ «Поступление на расчётный счёт»");
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const codes = conn.cfg?.accounting?.strictDefaults
          ? new Map<AccountDefault, string | undefined>()
          : await resolveDefaultAccounts(conn, ["receivable", "receivableAdvance"]);
        const settle = codes.get("receivable");
        const advance = codes.get("receivableAdvance");
        const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: configuredOperationKind(conn.cfg, "customerPayment", operationKind),
          Организация_Key: org.key,
          СчетОрганизации_Key: bank,
          Контрагент: counterpartyRef,
          ...(cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {}),
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          РасшифровкаПлатежа: [
            clean({
              LineNumber: 1,
              ДоговорКонтрагента_Key: contractRef,
              СпособПогашенияЗадолженности: "Автоматически", // иначе платёж не распределяется и нет проводок
              СуммаПлатежа: amount,
              СтавкаНДС: "БезНДС",
              СчетУчетаРасчетовСКонтрагентом_Key: settle,
              СчетУчетаРасчетовПоАвансам_Key: advance,
            }),
          ],
        });
        return createOrPreview(conn, set, payload, confirm, undefined, { settlementSide: "customer" });
      }),
  );

  // === Фаза 2: денежные документы (списание с р/с, ПКО, РКО) ===

  /**
   * Строка «Расшифровка платежа» для взаиморасчётов с контрагентом. Заполняется
   * только для операций оплаты поставщику/от покупателя при заданном договоре —
   * иначе платёж без взаиморасчётов (налог/ЗП/взнос), ТЧ не нужна.
   */
  async function buildSettlementRow(
    conn: Connection,
    operationKind: string,
    contractRef: string | undefined,
    amount: number,
  ): Promise<{ settle?: string; row: Record<string, unknown> } | undefined> {
    const supplier = /Поставщик/i.test(operationKind);
    const customer = /Покупател/i.test(operationKind);
    if (!contractRef || (!supplier && !customer)) return undefined;
    const codes = conn.cfg?.accounting?.strictDefaults
      ? new Map<AccountDefault, string | undefined>()
      : await resolveDefaultAccounts(
          conn,
          supplier ? ["payable", "payableAdvance"] : ["receivable", "receivableAdvance"],
        );
    const settle = codes.get(supplier ? "payable" : "receivable");
    const advance = codes.get(supplier ? "payableAdvance" : "receivableAdvance");
    return {
      settle,
      row: clean({
        LineNumber: 1,
        ДоговорКонтрагента_Key: contractRef,
        СпособПогашенияЗадолженности: "Автоматически",
        СуммаПлатежа: amount,
        СтавкаНДС: "БезНДС",
        СчетУчетаРасчетовСКонтрагентом_Key: settle,
        СчетУчетаРасчетовПоАвансам_Key: advance,
      }),
    };
  }

  const cpTypeOf = async (conn: Connection): Promise<Record<string, string>> => {
    const cpSet = resolveEntity(CATALOGS.counterparties, await conn.available());
    return cpSet ? { Контрагент_Type: `StandardODATA.${cpSet}` } : {};
  };

  server.registerTool(
    "write.money.create_bank_writeoff",
    {
      title: "Создать списание с расчётного счёта",
      description:
        "Создаёт документ «Списание с расчётного счёта» (исходящий платёж с банка). НЕПРОВЕДЁННЫЙ; " +
        "провести — post_document. Вид операции ОБЯЗАТЕЛЕН и не угадывается (напр. «ОплатаПоставщику», " +
        "«ПеречислениеНалога», «ПеречислениеЗП», «ПереводНаДругойСчет», «ПрочееСписание»). " +
        "Для оплаты поставщику с договором заполняется расшифровка: счёт расчётов — по настройке договора, источник виден в предпросмотре (обычно 60), сумма взаиморасчётов — для рублёвого договора. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        organization: organizationField,
        operationKind: z
          .string()
          .describe(
            "Вид операции (ОБЯЗАТЕЛЬНО, не угадывать): ОплатаПоставщику / ПеречислениеНалога / ПеречислениеЗП / ПереводНаДругойСчет / ПрочееСписание …",
          ),
        amount: z.number().positive().describe("Сумма списания"),
        counterpartyRef: z.string().optional().describe("Ref_Key контрагента (для расчётов)"),
        contractRef: z.string().optional().describe("Ref_Key договора (для расшифровки взаиморасчётов)"),
        purposeText: z.string().optional().describe("Назначение платежа (текст)"),
        cashflowItemRef: z.string().optional().describe("Ref_Key статьи ДДС"),
        orgBankAccount: orgBankAccountField,
        date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
        comment: z.string().optional().describe("Комментарий"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({
      database,
      organization,
      operationKind,
      amount,
      counterpartyRef,
      contractRef,
      purposeText,
      cashflowItemRef,
      orgBankAccount,
      date,
      comment,
      confirm,
    }) =>
      guard("write.money.create_bank_writeoff", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.bankOut, "Документ «Списание с расчётного счёта»");
        const org = await resolveOrg(conn, organization);
        const bank = await resolveOrgBankAccount(conn, org.key, orgBankAccount);
        const bd = await buildSettlementRow(conn, operationKind, contractRef, amount);
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидОперации: operationKind,
          Организация_Key: org.key,
          СчетОрганизации_Key: bank,
          ...(counterpartyRef ? { Контрагент: counterpartyRef, ...(await cpTypeOf(conn)) } : {}),
          ДоговорКонтрагента_Key: contractRef,
          СуммаДокумента: amount,
          НазначениеПлатежа: purposeText,
          СтатьяДвиженияДенежныхСредств_Key: cashflowItemRef,
          Комментарий: comment,
          ...(bd ? { РасшифровкаПлатежа: [bd.row] } : {}),
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  const cashDocInput = {
    database: databaseField,
    organization: organizationField,
    operationKind: z.string().describe("Вид операции (ОБЯЗАТЕЛЬНО, не угадывать)"),
    amount: z.number().positive().describe("Сумма"),
    cashRef: z.string().optional().describe("Ref_Key кассы организации (СчетКасса)"),
    counterpartyRef: z.string().optional().describe("Ref_Key контрагента (для расчётов)"),
    contractRef: z.string().optional().describe("Ref_Key договора"),
    basis: z.string().optional().describe("Основание (текст)"),
    cashflowItemRef: z.string().optional().describe("Ref_Key статьи ДДС"),
    date: z.string().optional().describe("Дата YYYY-MM-DD (по умолчанию сегодня)"),
    comment: z.string().optional().describe("Комментарий"),
    confirm: confirmField,
  };

  const buildCashPayload = async (
    conn: Connection,
    org: { key: string },
    a: {
      operationKind: string;
      amount: number;
      cashRef?: string;
      counterpartyRef?: string;
      contractRef?: string;
      basis?: string;
      cashflowItemRef?: string;
      date?: string;
      comment?: string;
    },
  ): Promise<Record<string, unknown>> => {
    const bd = await buildSettlementRow(conn, a.operationKind, a.contractRef, a.amount);
    return clean({
      Date: odataDate(a.date ? new Date(`${a.date}T00:00:00`) : new Date()),
      Posted: false,
      ВидОперации: a.operationKind,
      Организация_Key: org.key,
      СчетКасса_Key: a.cashRef,
      ...(a.counterpartyRef ? { Контрагент: a.counterpartyRef, ...(await cpTypeOf(conn)) } : {}),
      ДоговорКонтрагента_Key: a.contractRef,
      СуммаДокумента: a.amount,
      Основание: a.basis,
      СтатьяДвиженияДенежныхСредств_Key: a.cashflowItemRef,
      Комментарий: a.comment,
      ...(bd ? { РасшифровкаПлатежа: [bd.row] } : {}),
    });
  };

  server.registerTool(
    "write.money.create_cash_receipt",
    {
      title: "Создать приходный кассовый ордер (ПКО)",
      description:
        "Создаёт «Приходный кассовый ордер» (приём наличных в кассу). НЕПРОВЕДЁННЫЙ; провести — post_document. " +
        "Вид операции ОБЯЗАТЕЛЕН (напр. «ОплатаПокупателя», «ПрочийПриход», «ПолучениеНаличныхВБанке», " +
        "«ВозвратОтПодотчетногоЛица»). " +
        "Для оплаты по договору заполняется расшифровка: счёт расчётов — по настройке договора, источник виден в предпросмотре, сумма взаиморасчётов — для рублёвого договора. " +
        "dry-run/confirm.",
      inputSchema: z.object(cashDocInput),
      outputSchema: createResultSchema,
    },
    (a) =>
      guard("write.money.create_cash_receipt", async () => {
        const conn = ctx.db(a.database);
        const set = await requireEntity(conn, DOCUMENTS.cashIn, "Документ «Приходный кассовый ордер»");
        const org = await resolveOrg(conn, a.organization);
        const payload = await buildCashPayload(conn, org, a);
        return createOrPreview(conn, set, payload, a.confirm);
      }),
  );

  server.registerTool(
    "write.money.create_cash_payment",
    {
      title: "Создать расходный кассовый ордер (РКО)",
      description:
        "Создаёт «Расходный кассовый ордер» (выдача наличных из кассы). НЕПРОВЕДЁННЫЙ; провести — post_document. " +
        "Вид операции ОБЯЗАТЕЛЕН (напр. «ОплатаПоставщику», «ВыдачаПодотчетномуЛицу», «ВзносНаличнымиВБанк», " +
        "«ВыплатаЗаработнойПлатыПоВедомостям»). " +
        "Для оплаты по договору заполняется расшифровка: счёт расчётов — по настройке договора, источник виден в предпросмотре, сумма взаиморасчётов — для рублёвого договора. " +
        "dry-run/confirm.",
      inputSchema: z.object(cashDocInput),
      outputSchema: createResultSchema,
    },
    (a) =>
      guard("write.money.create_cash_payment", async () => {
        const conn = ctx.db(a.database);
        const set = await requireEntity(conn, DOCUMENTS.cashOut, "Документ «Расходный кассовый ордер»");
        const org = await resolveOrg(conn, a.organization);
        const payload = await buildCashPayload(conn, org, a);
        return createOrPreview(conn, set, payload, a.confirm);
      }),
  );

  server.registerTool(
    "write.sales.create_issued_invoice",
    {
      title: "Создать счёт-фактуру выданный",
      description:
        "Создаёт «Счёт-фактура выданный» НА ОСНОВАНИИ реализации (или иного документа отгрузки). " +
        "Реквизиты (организация, контрагент, договор, суммы, НДС) наследуются от документа-основания. " +
        "НЕПРОВЕДЁННЫЙ; провести — post_document. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        baseDocumentRef: z.string().describe("Ref_Key документа-основания (реализация товаров и услуг)"),
        baseDocumentEntity: z
          .string()
          .optional()
          .describe(
            "Имя набора документа-основания, если это не «Реализация…» (напр. Document_ОтчетКомитентуОПродажах)",
          ),
        operationCode: z.string().optional().describe("Код вида операции (по умолчанию 01)"),
        date: z.string().optional().describe("Дата и дата выставления YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, baseDocumentRef, baseDocumentEntity, operationCode, date, confirm }) =>
      guard("write.sales.create_issued_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, DOCUMENTS.issuedInvoice, "Документ «Счёт-фактура выданный»");
        const baseSet =
          baseDocumentEntity ??
          (await requireEntity(conn, DOCUMENTS.sales, "Документ-основание «Реализация товаров и услуг»"));
        const b = await invoiceBasis(conn, baseSet, baseDocumentRef);
        const d = odataDate(date ? new Date(`${date}T00:00:00`) : new Date());
        const baseGuid = baseDocumentRef.replace(/[{}']/g, "");
        const baseRow = clean({
          LineNumber: 1,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
        });
        const payload = clean({
          Date: d,
          Posted: false,
          ВидСчетаФактуры: "НаРеализацию",
          Организация_Key: b.org,
          Контрагент_Key: b.counterparty,
          ДоговорКонтрагента_Key: b.contract,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
          ДатаВыставления: d,
          КодВидаОперации: operationCode ?? "01",
          СуммаДокумента: b.total,
          СуммаНДСДокумента: b.vat,
          ДокументыОснования: [baseRow],
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.purchase.create_received_invoice",
    {
      title: "Создать счёт-фактуру полученный",
      description:
        "Создаёт «Счёт-фактура полученный» НА ОСНОВАНИИ поступления товаров и услуг. " +
        "Реквизиты наследуются от основания; дату входящего счёта-фактуры продавца указывают вручную " +
        "(в поступлении её нет). НЕПРОВЕДЁННЫЙ; провести — post_document. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        baseDocumentRef: z.string().describe("Ref_Key документа-основания (поступление товаров и услуг)"),
        baseDocumentEntity: z
          .string()
          .optional()
          .describe("Имя набора документа-основания, если это не «Поступление…»"),
        incomingDate: z.string().optional().describe("Дата входящего счёта-фактуры продавца YYYY-MM-DD"),
        operationCode: z.string().optional().describe("Код вида операции (по умолчанию 01)"),
        date: z.string().optional().describe("Дата регистрации YYYY-MM-DD (по умолчанию сегодня)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, baseDocumentRef, baseDocumentEntity, incomingDate, operationCode, date, confirm }) =>
      guard("write.purchase.create_received_invoice", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(
          conn,
          DOCUMENTS.receivedInvoice,
          "Документ «Счёт-фактура полученный»",
        );
        const baseSet =
          baseDocumentEntity ??
          (await requireEntity(
            conn,
            DOCUMENTS.purchases,
            "Документ-основание «Поступление товаров и услуг»",
          ));
        const b = await invoiceBasis(conn, baseSet, baseDocumentRef);
        const baseGuid = baseDocumentRef.replace(/[{}']/g, "");
        const baseRow = clean({
          LineNumber: 1,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
        });
        const payload = clean({
          Date: odataDate(date ? new Date(`${date}T00:00:00`) : new Date()),
          Posted: false,
          ВидСчетаФактуры: "НаПоступление",
          Организация_Key: b.org,
          Контрагент_Key: b.counterparty,
          ДоговорКонтрагента_Key: b.contract,
          ДатаВходящегоДокумента: incomingDate ? odataDate(new Date(`${incomingDate}T00:00:00`)) : undefined,
          ДокументОснование: baseGuid,
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
          НДСПредъявленКВычету: true,
          КодВидаОперации: operationCode ?? "01",
          СуммаДокумента: b.total,
          СуммаНДСДокумента: b.vat,
          ДокументыОснования: [baseRow],
        });
        return createOrPreview(conn, set, payload, confirm);
      }),
  );

  server.registerTool(
    "write.counterparty.create_bank_account",
    {
      title: "Создать банковский счёт контрагента",
      description:
        "Заводит расчётный счёт контрагента (подчинённый справочник): банк по БИК + номер счёта. " +
        "Можно сделать счёт основным. По умолчанию dry-run; создание — при confirm=true. " +
        "Ref контрагента — из find_counterparty.",
      inputSchema: z.object({
        database: databaseField,
        ownerRef: z.string().describe("Ref_Key контрагента-владельца счёта"),
        accountNumber: z.string().min(1).describe("Номер расчётного счёта"),
        bik: z.string().min(1).describe("БИК банка (ищется в справочнике «Банки»)"),
        currency: z.string().optional().describe("Валюта (код/название; по умолчанию рубль)"),
        label: z.string().optional().describe("Наименование счёта (по умолчанию — номер счёта)"),
        makeMain: z.boolean().default(false).describe("Сделать основным банковским счётом контрагента"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, ownerRef, accountNumber, bik, currency, label, makeMain, confirm }) =>
      guard("write.counterparty.create_bank_account", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, CATALOGS.bankAccounts, "Справочник «Банковские счета»");
        const bank = await resolveBankByBik(conn, bik);
        const cur = currency
          ? await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", currency)
          : await resolveCatalogItem(conn, CATALOGS.currencies, "Справочник «Валюты»", "643");
        const payload = clean({
          Description: label ?? accountNumber,
          Owner: ownerRef,
          Owner_Type: COUNTERPARTY_TYPE,
          НомерСчета: accountNumber,
          Банк_Key: bank.ref,
          ВалютаДенежныхСредств_Key: cur.ref,
        });
        return createSubordinate(
          conn,
          set,
          payload,
          confirm,
          makeMain
            ? {
                ownerSet: await resolveSet(conn, CATALOGS.counterparties, "Контрагенты"),
                ownerRef,
                field: "ОсновнойБанковскийСчет_Key",
              }
            : undefined,
        );
      }),
  );

  server.registerTool(
    "write.counterparty.create_contact_person",
    {
      title: "Создать контактное лицо (директор и т.п.)",
      description:
        "Заводит контактное лицо контрагента (директор, бухгалтер, менеджер): ФИО + должность. " +
        "Можно сделать основным контактным лицом. По умолчанию dry-run; создание — при confirm=true.",
      inputSchema: z.object({
        database: databaseField,
        ownerRef: z.string().describe("Ref_Key контрагента"),
        name: z.string().min(1).describe("ФИО (наименование контактного лица)"),
        lastName: z.string().optional().describe("Фамилия"),
        firstName: z.string().optional().describe("Имя"),
        position: z.string().optional().describe("Должность (напр. «Директор»)"),
        makeMain: z.boolean().default(false).describe("Сделать основным контактным лицом контрагента"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, ownerRef, name, lastName, firstName, position, makeMain, confirm }) =>
      guard("write.counterparty.create_contact_person", async () => {
        const conn = ctx.db(database);
        const set = await requireEntity(conn, ["Catalog_КонтактныеЛица"], "Справочник «Контактные лица»");
        const payload = clean({
          Description: name,
          Фамилия: lastName,
          Имя: firstName,
          Должность: position,
          ОбъектВладелец: ownerRef,
          ОбъектВладелец_Type: COUNTERPARTY_TYPE,
        });
        return createSubordinate(
          conn,
          set,
          payload,
          confirm,
          makeMain
            ? {
                ownerSet: await resolveSet(conn, CATALOGS.counterparties, "Контрагенты"),
                ownerRef,
                field: "ОсновноеКонтактноеЛицо_Key",
              }
            : undefined,
        );
      }),
  );

  server.registerTool(
    "write.entity.create_folder",
    {
      title: "Создать папку (группу) справочника",
      description:
        "Создаёт папку/группу в иерархическом справочнике 1С (по умолчанию — Контрагенты). " +
        "Опционально вкладывает в родительскую папку (parent — имя папки или Ref_Key). " +
        "В 1С «папки» = группы справочника; отдельных тегов в типовой Бухгалтерии нет. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z
          .string()
          .default("Catalog_Контрагенты")
          .describe("Иерархический справочник (по умолчанию Catalog_Контрагенты)"),
        name: z.string().min(1).describe("Название папки"),
        parent: z.string().optional().describe("Родительская папка — имя или Ref_Key (для вложенности)"),
        confirm: confirmField,
      }),
      outputSchema: createResultSchema,
    },
    ({ database, entitySet, name, parent, confirm }) =>
      guard("write.entity.create_folder", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const parentRef = await folderRefOf(conn, entitySet, parent);
        const payload = clean({ Description: name, IsFolder: true, Parent_Key: parentRef });
        return createOrPreview(conn, entitySet, payload, confirm);
      }),
  );

  server.registerTool(
    "write.entity.move_to_folder",
    {
      title: "Переместить объект в папку",
      description:
        "Перемещает элемент справочника (напр. контрагента) в папку/группу — задаёт родителя (Parent_Key). " +
        "folder — имя папки или Ref_Key. По умолчанию справочник Контрагенты. dry-run/confirm.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z
          .string()
          .default("Catalog_Контрагенты")
          .describe("Справочник (по умолчанию Catalog_Контрагенты)"),
        ref: z.string().describe("Ref_Key перемещаемого элемента"),
        folder: z.string().describe("Папка назначения — имя или Ref_Key"),
        confirm: confirmField,
      }),
      outputSchema: patchResultSchema,
    },
    ({ database, entitySet, ref, folder, confirm }) =>
      guard("write.entity.move_to_folder", async () => {
        const conn = ctx.db(database);
        ensurePublished(await conn.available(), entitySet);
        const parentRef = await folderRefOf(conn, entitySet, folder);
        return patchOrPreview(conn, entitySet, ref, { Parent_Key: parentRef }, confirm);
      }),
  );
}
