import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { buildQuery, cmp, odataGuid } from "../odata/query.js";
import { withoutAccounting } from "../odata/configuration.js";
import { InputError } from "../errors.js";
import { createCore, docDate, resolveCatalogItem, resolveOrgRef } from "./write.js";
import { partnerOf } from "./trade-counterparty.js";
import { seriesSavedNote } from "./trade-warehouse.js";

/**
 * Продажи и закупки Управления торговлей: реализация (товары и услуги), приобретение, возвраты, счёт клиенту, заказы
 * клиента и поставщику, счета-фактуры выданный и полученный. Проверено на копии рабочей УТ 11.5 (21.11.2026): через
 * OData 1С не заполняет налогообложение НДС, график оплаты и склад строки — без них реализация не проводится; объект
 * расчётов, виды запасов и ключи аналитики заполняет сама. Серии — статус 2 в строке и таблица «Серии» со складом,
 * как во всех документах базы (без статуса 1С очищает серии при записи).
 */

export const SALES_DOCS = {
  shipment: "Document_РеализацияТоваровУслуг",
  purchase: "Document_ПриобретениеТоваровУслуг",
  returnFromCustomer: "Document_ВозвратТоваровОтКлиента",
  returnToSupplier: "Document_ВозвратТоваровПоставщику",
  invoice: "Document_СчетНаОплатуКлиенту",
  customerOrder: "Document_ЗаказКлиента",
  supplierOrder: "Document_ЗаказПоставщику",
  issuedInvoice: "Document_СчетФактураВыданный",
  receivedInvoice: "Document_СчетФактураПолученный",
} as const;
export type SalesKind = keyof typeof SALES_DOCS;

/** Хозяйственные операции: первая — по умолчанию (решения владельца, 21.11.2026: комиссия — приём и возврат). */
export const SALES_OPS: Partial<Record<SalesKind, readonly string[]>> = {
  shipment: ["РеализацияКлиенту"],
  purchase: ["ЗакупкаУПоставщика", "ПриемНаКомиссию"],
  returnFromCustomer: ["ВозвратТоваровОтКлиента"],
  returnToSupplier: ["ВозвратТоваровПоставщику", "ВозвратТоваровКомитенту"],
  customerOrder: ["РеализацияКлиенту"],
  supplierOrder: ["ЗакупкаУПоставщика", "ПриемНаКомиссию"],
};
const SALES_EXCLUDED: Record<string, string> = {
  ВозвратОтРозничногоПокупателя: "возврат от розничного покупателя оформляет касса (розница)",
  ПоставкаПодПринципала: "поставка под принципала (агентская схема) коннектором не поддержана",
};

const HEAD: Record<SalesKind, readonly string[]> = {
  shipment: ["ХозяйственнаяОперация", "Соглашение_Key", "НалогообложениеНДС", "Статус"],
  purchase: ["ХозяйственнаяОперация", "НалогообложениеНДС", "Склад_Key"],
  returnFromCustomer: ["ХозяйственнаяОперация", "Соглашение_Key", "НалогообложениеНДС"],
  returnToSupplier: ["ХозяйственнаяОперация", "ДокументПоступления_Key", "НалогообложениеНДС"],
  invoice: ["Партнер_Key", "Договор_Key", "ДокументОснование"],
  customerOrder: ["ХозяйственнаяОперация", "Соглашение_Key", "Статус"],
  supplierOrder: ["ХозяйственнаяОперация", "Статус", "Склад_Key"],
  issuedInvoice: ["ДокументОснование", "ДатаВыставления", "КодВидаОперации"],
  receivedInvoice: ["ДатаСоставления", "КодВидаОперации", "Контрагент"],
};

/** Документ продаж или закупок УТ: база без регистра бухгалтерии, у документа поля УТ. */
export async function tradeSalesSet(conn: Connection, kind: SalesKind): Promise<string | undefined> {
  const set = SALES_DOCS[kind];
  if (!(await conn.available()).has(set) || !(await withoutAccounting(conn))) return undefined;
  const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name) ?? []);
  return HEAD[kind].every((f) => props.has(f)) ? set : undefined;
}

// ─── Строки ───

export interface SalesLine {
  nomenclatureRef: string;
  quantity: number;
  price: number;
  /** Ставка как в Бухгалтерии: БезНДС, НДС0, НДС5, НДС7, НДС10, НДС20, НДС22. */
  vatRate: string;
  content?: string | undefined;
  seriesRef?: string | undefined;
}
const RATE: Record<string, { percent: number; name: string }> = {
  БезНДС: { percent: 0, name: "Без НДС" },
  НДС0: { percent: 0, name: "0%" },
  НДС5: { percent: 5, name: "5%" },
  НДС7: { percent: 7, name: "7%" },
  НДС10: { percent: 10, name: "10%" },
  НДС20: { percent: 20, name: "20%" },
  НДС22: { percent: 22, name: "22%" },
};
const round2 = (n: number) => Math.round(n * 100) / 100;
const guidOf = (s: string) => s.replace(/[{}']/g, "").trim();
const refuse = (text: string): never => {
  throw new InputError(`${text} Ничего не записано.`);
};

async function rateRefs(conn: Connection, lines: readonly SalesLine[]): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const code of new Set(lines.map((l) => l.vatRate))) {
    const r = RATE[code];
    if (!r) refuse(`Ставка НДС «${code}» не поддержана; допустимы: ${Object.keys(RATE).join(", ")}.`);
    out.set(code, (await resolveCatalogItem(conn, ["Catalog_СтавкиНДС"], "Ставка НДС", r!.name)).ref);
  }
  return out;
}

/**
 * Строки товаров или услуг: количество упаковок равно количеству (без упаковок, как в документах базы), суммы по
 * признаку «цена включает НДС», склад в строке у товаров (без него реализация не проводится), серии — в таблице.
 */
function rows(
  lines: readonly SalesLine[],
  rates: Map<string, string>,
  opts: { priceIncludesVat: boolean; warehouse?: string | undefined; services?: boolean },
): { Товары: Array<Record<string, unknown>>; Серии?: Array<Record<string, unknown>>; total: number } {
  let total = 0;
  const Товары = lines.map((l, i) => {
    const r = RATE[l.vatRate]!;
    const sum = round2(l.quantity * l.price);
    const vat = round2(
      opts.priceIncludesVat ? (sum * r.percent) / (100 + r.percent) : (sum * r.percent) / 100,
    );
    const withVat = opts.priceIncludesVat ? sum : round2(sum + vat);
    total += withVat;
    return {
      LineNumber: i + 1,
      Номенклатура_Key: guidOf(l.nomenclatureRef),
      Количество: l.quantity,
      ...(opts.services ? {} : { КоличествоУпаковок: l.quantity }),
      Цена: l.price,
      Сумма: sum,
      СтавкаНДС_Key: rates.get(l.vatRate),
      ...(vat ? { СуммаНДС: vat } : {}),
      СуммаСНДС: withVat,
      ...(l.content ? { Содержание: l.content } : {}),
      ...(opts.warehouse && !opts.services ? { Склад_Key: opts.warehouse } : {}),
      ...(l.seriesRef ? { СтатусУказанияСерий: 2 } : {}),
    };
  });
  const series = lines.filter((l) => l.seriesRef);
  return {
    Товары,
    total: round2(total),
    ...(series.length
      ? {
          Серии: series.map((l, i) => ({
            LineNumber: i + 1,
            Номенклатура_Key: guidOf(l.nomenclatureRef),
            Серия_Key: guidOf(l.seriesRef!),
            Количество: l.quantity,
            ...(opts.warehouse ? { Склад_Key: opts.warehouse } : {}),
          })),
        }
      : {}),
  };
}

/**
 * Налогообложение НДС: через OData 1С его у реализации не заполняет, и проведение отказывает. Задано — как есть; нет —
 * как в последнем проведённом документе этого вида у организации (у организации оно одно), с заметкой; нет такого —
 * отказ с просьбой указать.
 */
async function vatTaxation(
  conn: Connection,
  set: string,
  org: string,
  given: string | undefined,
): Promise<{ value: string; note?: string }> {
  if (given) return { value: given };
  // Сначала документ того же вида, затем реализация: налогообложение у организации одно (заказов клиента может не быть).
  let last: ODataEntity | undefined;
  for (const from of [...new Set([set, SALES_DOCS.shipment])]) {
    last = (
      await conn.client.getCollection<ODataEntity>(
        `${from}${buildQuery({
          filter: `${cmp("Организация_Key", "eq", odataGuid(org))} and Posted eq true and DeletionMark eq false`,
          orderby: "Date desc",
          top: 1,
          select: ["Number", "Date", "НалогообложениеНДС"],
        })}`,
      )
    ).value[0];
    if (last?.["НалогообложениеНДС"]) break;
  }
  const value = String(last?.["НалогообложениеНДС"] ?? "");
  if (!value)
    refuse(
      "Налогообложение НДС не задано, а взять его не из чего: у организации нет проведённых документов этого вида. Укажите vatTaxation (например, ПродажаНеОблагаетсяНДС или ПродажаОблагаетсяНДС).",
    );
  return {
    value,
    note: `Налогообложение НДС «${value}» — как в последнем проведённом документе организации № ${String(last!["Number"])} от ${String(last!["Date"]).slice(0, 10)}.`,
  };
}

// ─── Создание ───

export interface SalesInput {
  organization?: string | undefined;
  operationKind?: string | undefined;
  counterpartyRef?: string | undefined;
  contractRef?: string | undefined;
  agreementRef?: string | undefined;
  /** Склад: Ref_Key (разрешает инструмент). */
  warehouse?: string | undefined;
  lines?: SalesLine[] | undefined;
  priceIncludesVat?: boolean | undefined;
  vatTaxation?: string | undefined;
  paymentDueDate?: string | undefined;
  status?: string | undefined;
  /** Основание: реализация для возврата клиента и счёта-фактуры, приобретение для возврата поставщику. */
  basisRef?: string | undefined;
  basisSet?: string | undefined;
  orgBankAccount?: string | undefined;
  /** Реализация услуг (акт): строки — услуги, без склада. */
  services?: boolean | undefined;
  incomingNumber?: string | undefined;
  incomingDate?: string | undefined;
  operationCode?: string | undefined;
  date?: string | undefined;
  comment?: string | undefined;
}

const AGREEMENTS: Record<"client" | "supplier", string> = {
  client: "Catalog_СоглашенияСКлиентами",
  supplier: "Catalog_СоглашенияСПоставщиками",
};

/**
 * Соглашение: задано — как есть; нет — единственное действующее соглашение партнёра с организацией (как договор по
 * умолчанию в Бухгалтерии), с заметкой; несколько или ни одного — отказ со списком.
 */
async function agreementOf(
  conn: Connection,
  side: "client" | "supplier",
  partner: string,
  org: string,
  given: string | undefined,
): Promise<{ ref: string; note?: string }> {
  if (given) return { ref: guidOf(given) };
  const set = AGREEMENTS[side];
  const rows = (
    await conn.client.getCollection<ODataEntity>(
      `${set}${buildQuery({
        filter: `${cmp("Партнер_Key", "eq", odataGuid(partner))} and DeletionMark eq false`,
        select: ["Ref_Key", "Description", "Организация_Key", "Статус"],
      })}`,
    )
  ).value.filter(
    (a) =>
      (!a["Организация_Key"] ||
        /^0{8}-/.test(String(a["Организация_Key"])) ||
        a["Организация_Key"] === org) &&
      a["Статус"] !== "Закрыто",
  );
  if (rows.length === 1)
    return {
      ref: String(rows[0]!["Ref_Key"]),
      note: `Соглашение — единственное у партнёра: «${String(rows[0]!["Description"])}».`,
    };
  return refuse(
    `Укажите agreementRef — соглашение с ${side === "client" ? "клиентом" : "поставщиком"}: ${rows.length ? `у партнёра их несколько (${rows.map((a) => `«${a["Description"]}» ${a["Ref_Key"]}`).join(", ")})` : "у партнёра нет действующих"}.`,
  );
}

const DAY = (d: string) => `${d.slice(0, 10)}T00:00:00`;

/** Создание документа продаж или закупок УТ. */
export async function createTradeSales(
  conn: Connection,
  kind: SalesKind,
  input: SalesInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const set = SALES_DOCS[kind];
  const ops = SALES_OPS[kind];
  const op = input.operationKind ?? ops?.[0];
  if (ops && (!op || !ops.includes(op))) {
    const why = op && Object.hasOwn(SALES_EXCLUDED, op) ? SALES_EXCLUDED[op] : undefined;
    refuse(
      `Хозяйственная операция «${op}» ${why ? `не поддерживается: ${why}.` : "здесь не поддержана."} Поддержаны: ${ops.join(", ")}.`,
    );
  }
  const date = await docDate(conn, input.date);
  const notes: string[] = [];
  const rub = (await resolveCatalogItem(conn, ["Catalog_Валюты"], "Справочник «Валюты»", "643")).ref;

  // Счета-фактуры — на основании: организация и контрагент из документа-основания.
  if (kind === "issuedInvoice" || kind === "receivedInvoice") {
    if (!input.basisRef)
      refuse(
        `Счёт-фактура ${kind === "issuedInvoice" ? "выданный" : "полученный"} создаётся на основании документа: укажите basisRef.`,
      );
    const baseSet = input.basisSet ?? (kind === "issuedInvoice" ? SALES_DOCS.shipment : SALES_DOCS.purchase);
    const base = await conn.client.getEntity<ODataEntity>(
      `${baseSet}(${odataGuid(guidOf(input.basisRef!))})?$format=json`,
    );
    if (base["Posted"] !== true) refuse("Документ-основание не проведён: сначала проведите его.");
    const head: Record<string, unknown> = {
      Date: date,
      Posted: false,
      Организация_Key: base["Организация_Key"],
      Контрагент: base["Контрагент_Key"],
      Контрагент_Type: "StandardODATA.Catalog_Контрагенты",
      Партнер_Key: base["Партнер_Key"],
      Валюта_Key: base["Валюта_Key"] ?? rub,
      КодВидаОперации: input.operationCode ?? "01",
      ДокументыОснования: [
        {
          LineNumber: 1,
          ДокументОснование: guidOf(input.basisRef!),
          ДокументОснование_Type: `StandardODATA.${baseSet}`,
        },
      ],
    };
    if (base["Договор_Key"] && !/^0{8}-/.test(String(base["Договор_Key"]))) {
      head["Договор"] = base["Договор_Key"];
      head["Договор_Type"] = "StandardODATA.Catalog_ДоговорыКонтрагентов";
    }
    if (kind === "issuedInvoice") {
      head["ДокументОснование"] = guidOf(input.basisRef!);
      head["ДокументОснование_Type"] = `StandardODATA.${baseSet}`;
      head["ДатаВыставления"] = date;
      if (input.comment) head["Комментарий"] = input.comment;
    } else {
      if (!input.incomingNumber || !input.incomingDate)
        refuse(
          "Счёт-фактура полученный: укажите incomingNumber и incomingDate — номер и дату счёта-фактуры поставщика.",
        );
      // Номер входящего счёта-фактуры в УТ — номер самого документа (отдельного поля нет).
      head["Number"] = input.incomingNumber;
      head["ДатаСоставления"] = DAY(input.incomingDate!);
    }
    return createCore(conn, set, head, confirm, notes, () => undefined);
  }

  if (!input.counterpartyRef) refuse("Укажите counterpartyRef — контрагента.");
  const cp = guidOf(input.counterpartyRef!);
  const partner = await partnerOf(conn, cp);
  if (!partner) refuse("У контрагента нет партнёра: в Управлении торговлей документы ведутся по партнёру.");
  const org = await resolveOrgRef(conn, input.organization);
  const head: Record<string, unknown> = {
    Date: date,
    Posted: false,
    Организация_Key: org,
    Партнер_Key: partner,
    Контрагент_Key: cp,
    Валюта_Key: rub,
    ...(op ? { ХозяйственнаяОперация: op } : {}),
    ...(input.contractRef ? { Договор_Key: guidOf(input.contractRef) } : {}),
    ...(input.comment ? { Комментарий: input.comment } : {}),
  };

  // Счёт клиенту в УТ — сумма и график оплаты, без строк товаров.
  if (kind === "invoice") {
    const lines = input.lines ?? [];
    const total = round2(lines.reduce((a, l) => a + l.quantity * l.price, 0));
    if (!(total > 0)) refuse("Счёт клиенту: сумма по строкам должна быть больше нуля.");
    Object.assign(head, {
      СуммаДокумента: total,
      ЭтапыГрафикаОплаты: [
        {
          LineNumber: 1,
          ДатаПлатежа: DAY(input.paymentDueDate ?? date),
          ПроцентПлатежа: 100,
          СуммаПлатежа: total,
        },
      ],
    });
    if (input.basisRef) {
      head["ДокументОснование"] = guidOf(input.basisRef);
      head["ДокументОснование_Type"] = `StandardODATA.${input.basisSet ?? SALES_DOCS.shipment}`;
    }
    notes.push(
      "Счёт на оплату клиенту в Управлении торговлей хранит сумму и график оплаты, без строк товаров.",
    );
    return createCore(conn, set, head, confirm, notes, () => undefined);
  }

  const lines = input.lines ?? [];
  if (!lines.length) refuse("Укажите строки документа (lines).");
  const goods = !input.services;
  if (goods && !input.warehouse) refuse("Укажите склад (warehouse).");
  if (input.services && lines.some((l) => l.seriesRef)) refuse("У услуг серий нет.");
  const priceIncludesVat = input.priceIncludesVat ?? true;
  head["ЦенаВключаетНДС"] = priceIncludesVat;
  if (goods) head["Склад_Key"] = input.warehouse;

  const side: "client" | "supplier" =
    kind === "shipment" || kind === "returnFromCustomer" || kind === "customerOrder" ? "client" : "supplier";
  if (side === "client" || input.agreementRef) {
    const a = await agreementOf(conn, side, partner!, org, input.agreementRef);
    head["Соглашение_Key"] = a.ref;
    if (a.note) notes.push(a.note);
  }
  // Налогообложение: у продаж 1С через OData его не заполняет; у закупок заполняет сама.
  if (side === "client") {
    const v = await vatTaxation(conn, set, org, input.vatTaxation);
    head["НалогообложениеНДС"] = v.value;
    if (v.note) notes.push(v.note);
  } else if (input.vatTaxation) head["НалогообложениеНДС"] = input.vatTaxation;

  const rates = await rateRefs(conn, lines);
  if (head["НалогообложениеНДС"] === "ПродажаОблагаетсяНДС" && lines.some((l) => l.vatRate === "БезНДС"))
    notes.push("Организация продаёт с НДС, а у строк ставка «Без НДС» — проверьте vatRate.");
  const built = rows(lines, rates, {
    priceIncludesVat,
    warehouse: input.warehouse,
    services: input.services ?? false,
  });
  head["СуммаДокумента"] = built.total;
  head["Товары"] = built.Товары;
  if (built.Серии) head["Серии"] = built.Серии;

  if (kind === "shipment") {
    head["Статус"] = input.status ?? "Отгружено";
    // График оплаты 1С через OData не заполняет: одна строка, как у всех реализаций базы.
    head["ЭтапыГрафикаОплаты"] = [
      {
        LineNumber: 1,
        ВариантОплаты: "КредитПослеОтгрузки",
        ДатаПлатежа: DAY(input.paymentDueDate ?? date),
        СуммаПлатежа: built.total,
        ПроцентПлатежа: 100,
        СуммаВзаиморасчетов: built.total,
        Сдвиг: 0,
        ВариантОтсчета: "ОтДатыОтгрузки",
      },
    ];
    if (input.services)
      notes.push(
        "Услуги в Управлении торговлей продаются реализацией: акт выполненных работ эта конфигурация через OData не проводит.",
      );
  }
  if (kind === "customerOrder") head["Статус"] = input.status ?? "КОбеспечению";
  if (kind === "supplierOrder") head["Статус"] = input.status ?? "НеСогласован";
  if (kind === "returnToSupplier" || kind === "returnFromCustomer") {
    if (!input.basisRef)
      refuse(
        `Возврат ${kind === "returnToSupplier" ? "поставщику — укажите basisRef: приобретение" : "от клиента — укажите basisRef: реализацию"}, по которому возвращается товар.`,
      );
    const basis = guidOf(input.basisRef!);
    if (kind === "returnToSupplier") {
      head["ДокументПоступления_Key"] = basis;
      head["Товары"] = built.Товары.map((r) => ({ ...r, ДокументПоступления_Key: basis }));
    } else
      head["Товары"] = built.Товары.map((r) => ({
        ...r,
        ДокументРеализации: basis,
        ДокументРеализации_Type: `StandardODATA.${SALES_DOCS.shipment}`,
      }));
  }
  const sent = built.Серии?.length ?? 0;
  return createCore(conn, set, head, confirm, notes, (created) => seriesSavedNote(sent, created));
}

// ─── Проверка проведения ───

type Dir = "Receipt" | "Expense";
/** Что двигает документ: склад (направление) и расчёты с партнёром (регистр, направление). */
const MOVES: Partial<
  Record<SalesKind, { stock: Dir; party?: { register: string; dir: Dir; ops?: readonly string[] } }>
> = {
  shipment: {
    stock: "Expense",
    party: { register: "AccumulationRegister_РасчетыСКлиентами", dir: "Receipt" },
  },
  purchase: {
    stock: "Receipt",
    party: {
      register: "AccumulationRegister_РасчетыСПоставщиками",
      dir: "Expense",
      ops: ["ЗакупкаУПоставщика"],
    },
  },
  returnToSupplier: {
    stock: "Expense",
    party: {
      register: "AccumulationRegister_РасчетыСПоставщиками",
      dir: "Receipt",
      ops: ["ВозвратТоваровПоставщику"],
    },
  },
  returnFromCustomer: {
    stock: "Receipt",
    party: { register: "AccumulationRegister_РасчетыСКлиентами", dir: "Expense" },
  },
};
const STOCK = "AccumulationRegister_ТоварыНаСкладах";
const SERIES_REG = "AccumulationRegister_ДвиженияСерийТоваров";

async function recordsOf(
  conn: Connection,
  register: string,
  set: string,
  guid: string,
): Promise<ODataEntity[]> {
  try {
    const s = await conn.client.getEntity<ODataEntity>(
      `${register}(Recorder='${guid}', Recorder_Type='StandardODATA.${set}')?$format=json`,
    );
    return ((s["RecordSet"] as ODataEntity[] | undefined) ?? []).filter((r) => r["Active"] !== false);
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found") return [];
    throw e;
  }
}
const qty = (n: number) => Math.round(n * 1000);
const cents = (n: number) => Math.round(n * 100);
const low = (v: unknown) => String(v ?? "").toLowerCase();
const empty = (v: unknown) => !v || /^0{8}-/.test(String(v));

/**
 * Проверка проведения документа продаж или закупок УТ: по складу — расход или приход по каждой позиции со складом в
 * строке, по расчётам — сумма документа, по сериям — количество каждой серии таблицы «Серии»; после отмены — пусто.
 * Счета, заказы и счета-фактуры — только состояние документа. undefined — не документ продаж и закупок УТ.
 */
export async function verifySalesMovements(
  conn: Connection,
  entitySet: string,
  ref: string,
  post: boolean,
): Promise<Record<string, unknown> | undefined> {
  const kind = (Object.keys(SALES_DOCS) as SalesKind[]).find((k) => SALES_DOCS[k] === entitySet);
  if (!kind || !(await tradeSalesSet(conn, kind))) return undefined;
  const spec = MOVES[kind];
  if (!spec)
    return {
      movementsNote:
        "Счёт, заказ и счёт-фактура движений по складу и расчётам не делают: проверено состояние документа.",
    };
  const guid = guidOf(ref);
  try {
    const doc = await conn.client.getEntity<ODataEntity>(`${entitySet}(${odataGuid(guid)})?$format=json`);
    const op = String(doc["ХозяйственнаяОперация"] ?? "");
    const party = spec.party && (!spec.party.ops || spec.party.ops.includes(op)) ? spec.party : undefined;
    const regs = [STOCK, ...(party ? [party.register] : []), SERIES_REG];
    const found: Record<string, ODataEntity[]> = {};
    for (const r of regs) found[r] = await recordsOf(conn, r, entitySet, guid);
    const mismatches: string[] = [];
    if (!post) {
      for (const [r, rows] of Object.entries(found))
        if (rows.length) mismatches.push(`после отмены проведения в ${r} осталось записей: ${rows.length}`);
    } else {
      // Склад: позиции со складом в строке — у товаров он всегда в строке (без него 1С не проводит), у услуг его нет.
      const want = new Map<string, number>();
      for (const r of (doc["Товары"] as ODataEntity[] | undefined) ?? []) {
        if (empty(r["Склад_Key"])) continue;
        const k = `${low(r["Склад_Key"])} ${low(r["Номенклатура_Key"])}`;
        want.set(k, (want.get(k) ?? 0) + qty(Number(r["Количество"] ?? 0)));
      }
      const got = new Map<string, number>();
      for (const r of found[STOCK]!) {
        const k = `${low(r["Склад_Key"])} ${low(r["Номенклатура_Key"])}`;
        const sign = r["RecordType"] === spec.stock ? 1 : -1;
        got.set(k, (got.get(k) ?? 0) + sign * qty(Number(r["ВНаличии"] ?? 0)));
      }
      for (const [k, n] of want) {
        const g = got.get(k) ?? 0;
        if (g !== n) mismatches.push(`склад ${k}: в регистре ${g / 1000}, в документе ${n / 1000}`);
        got.delete(k);
      }
      for (const [k, n] of got) if (n) mismatches.push(`склад ${k}: лишнее движение ${n / 1000}`);
      if (party) {
        const sum = found[party.register]!.filter((r) => r["RecordType"] === party.dir).reduce(
          (a, r) => a + cents(Number(r["Сумма"] ?? 0)),
          0,
        );
        const amount = cents(Number(doc["СуммаДокумента"] ?? 0));
        if (sum !== amount) mismatches.push(`${party.register}: ${sum / 100}, в документе ${amount / 100}`);
      }
      for (const s of (doc["Серии"] as ODataEntity[] | undefined) ?? []) {
        const n = found[SERIES_REG]!.filter((r) => low(r["Серия_Key"]) === low(s["Серия_Key"])).reduce(
          (a, r) => a + qty(Number(r["Количество"] ?? 0)),
          0,
        );
        if (n !== qty(Number(s["Количество"] ?? 0)))
          mismatches.push(
            `серия ${low(s["Серия_Key"])}: в регистре серий ${n / 1000}, в документе ${Number(s["Количество"])}`,
          );
      }
    }
    const matched = !mismatches.length;
    return {
      movementsVerified: matched,
      movements: {
        register: regs.join(", "),
        found: Object.values(found).reduce((a, r) => a + r.length, 0),
        ...(matched ? {} : { mismatches: mismatches.slice(0, 20) }),
      },
      ...(matched
        ? {}
        : {
            verified: false,
            note: post
              ? "Документ проведён, но движения по складу, расчётам или сериям не совпали со строками. Проверьте документ в 1С; отменить проведение — write.document.post_document с post=false."
              : "Проведение отменено, но движения документа остались. Проверьте его в 1С.",
          }),
    };
  } catch (e) {
    return {
      movementsVerified: false,
      verified: false,
      note: `Действие выполнено, но движения прочитать не удалось (${e instanceof Error ? e.message : String(e)}). Проверьте документ в 1С, прежде чем повторять.`,
    };
  }
}
