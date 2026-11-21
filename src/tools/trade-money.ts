import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { buildQuery, cmp, odataGuid } from "../odata/query.js";
import { withoutAccounting } from "../odata/configuration.js";
import { InputError } from "../errors.js";
import { createCore, docDate, resolveCatalogItem, resolveOrgRef } from "./write.js";
import { partnerOf } from "./trade-counterparty.js";

/**
 * Денежные документы Управления торговлей: ПКО, РКО, поступление и списание безналичных, заявка на расходование.
 * Проверено на копии рабочей УТ 11.5 (21.11.2026): по каждой поддержанной хозяйственной операции документ с
 * минимальным набором полей 1С принимает, объект расчётов заполняет сама при записи и проводит. Безналичные — с
 * признаком «проведено банком»: без него движение идёт в «к зачислению/к списанию», а не в остаток.
 */

export const MONEY_DOCS = {
  cashIn: "Document_ПриходныйКассовыйОрдер",
  cashOut: "Document_РасходныйКассовыйОрдер",
  bankIn: "Document_ПоступлениеБезналичныхДенежныхСредств",
  bankOut: "Document_СписаниеБезналичныхДенежныхСредств",
  request: "Document_ЗаявкаНаРасходованиеДенежныхСредств",
} as const;
export type MoneyKind = keyof typeof MONEY_DOCS;

type Party = "client" | "supplier" | "refund";
interface OpSpec {
  /** Расчёты с партнёром: оплата клиента, поставщику, возврат клиенту. */
  party?: Party;
  /** Прочее: статья доходов или расходов в строке. */
  item?: "income" | "expense";
  /** Что ещё нужно операции. */
  needs?: ReadonlyArray<"person" | "fromCash" | "toBank" | "fromBank" | "counterparty" | "budget">;
  /** Строка расшифровки: у перемещений денег её нет. */
  noRows?: boolean;
}

/** Поддержанные хозяйственные операции (решение владельца, 21.11.2026). */
export const MONEY_OPS: Record<MoneyKind, Record<string, OpSpec>> = {
  cashIn: {
    ПоступлениеОплатыОтКлиента: { party: "client" },
    ПрочееПоступлениеДенежныхСредств: { item: "income" },
    ПоступлениеДенежныхСредствИзДругойКассы: { needs: ["fromCash"], noRows: true },
  },
  cashOut: {
    ОплатаПоставщику: { party: "supplier" },
    ПрочаяВыдачаДенежныхСредств: { item: "expense" },
    ВыдачаДенежныхСредствПодотчетнику: { needs: ["person"] },
    ИнкассацияДенежныхСредствВБанк: { needs: ["toBank"], noRows: true },
  },
  bankIn: {
    ПоступлениеОплатыОтКлиента: { party: "client" },
    ПрочееПоступлениеДенежныхСредств: { item: "income" },
    ПоступлениеДенежныхСредствСДругогоСчета: { needs: ["fromBank"], noRows: true },
    ПоступлениеДенежныхСредствИзКассыНаРасчетныйСчет: { needs: ["fromCash"], noRows: true },
    ПоступлениеОплатыПоПлатежнойКарте: { needs: ["counterparty"], noRows: true },
  },
  bankOut: {
    ОплатаПоставщику: { party: "supplier" },
    ПрочаяВыдачаДенежныхСредств: { item: "expense" },
    ВозвратОплатыКлиенту: { party: "refund" },
    ПеречислениеВБюджет: { needs: ["counterparty", "budget"] },
  },
  request: {
    ОплатаПоставщику: { party: "supplier" },
    ПрочаяВыдачаДенежныхСредств: { item: "expense" },
    ВозвратОплатыКлиенту: { party: "refund" },
    ПеречислениеВБюджет: { needs: ["counterparty", "budget"] },
  },
};

/** Исключения по замыслу: розница и зарплата делаются своими механизмами 1С. */
const EXCLUDED: Record<string, string> = {
  ПоступлениеДенежныхСредствИзКассыККМ: "поступление из кассы ККМ оформляет розница (касса), а не коннектор",
  ВыдачаДенежныхСредствВКассуККМ: "выдачу в кассу ККМ оформляет розница (касса), а не коннектор",
  ВыплатаЗарплатыПоЗарплатномуПроекту: "зарплата — в зарплатной программе",
  ВыплатаЗарплаты: "зарплата — в зарплатной программе",
  ВозвратНеперечисленнойЗарплатыПоЗарплатномуПроекту: "зарплата — в зарплатной программе",
};

const HEAD: Record<MoneyKind, readonly string[]> = {
  cashIn: ["Касса_Key", "ХозяйственнаяОперация", "ПринятоОт", "КассаОтправитель"],
  cashOut: ["Касса_Key", "ХозяйственнаяОперация", "Выдать", "ПодотчетноеЛицо_Key"],
  bankIn: ["БанковскийСчет_Key", "ХозяйственнаяОперация", "ПроведеноБанком", "БанковскийСчетОтправитель_Key"],
  bankOut: ["БанковскийСчет_Key", "ХозяйственнаяОперация", "ПроведеноБанком", "КодБК"],
  request: ["ХозяйственнаяОперация", "Статус", "ФормаОплатыЗаявки", "ЖелательнаяДатаПлатежа"],
};

/** Денежный документ УТ: база без регистра бухгалтерии, у документа поля УТ (хозяйственная операция, касса). */
export async function tradeMoneySet(conn: Connection, kind: MoneyKind): Promise<string | undefined> {
  const set = MONEY_DOCS[kind];
  if (!(await conn.available()).has(set) || !(await withoutAccounting(conn))) return undefined;
  const meta = await conn.getMetadata();
  const head = new Set(meta.entities.get(set)?.properties.map((p) => p.name) ?? []);
  const rows = new Set(meta.entities.get(`${set}_РасшифровкаПлатежа`)?.properties.map((p) => p.name) ?? []);
  return HEAD[kind].every((f) => head.has(f)) && rows.has("Сумма") ? set : undefined;
}

export interface BudgetInput {
  kbk: string;
  oktmo: string;
  payerStatus: string;
  kind?: string | undefined;
  basis?: string | undefined;
  period?: string | undefined;
  number?: string | undefined;
  date?: string | undefined;
}

export interface MoneyInput {
  organization?: string | undefined;
  operationKind: string;
  amount: number;
  /** Касса: название, код или Ref_Key; одна у организации — сама. */
  cash?: string | undefined;
  orgBankAccount?: string | undefined;
  counterpartyRef?: string | undefined;
  contractRef?: string | undefined;
  cashflowItem?: string | undefined;
  incomeItem?: string | undefined;
  expenseItem?: string | undefined;
  department?: string | undefined;
  person?: string | undefined;
  fromCash?: string | undefined;
  toBankAccount?: string | undefined;
  fromBankAccount?: string | undefined;
  /** Кому выдано / от кого принято (текст ордера); по умолчанию — наименование контрагента или лица. */
  partyText?: string | undefined;
  basis?: string | undefined;
  purpose?: string | undefined;
  budget?: BudgetInput | undefined;
  /** Заявка: наличными из кассы или безналичным со счёта. */
  payForm?: "cash" | "bank" | undefined;
  status?: "НеСогласована" | "Согласована" | "КОплате" | undefined;
  paymentDate?: string | undefined;
  date?: string | undefined;
  comment?: string | undefined;
  /** Несколько строк расшифровки по документам расчётов (1.1). */
  lines?:
    Array<{ contractRef?: string | undefined; basisRef?: string | undefined; amount: number }> | undefined;
}

/**
 * Строки по документам расчётов: какие документы могут быть основанием строки. Объект расчётов строки — запись
 * справочника «Объекты расчётов» с Объект = документ (на копии УТ 11.5: закупка и заказ поставщику — «Накладная» и
 * «Заказ»). Проверено проведением 21.11.2026: оплата поставщику двумя строками по двум закупкам — расчёты по каждой.
 */
const LINE_BASIS: Record<string, readonly string[]> = {
  "bankOut:ОплатаПоставщику": [
    "Document_ПриобретениеТоваровУслуг",
    "Document_ЗаказПоставщику",
    "Document_ПриобретениеУслугПрочихАктивов",
  ],
  "bankIn:ПоступлениеОплатыОтКлиента": ["Document_РеализацияТоваровУслуг", "Document_ЗаказКлиента"],
};

async function settlementObjectRows(
  conn: Connection,
  kind: MoneyKind,
  input: MoneyInput,
  row: Record<string, unknown>,
  partner: unknown,
  org: string,
): Promise<Record<string, unknown>[]> {
  const lines = input.lines!;
  const sets = LINE_BASIS[`${kind}:${input.operationKind}`];
  if (!sets)
    refuse(
      `lines — для оплаты поставщику (списание безналичных) и поступления оплаты от клиента (поступление безналичных); у «${input.operationKind}» строк по документам нет.`,
    );
  if (lines.some((l) => l.contractRef !== undefined || !l.basisRef))
    refuse(
      "Управление торговлей: строки lines — по документам расчётов, в каждой basisRef (договор задаётся contractRef документа).",
    );
  const sum = Math.round(lines.reduce((t, l) => t + l.amount, 0) * 100) / 100;
  if (Math.abs(sum - input.amount) >= 0.005)
    refuse(
      `Сумма строк lines ${sum} не равна сумме документа ${input.amount}: суммы строк должны давать amount.`,
    );
  const refs = lines.map((l) => guidOf(l.basisRef!).toLowerCase());
  if (new Set(refs).size !== refs.length)
    refuse("В lines документ повторяется: объедините суммы в одну строку.");
  const available = await conn.available();
  if (!available.has("Catalog_ОбъектыРасчетов"))
    refuse(
      "Строки по документам расчётов: справочник «Объекты расчётов» не опубликован в OData — добавьте его в состав стандартного интерфейса OData.",
    );
  const rows: Record<string, unknown>[] = [];
  for (const [i, ref] of refs.entries()) {
    let obj: ODataEntity | undefined;
    for (const set of sets!) {
      if (!available.has(set)) continue;
      obj = (
        await conn.client.getCollection<ODataEntity>(
          `Catalog_ОбъектыРасчетов${buildQuery({
            filter: `Объект eq cast(guid'${ref}', '${set}')`,
            select: ["Ref_Key", "Партнер_Key", "Организация_Key", "DeletionMark"],
          })}`,
        )
      ).value[0];
      if (obj) break;
    }
    if (!obj)
      refuse(
        `Строка ${i + 1}: у документа basisRef нет объекта расчётов — нужен записанный в 1С документ: ${sets!.map((x) => x.replace("Document_", "")).join(", ")}.`,
      );
    if (obj!["DeletionMark"] === true)
      refuse(`Строка ${i + 1}: объект расчётов документа помечен на удаление.`);
    if (String(obj!["Партнер_Key"] ?? "").toLowerCase() !== String(partner ?? "").toLowerCase())
      refuse(
        `Строка ${i + 1}: документ basisRef — расчёты с другим партнёром, не с контрагентом этой оплаты.`,
      );
    if (String(obj!["Организация_Key"] ?? "").toLowerCase() !== org.toLowerCase())
      refuse(`Строка ${i + 1}: документ basisRef — другой организации.`);
    rows.push({
      ...row,
      LineNumber: i + 1,
      Сумма: lines[i]!.amount,
      ...("СуммаВзаиморасчетов" in row ? { СуммаВзаиморасчетов: lines[i]!.amount } : {}),
      ОбъектРасчетов_Key: obj!["Ref_Key"],
    });
  }
  return rows;
}

const RUB_CODE = "643";
const CASHES = "Catalog_Кассы";
const ORG_BANKS = "Catalog_БанковскиеСчетаОрганизаций";
const PERSONS = "Catalog_ФизическиеЛица";
const CF_ITEMS = "Catalog_СтатьиДвиженияДенежныхСредств";
const DEPARTMENTS = "Catalog_СтруктураПредприятия";
const ASSETS = "ChartOfCharacteristicTypes_СтатьиАктивовПассивов";
const guidOf = (s: string) => s.replace(/[{}']/g, "").trim();
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Касса или счёт организации: по Ref_Key, коду или наименованию среди её собственных; одна — сама. */
async function ownedBy(
  conn: Connection,
  set: string,
  org: string,
  query: string | undefined,
  label: string,
): Promise<string> {
  const hasCode = !!(await conn.getMetadata()).entities.get(set)?.properties.some((p) => p.name === "Code");
  const rows = (
    await conn.client.getCollection<ODataEntity>(
      `${set}${buildQuery({
        filter: `${cmp("Owner_Key", "eq", odataGuid(org))} and DeletionMark eq false`,
        // У касс и счетов организации в УТ кода нет: выбираем только то, что есть в схеме.
        select: ["Ref_Key", "Description", ...(hasCode ? ["Code"] : [])],
      })}`,
    )
  ).value;
  if (!query) {
    if (rows.length === 1) return String(rows[0]!["Ref_Key"]);
    throw new InputError(
      `${label}: у организации ${rows.length ? `их несколько (${rows.map((r) => `«${r["Description"]}»`).join(", ")})` : "нет ни одного"} — укажите. Ничего не записано.`,
    );
  }
  const q = query.trim();
  const hit = rows.filter((r) =>
    GUID.test(q)
      ? String(r["Ref_Key"]).toLowerCase() === guidOf(q).toLowerCase()
      : r["Description"] === q || (hasCode && String(r["Code"] ?? "").trim() === q),
  );
  if (hit.length === 1) return String(hit[0]!["Ref_Key"]);
  throw new InputError(
    `${label} «${q}» ${hit.length ? "найдено несколько" : "не найдено"} среди ${rows.map((r) => `«${r["Description"]}»`).join(", ") || "—"}. Ничего не записано.`,
  );
}

/** Статья доходов или расходов «прочего»: в УТ — план видов характеристик статей или статей активов и пассивов. */
async function articleOf(conn: Connection, which: "income" | "expense", query: string) {
  const own =
    which === "income"
      ? "ChartOfCharacteristicTypes_СтатьиДоходов"
      : "ChartOfCharacteristicTypes_СтатьиРасходов";
  const label = which === "income" ? "Статья доходов" : "Статья расходов";
  for (const set of [own, ASSETS]) {
    if (!(await conn.available()).has(set)) continue;
    try {
      const a = await resolveCatalogItem(conn, [set], label, query);
      return { ref: a.ref, type: `StandardODATA.${set}` };
    } catch (e) {
      if (!(e instanceof InputError) || !/не найдено/.test(e.message)) throw e;
    }
  }
  throw new InputError(
    `${label} «${query}» не найдена ни среди статей ${which === "income" ? "доходов" : "расходов"}, ни среди статей активов и пассивов. Ничего не записано.`,
  );
}

async function nameOf(conn: Connection, set: string, ref: string): Promise<string> {
  try {
    const x = await conn.client.getEntity<ODataEntity>(
      `${set}(${odataGuid(guidOf(ref))})${buildQuery({ select: ["Description"] })}`,
    );
    return String(x["Description"] ?? "");
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found")
      throw new InputError(`Объект ${ref} в ${set} не найден. Ничего не записано.`);
    throw e;
  }
}

const refuse = (text: string): never => {
  throw new InputError(`${text} Ничего не записано.`);
};

/** Создание денежного документа УТ по хозяйственной операции. */
export async function createTradeMoney(
  conn: Connection,
  kind: MoneyKind,
  input: MoneyInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const op = Object.hasOwn(MONEY_OPS[kind], input.operationKind)
    ? MONEY_OPS[kind][input.operationKind]
    : undefined;
  if (!op) {
    const why = Object.hasOwn(EXCLUDED, input.operationKind) ? EXCLUDED[input.operationKind] : undefined;
    refuse(
      `Хозяйственная операция «${input.operationKind}» ${why ? `не поддерживается: ${why}.` : "в этом документе Управления торговлей не поддержана."} ` +
        `Поддержаны: ${Object.keys(MONEY_OPS[kind]).join(", ")}.`,
    );
  }
  const spec = op!;
  const set = MONEY_DOCS[kind];
  const org = await resolveOrgRef(conn, input.organization);
  const rub = (await resolveCatalogItem(conn, ["Catalog_Валюты"], "Справочник «Валюты»", RUB_CODE)).ref;
  const date = await docDate(conn, input.date);
  const day = `${date.slice(0, 10)}T00:00:00`;
  const notes: string[] = [];
  const head: Record<string, unknown> = {
    Date: date,
    Posted: false,
    Организация_Key: org,
    ХозяйственнаяОперация: input.operationKind,
    СуммаДокумента: input.amount,
    Валюта_Key: rub,
    ...(input.comment ? { Комментарий: input.comment } : {}),
  };
  // Касса или счёт документа.
  const isCash = kind === "cashIn" || kind === "cashOut" || (kind === "request" && input.payForm === "cash");
  if (kind === "request") {
    if (!input.payForm)
      refuse("Заявка на расходование: укажите payForm — cash (из кассы) или bank (со счёта).");
    head["Статус"] = input.status ?? "НеСогласована";
    head["ФормаОплатыНаличная"] = isCash;
    head["ФормаОплатыБезналичная"] = !isCash;
    head["ФормаОплатыЗаявки"] = isCash ? "Наличная" : "Безналичная";
    const pay = input.paymentDate ? `${input.paymentDate}T00:00:00` : day;
    head["ЖелательнаяДатаПлатежа"] = pay;
    head["ДатаПлатежа"] = pay;
    if (!input.status)
      notes.push(
        "Заявка создаётся «Не согласована»: согласует её ответственный в 1С (status — чтобы задать иначе).",
      );
  }
  if (isCash) head["Касса_Key"] = await ownedBy(conn, CASHES, org, input.cash, "Касса");
  else
    head["БанковскийСчет_Key"] = await ownedBy(
      conn,
      ORG_BANKS,
      org,
      input.orgBankAccount,
      "Банковский счёт организации",
    );
  if (kind === "bankIn" || kind === "bankOut") {
    // Выписка уже прошла банк: так движение попадает в остаток, а не в «к зачислению/к списанию».
    head["ПроведеноБанком"] = true;
    head["ДатаПроведенияБанком"] = day;
  }
  // Статья ДДС — у всех, кроме перемещений денег (у них её нет в строке; 1С берёт из настройки).
  const cf = input.cashflowItem
    ? (await resolveCatalogItem(conn, [CF_ITEMS], "Статья движения денежных средств", input.cashflowItem)).ref
    : undefined;
  const department = input.department
    ? (await resolveCatalogItem(conn, [DEPARTMENTS], "Подразделение", input.department)).ref
    : undefined;
  if (department) head["Подразделение_Key"] = department;
  if (!spec.noRows && !cf)
    refuse(
      "Укажите cashflowItem — статью движения денежных средств (название, код или Ref_Key); коннектор её не угадывает.",
    );
  if (kind === "bankIn" && input.operationKind === "ПоступлениеОплатыПоПлатежнойКарте" && !cf)
    refuse("Оплата картой: укажите cashflowItem — статью движения денежных средств.");
  const row: Record<string, unknown> = {
    LineNumber: 1,
    Сумма: input.amount,
    ...(cf ? { СтатьяДвиженияДенежныхСредств_Key: cf } : {}),
  };
  // Контрагент и партнёр.
  const needsCp = spec.party || spec.needs?.includes("counterparty");
  if (needsCp && !input.counterpartyRef)
    refuse(`Операция «${input.operationKind}»: укажите counterpartyRef — контрагента.`);
  if (input.counterpartyRef) {
    const cp = guidOf(input.counterpartyRef);
    head["Контрагент_Key"] = cp;
    const partner = await partnerOf(conn, cp);
    if (spec.party) {
      if (!partner) refuse("У контрагента нет партнёра: в Управлении торговлей расчёты ведутся по партнёру.");
      head["Партнер_Key"] = partner;
      Object.assign(row, {
        Партнер_Key: partner,
        ВалютаВзаиморасчетов_Key: rub,
        СуммаВзаиморасчетов: input.amount,
        Организация_Key: org,
      });
    } else if (partner) head["Партнер_Key"] = partner;
    if (input.contractRef) {
      head["Договор"] = guidOf(input.contractRef);
      head["Договор_Type"] = "StandardODATA.Catalog_ДоговорыКонтрагентов";
    }
    if (!input.partyText && (kind === "cashIn" || kind === "cashOut"))
      input = { ...input, partyText: await nameOf(conn, "Catalog_Контрагенты", cp) };
  } else if (input.contractRef) refuse("Договор указан без контрагента (counterpartyRef).");
  if (spec.item) {
    const which = spec.item;
    const q = which === "income" ? input.incomeItem : input.expenseItem;
    if (!q)
      refuse(
        `Операция «${input.operationKind}»: укажите ${which === "income" ? "incomeItem — статью доходов" : "expenseItem — статью расходов"} (или статью активов и пассивов); коннектор её не угадывает.`,
      );
    const a = await articleOf(conn, which, q!);
    const field = which === "income" ? "СтатьяДоходов" : "СтатьяРасходов";
    Object.assign(row, {
      [field]: a.ref,
      [`${field}_Type`]: a.type,
      ВалютаВзаиморасчетов_Key: rub,
      СуммаВзаиморасчетов: input.amount,
      ...(department ? { Подразделение_Key: department } : {}),
    });
  }
  if (spec.needs?.includes("person")) {
    if (!input.person)
      refuse("Выдача подотчётнику: укажите person — физическое лицо (название или Ref_Key).");
    const p = await resolveCatalogItem(conn, [PERSONS], "Физическое лицо", input.person!);
    head["ПодотчетноеЛицо_Key"] = p.ref;
    if (!input.partyText) input = { ...input, partyText: p.name ?? "" };
    Object.assign(row, {
      ВалютаВзаиморасчетов_Key: rub,
      СуммаВзаиморасчетов: input.amount,
      Организация_Key: org,
    });
  }
  if (spec.needs?.includes("fromCash")) {
    if (!input.fromCash) refuse("Укажите fromCash — кассу-отправитель.");
    const from = await ownedBy(conn, CASHES, org, input.fromCash, "Касса-отправитель");
    if (kind === "cashIn") {
      if (from === head["Касса_Key"]) refuse("Касса-отправитель совпадает с кассой документа.");
      head["КассаОтправитель"] = from;
      head["КассаОтправитель_Type"] = `StandardODATA.${CASHES}`;
    } else head["КассаОтправитель_Key"] = from;
  }
  if (spec.needs?.includes("toBank")) {
    if (!input.toBankAccount)
      refuse("Инкассация: укажите toBankAccount — счёт организации, куда сдаются деньги.");
    head["БанковскийСчет_Key"] = await ownedBy(
      conn,
      ORG_BANKS,
      org,
      input.toBankAccount,
      "Счёт для инкассации",
    );
  }
  if (spec.needs?.includes("fromBank")) {
    if (!input.fromBankAccount)
      refuse("Поступление с другого счёта: укажите fromBankAccount — счёт-отправитель организации.");
    const from = await ownedBy(conn, ORG_BANKS, org, input.fromBankAccount, "Счёт-отправитель");
    if (from === head["БанковскийСчет_Key"]) refuse("Счёт-отправитель совпадает со счётом документа.");
    head["БанковскийСчетОтправитель_Key"] = from;
  }
  if (spec.needs?.includes("counterparty") && input.operationKind === "ПоступлениеОплатыПоПлатежнойКарте")
    head["СтатьяДвиженияДенежныхСредств_Key"] = cf;
  if (spec.needs?.includes("budget")) {
    const b = input.budget;
    if (!b?.kbk || !b.oktmo || !b.payerStatus)
      refuse("Перечисление в бюджет: укажите budget — kbk, oktmo и payerStatus (статус составителя).");
    Object.assign(head, {
      ПеречислениеВБюджет: true,
      ВидПеречисленияВБюджет: b!.kind ?? "НалоговыйПлатеж",
      КодБК: b!.kbk,
      КодОКАТО: b!.oktmo,
      СтатусСоставителя: b!.payerStatus,
      ПоказательОснования: b!.basis ?? "0",
      ПоказательПериода: b!.period ?? "0",
      ПоказательНомера: b!.number ?? "0",
      ПоказательДаты: b!.date ?? "0",
    });
  }
  // Тексты ордеров и назначение платежа.
  if (kind === "cashIn" && input.partyText) head["ПринятоОт"] = input.partyText;
  if (kind === "cashOut" && input.partyText) head["Выдать"] = input.partyText;
  if ((kind === "cashIn" || kind === "cashOut") && input.basis) head["Основание"] = input.basis;
  if ((kind === "bankIn" || kind === "bankOut" || kind === "request") && input.purpose)
    head["НазначениеПлатежа"] = input.purpose;
  if (kind === "request" && cf) head["СтатьяДвиженияДенежныхСредств_Key"] = cf;
  const rows = input.lines
    ? await settlementObjectRows(conn, kind, input, row, head["Партнер_Key"], org)
    : [row];
  if (input.lines)
    notes.push(
      `Строк расшифровки: ${rows.length} — по объектам расчётов документов basisRef; проведение разнесёт оплату по ним.`,
    );
  const payload = { ...head, ...(spec.noRows ? {} : { РасшифровкаПлатежа: rows }) };
  return createCore(conn, set, payload, confirm, notes, () => undefined);
}

// ─── Проверка проведения ───

const MAIN_REGISTER: Record<
  MoneyKind,
  { register: string; account?: string; direction: "Receipt" | "Expense" }
> = {
  cashIn: {
    register: "AccumulationRegister_ДенежныеСредстваНаличные",
    account: "Касса_Key",
    direction: "Receipt",
  },
  cashOut: {
    register: "AccumulationRegister_ДенежныеСредстваНаличные",
    account: "Касса_Key",
    direction: "Expense",
  },
  bankIn: {
    register: "AccumulationRegister_ДенежныеСредстваБезналичные",
    account: "БанковскийСчет_Key",
    direction: "Receipt",
  },
  bankOut: {
    register: "AccumulationRegister_ДенежныеСредстваБезналичные",
    account: "БанковскийСчет_Key",
    direction: "Expense",
  },
  request: { register: "AccumulationRegister_ДенежныеСредстваКВыплате", direction: "Receipt" },
};
/** Расчёты с партнёром: оплата клиента уменьшает его долг (расход), оплата поставщику и возврат клиенту — приход. */
const PARTY_REGISTER: Record<Party, { register: string; direction: "Receipt" | "Expense" }> = {
  client: { register: "AccumulationRegister_РасчетыСКлиентами", direction: "Expense" },
  supplier: { register: "AccumulationRegister_РасчетыСПоставщиками", direction: "Receipt" },
  refund: { register: "AccumulationRegister_РасчетыСКлиентами", direction: "Receipt" },
};

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
    // Живая УТ 11.5: у документа без движений набора записей нет — 404.
    if (e instanceof ODataError && e.kind === "not_found") return [];
    throw e;
  }
}

const cents = (n: number) => Math.round(n * 100);

/**
 * Проверка проведения денежного документа УТ: в регистре денег по кассе или счёту документа — движение на его сумму в
 * нужную сторону; у расчётов с партнёром — та же сумма в регистре расчётов; после отмены проведения — ничего.
 * undefined — не денежный документ УТ (проверяют другие).
 */
export async function verifyMoneyMovements(
  conn: Connection,
  entitySet: string,
  ref: string,
  post: boolean,
): Promise<Record<string, unknown> | undefined> {
  const kind = (Object.keys(MONEY_DOCS) as MoneyKind[]).find((k) => MONEY_DOCS[k] === entitySet);
  if (!kind || !(await tradeMoneySet(conn, kind))) return undefined;
  const guid = guidOf(ref);
  try {
    const doc = await conn.client.getEntity<ODataEntity>(`${entitySet}(${odataGuid(guid)})?$format=json`);
    const opName = String(doc["ХозяйственнаяОперация"]);
    const spec = Object.hasOwn(MONEY_OPS[kind], opName) ? MONEY_OPS[kind][opName] : undefined;
    const main = MAIN_REGISTER[kind];
    // Заявка денег не двигает и расчётов не трогает (на копии — только «к выплате»).
    const party = spec?.party && kind !== "request" ? PARTY_REGISTER[spec.party] : undefined;
    const registers = [main.register, ...(party ? [party.register] : [])];
    const found: Record<string, ODataEntity[]> = {};
    for (const r of [...new Set(registers)]) found[r] = await recordsOf(conn, r, entitySet, guid);
    const mismatches: string[] = [];
    let foreign = false;
    if (!post) {
      for (const [r, rows] of Object.entries(found))
        if (rows.length) mismatches.push(`после отмены проведения в ${r} осталось записей: ${rows.length}`);
    } else if (!spec) {
      return {
        movementsNote: `Хозяйственная операция «${String(doc["ХозяйственнаяОперация"])}» коннектором не поддержана — сверка движений не выполнялась.`,
      };
    } else {
      const amount = cents(Number(doc["СуммаДокумента"] ?? 0));
      const sum = (rows: ODataEntity[], dir: string, filter?: (r: ODataEntity) => boolean) =>
        rows
          .filter((r) => r["RecordType"] === dir && (!filter || filter(r)))
          .reduce((a, r) => a + cents(Number(r["Сумма"] ?? 0)), 0);
      const acc = main.account ? String(doc[main.account] ?? "").toLowerCase() : "";
      const own = (r: ODataEntity) => !main.account || String(r[main.account] ?? "").toLowerCase() === acc;
      const got = sum(found[main.register]!, main.direction, own);
      if (got !== amount)
        mismatches.push(
          `${main.register}: ${main.direction === "Receipt" ? "приход" : "расход"} ${got / 100}, в документе ${amount / 100}`,
        );
      const other = found[main.register]!.filter(
        (r) => own(r) && r["RecordType"] !== main.direction && cents(Number(r["Сумма"] ?? 0)) !== 0,
      );
      if (other.length)
        mismatches.push(`${main.register}: лишние движения в обратную сторону: ${other.length}`);
      if (party && settlementInDocCurrency(found[party.register]!, doc)) {
        const p = sum(found[party.register]!, party.direction);
        if (p !== amount) mismatches.push(`${party.register}: ${p / 100}, в документе ${amount / 100}`);
      } else if (party) foreign = true;
    }
    const matched = !mismatches.length;
    return {
      ...(foreign ? { movementsNote: FOREIGN_NOTE } : {}),
      movementsVerified: matched,
      movements: {
        register: registers.join(", "),
        found: Object.values(found).reduce((a, r) => a + r.length, 0),
        ...(mismatches.length ? { mismatches } : {}),
      },
      ...(matched
        ? {}
        : {
            verified: false,
            note: post
              ? "Документ проведён, но движения денег или расчётов не совпали с его суммой. Проверьте документ в 1С; отменить проведение — write.document.post_document с post=false."
              : "Проведение отменено, но движения документа остались. Проверьте его в 1С.",
          }),
    };
  } catch (e) {
    return {
      movementsVerified: false,
      verified: false,
      note: `Действие выполнено, но движения денег прочитать не удалось (${e instanceof Error ? e.message : String(e)}). Проверьте документ в 1С, прежде чем повторять.`,
    };
  }
}

/**
 * Расчёты с партнёром в валюте документа: «Сумма» регистра расчётов — в валюте взаиморасчётов (Валюта_Key), и при
 * договоре в у.е. она законно отличается от суммы документа (ревью выпуска 5). Валюта не опубликована — считаем, что
 * в валюте документа, как раньше.
 */
export function settlementInDocCurrency(rows: readonly ODataEntity[], doc: ODataEntity): boolean {
  const docCur = String(doc["Валюта_Key"] ?? "").toLowerCase();
  if (!docCur || /^0{8}-/.test(docCur)) return true;
  return rows.every((r) => {
    const c = String(r["Валюта_Key"] ?? "").toLowerCase();
    return !c || /^0{8}-/.test(c) || c === docCur;
  });
}
export const FOREIGN_NOTE =
  "Расчёты с партнёром — в валюте договора, отличной от валюты документа: сумма в расчётах не сверялась — по курсу коннектор её не пересчитывает.";

/** Платёжное поручение Бухгалтерии в УТ — заявка на расходование и списание безналичных. */
export const PAYOUT_ORDER_IN_TRADE =
  "В Управлении торговлей платёжного поручения как документа Бухгалтерии нет: оплату оформляют заявкой на " +
  "расходование (write.money.create_payment_request) и списанием безналичных (write.money.create_bank_writeoff). " +
  "Ничего не записано.";

// ─── Копия и правка полей ───

/** Поля шапки денежного документа УТ, которые можно менять (update_entity) и задавать копии (copy_document). */
const MONEY_EDITABLE = new Set([
  "Комментарий",
  "Date",
  "СуммаДокумента",
  "НазначениеПлатежа",
  "Основание",
  "ПринятоОт",
  "Выдать",
  "Подразделение_Key",
  "Статус",
  "ЖелательнаяДатаПлатежа",
  "ДатаПлатежа",
  "ДатаПроведенияБанком",
]);
const moneyWhitelistRefusal = (bad: string[]) =>
  `Поля ${bad.join(", ")} в денежном документе Управления торговлей коннектор не меняет: от них зависят расчёты и ` +
  `объект расчётов, который 1С заполняет сама. Можно: ${[...MONEY_EDITABLE].join(", ")}; другое — новый документ. ` +
  "Ничего не записано.";

/** Кинд денежного документа УТ, если коннектор умеет его править; undefined — Бухгалтерия, как раньше. */
export async function tradeMoneyEditKind(
  conn: Connection,
  entitySet: string,
): Promise<MoneyKind | undefined> {
  const kind = (Object.keys(MONEY_DOCS) as MoneyKind[]).find((k) => MONEY_DOCS[k] === entitySet);
  return kind && (await tradeMoneySet(conn, kind)) ? kind : undefined;
}

/** Сумма документа и одной строки расшифровки меняются вместе; строк несколько — разнести сумму нельзя. */
/** Процент ставки НДС по перечислению справочника (НДС22, НДС20_120 → 22, 20; БезНДС, НДС0 → 0). */
async function vatPercent(conn: Connection, ref: string): Promise<number | undefined> {
  const r = await conn.client
    .getEntity<ODataEntity>(
      `Catalog_СтавкиНДС(${odataGuid(ref)})${buildQuery({ select: ["ПеречислениеСтавкаНДС"] })}`,
    )
    .catch(() => undefined);
  const code = String(r?.["ПеречислениеСтавкаНДС"] ?? "");
  if (code === "БезНДС" || code === "НДС0") return 0;
  const m = /^НДС(\d+)(?:_\d+)?$/.exec(code);
  return m ? Number(m[1]) : undefined;
}

/**
 * Новая сумма документа — в единственную строку расшифровки; НДС строки пересчитывается по её ставке, как 1С в форме
 * при смене суммы: сумма × ставка / (100 + ставка), сумма платежа включает НДС (на копии рабочей УТ 11.5 так
 * посчитаны 83 из 90 строк с НДС, остальные — НДС, перенесённый из счёта поставщика). Предпросмотр показывает
 * «было → станет». Взаиморасчёты в другой валюте и несколько строк — отказ.
 */
async function withAmount(
  conn: Connection,
  doc: ODataEntity,
  fields: Record<string, unknown>,
): Promise<{ fields: Record<string, unknown>; note?: string } | string> {
  if (!("СуммаДокумента" in fields)) return { fields };
  const amount = Number(fields["СуммаДокумента"]);
  if (!(amount > 0)) return "Сумма документа должна быть больше нуля.";
  const rows = (doc["РасшифровкаПлатежа"] as ODataEntity[] | undefined) ?? [];
  if (rows.length > 1)
    return "В расшифровке платежа несколько строк: какую сумму в какую строку — не определить. Измените документ в 1С.";
  if (!rows.length) return { fields };
  const row = Object.fromEntries(
    Object.entries(rows[0]!).filter(([k]) => !k.includes("@") && k !== "Ref_Key"),
  );
  // Взаиморасчёты в другой валюте (у.е.) по новой сумме коннектор не пересчитывает (ревью выпуска 4).
  const settleCur = String(row["ВалютаВзаиморасчетов_Key"] ?? "");
  const docCur = String(doc["Валюта_Key"] ?? "");
  if (settleCur && !/^0{8}-/.test(settleCur) && docCur && settleCur !== docCur)
    return "Взаиморасчёты по строке — в другой валюте: сумму взаиморасчётов по курсу коннектор не пересчитывает. Измените документ в 1С.";
  const out: Record<string, unknown> = {
    ...row,
    Сумма: amount,
    ...("СуммаВзаиморасчетов" in row ? { СуммаВзаиморасчетов: amount } : {}),
  };
  let note: string | undefined;
  const rateRef = String(row["СтавкаНДС_Key"] ?? "");
  const oldVat = Number(row["СуммаНДС"] ?? 0);
  if (rateRef && !/^0{8}-/.test(rateRef)) {
    const p = await vatPercent(conn, rateRef);
    if (p === undefined)
      return "Ставка НДС строки расшифровки не распознана — НДС по новой сумме не пересчитать. Измените документ в 1С.";
    const vat = Math.round(Number(((amount * p * 100) / (100 + p)).toFixed(6))) / 100;
    out["СуммаНДС"] = vat;
    if (vat !== oldVat) note = `НДС строки пересчитан по ставке ${p}% от новой суммы: ${oldVat} → ${vat}.`;
  } else if (oldVat !== 0)
    return "В строке расшифровки есть НДС без ставки — пересчитать его не по чему. Измените документ в 1С.";
  return { fields: { ...fields, РасшифровкаПлатежа: [out] }, ...(note ? { note } : {}) };
}

/**
 * Смена даты документа служебные даты не трогает: у заявки дата платежа бывает позже даты документа, у выписки дата
 * банка — своя (второй раунд ревью выпуска 4: перенос затирал будущую дату платежа). Только заметка; задать их можно
 * явно — поля есть в белом списке.
 */
function dateNote(entitySet: string, fields: Record<string, unknown>): string | undefined {
  if (typeof fields["Date"] !== "string") return undefined;
  if (
    (entitySet === MONEY_DOCS.bankIn || entitySet === MONEY_DOCS.bankOut) &&
    !("ДатаПроведенияБанком" in fields)
  )
    return "Дата «проведено банком» не изменена — при необходимости задайте ДатаПроведенияБанком.";
  if (entitySet === MONEY_DOCS.request && !("ДатаПлатежа" in fields || "ЖелательнаяДатаПлатежа" in fields))
    return "Даты платежа заявки не изменены — при необходимости задайте ЖелательнаяДатаПлатежа и ДатаПлатежа.";
  return undefined;
}

/** Изменение полей шапки по белому списку, только у непроведённого документа. */
export async function updateTradeMoney(
  conn: Connection,
  entitySet: string,
  ref: string,
  fields: Record<string, unknown>,
  confirm: boolean,
  patch: (fields: Record<string, unknown>) => Promise<CallToolResult>,
): Promise<CallToolResult> {
  const bad = Object.keys(fields).filter((k) => !MONEY_EDITABLE.has(k));
  if (bad.length) throw new InputError(moneyWhitelistRefusal(bad));
  const doc = await conn.client.getEntity<ODataEntity>(
    `${entitySet}(${odataGuid(guidOf(ref))})?$format=json`,
  );
  if (doc["Posted"] === true)
    throw new InputError(
      "Документ проведён. Сначала отмените проведение (write.document.post_document с post=false), затем меняйте его. Ничего не записано.",
    );
  const amounted = await withAmount(conn, doc, fields);
  if (typeof amounted === "string") throw new InputError(`${amounted} Ничего не записано.`);
  const full = amounted.fields;
  const extra = [dateNote(entitySet, fields), amounted.note].filter((n): n is string => !!n);
  const patched = await patch(full);
  const result =
    extra.length && !patched.isError && patched.structuredContent
      ? {
          ...patched,
          structuredContent: {
            ...(patched.structuredContent as Record<string, unknown>),
            notes: [
              ...(((patched.structuredContent as Record<string, unknown>)["notes"] as string[] | undefined) ??
                []),
              ...extra,
            ],
          },
        }
      : patched;
  const sc = result.structuredContent as Record<string, unknown> | undefined;
  if (!confirm || result.isError || !sc || sc["dryRun"] || !("РасшифровкаПлатежа" in full)) return result;
  // Общая проверка правки табличные части не сверяет: сумму строки расшифровки перечитываем сами.
  const after = await conn.client.getEntity<ODataEntity>(
    `${entitySet}(${odataGuid(guidOf(ref))})?$format=json`,
  );
  const want = Number(full["СуммаДокумента"]);
  const rows = (after["РасшифровкаПлатежа"] as ODataEntity[] | undefined) ?? [];
  const sentVat = ((full["РасшифровкаПлатежа"] as ODataEntity[])[0] ?? {})["СуммаНДС"];
  const ok =
    rows.length === 1 &&
    cents(Number(rows[0]!["Сумма"])) === cents(want) &&
    cents(Number(after["СуммаДокумента"])) === cents(want) &&
    (sentVat === undefined || cents(Number(rows[0]!["СуммаНДС"] ?? 0)) === cents(Number(sentVat)));
  return {
    ...result,
    // Расшифровку сверили здесь: общая заметка «не сверялись (табличные части)» была бы неправдой.
    structuredContent: ok
      ? { ...Object.fromEntries(Object.entries(sc).filter(([k]) => k !== "readbackNote")), verified: true }
      : {
          ...sc,
          verified: false,
          note: `1С сохранила сумму иначе: в документе ${String(after["СуммаДокумента"])}, в расшифровке ${rows.map((r) => `${String(r["Сумма"])} (НДС ${String(r["СуммаНДС"] ?? 0)})`).join(", ")}. Проверьте документ в 1С.`,
        },
  };
}

/**
 * Поля, которые копия не берёт у образца: объект расчётов и связи с исходными документами (1С заполнит объект
 * расчётов сама, как при создании), данные банковской выписки.
 */
const MONEY_NOT_COPIED = new Set([
  "ДокументОснование",
  "ДокументОснование_Type",
  "ЗаявкаНаРасходованиеДенежныхСредств",
  "ЗаявкаНаРасходованиеДенежныхСредств_Type",
  "ОбъектРасчетов_Key",
  "ДанныеВыписки",
  "ДатаЗагрузки",
  "ДатаВыгрузки",
  "ИдентификаторПлатежа",
  "ИдентификаторДокумента",
  "НомерВходящегоДокумента",
  "ДатаВходящегоДокумента",
  "Автор_Key",
  "ДополнительныеРеквизиты",
]);
const MONEY_ROW_NOT_COPIED = new Set([
  "ОбъектРасчетов_Key",
  "ИдентификаторСтроки",
  "ОснованиеПлатежа",
  "ОснованиеПлатежа_Type",
  "ЗаявкаНаРасходованиеДенежныхСредств",
  "ЗаявкаНаРасходованиеДенежныхСредств_Type",
]);

/** Копия денежного документа УТ: непроведённая, с новой датой, без объекта расчётов и связей образца. */
export async function copyTradeMoney(
  conn: Connection,
  entitySet: string,
  kind: MoneyKind,
  ref: string,
  input: {
    date?: string | undefined;
    fields?: Record<string, unknown> | undefined;
    lines?: unknown[] | undefined;
  },
  confirm: boolean,
  strip: (e: ODataEntity) => Record<string, unknown>,
): Promise<CallToolResult> {
  if (input.lines?.length)
    throw new InputError(
      "У денежного документа строк товаров нет: правки lines не применяются; сумму задайте в fields.СуммаДокумента. Ничего не записано.",
    );
  const bad = Object.keys(input.fields ?? {}).filter((k) => !MONEY_EDITABLE.has(k));
  if (bad.length) throw new InputError(moneyWhitelistRefusal(bad));
  const source = await conn.client.getEntity<ODataEntity>(
    `${entitySet}(${odataGuid(guidOf(ref))})?$format=json`,
  );
  const srcOp = String(source["ХозяйственнаяОперация"]);
  const spec = Object.hasOwn(MONEY_OPS[kind], srcOp) ? MONEY_OPS[kind][srcOp] : undefined;
  if (!spec)
    throw new InputError(
      `Хозяйственная операция образца «${String(source["ХозяйственнаяОперация"])}» коннектором не поддержана — копию не создаю. Ничего не записано.`,
    );
  const base = strip(source);
  const payload: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(base)) if (!MONEY_NOT_COPIED.has(k)) payload[k] = v;
  payload["РасшифровкаПлатежа"] = (
    (base["РасшифровкаПлатежа"] as Record<string, unknown>[] | undefined) ?? []
  ).map((r) => Object.fromEntries(Object.entries(r).filter(([k]) => !MONEY_ROW_NOT_COPIED.has(k))));
  const date = await docDate(conn, input.date);
  payload["Date"] = date;
  payload["Posted"] = false;
  const day = `${date.slice(0, 10)}T00:00:00`;
  if (kind === "bankIn" || kind === "bankOut") {
    payload["ПроведеноБанком"] = true;
    payload["ДатаПроведенияБанком"] = day;
  }
  if (kind === "request") {
    payload["ЖелательнаяДатаПлатежа"] = day;
    payload["ДатаПлатежа"] = day;
    payload["Статус"] = "НеСогласована";
  }
  const full = await withAmount(
    conn,
    { ...source, РасшифровкаПлатежа: payload["РасшифровкаПлатежа"] },
    input.fields ?? {},
  );
  if (typeof full === "string") throw new InputError(`${full} Ничего не записано.`);
  Object.assign(payload, full.fields);
  if (!(payload["РасшифровкаПлатежа"] as unknown[]).length) delete payload["РасшифровкаПлатежа"];
  return createCore(
    conn,
    entitySet,
    payload,
    confirm,
    [
      "Копия — без объекта расчётов и связей образца (основание, заявка): 1С заполнит объект расчётов при записи.",
      ...(full.note ? [full.note] : []),
      ...(kind === "request" && !("Статус" in (input.fields ?? {}))
        ? ["Копия заявки — «Не согласована»."]
        : []),
    ],
    () => undefined,
  );
}
