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
  const op = MONEY_OPS[kind][input.operationKind];
  if (!op) {
    const why = EXCLUDED[input.operationKind];
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
  const payload = { ...head, ...(spec.noRows ? {} : { РасшифровкаПлатежа: [row] }) };
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
    const spec = MONEY_OPS[kind][String(doc["ХозяйственнаяОперация"])];
    const main = MAIN_REGISTER[kind];
    // Заявка денег не двигает и расчётов не трогает (на копии — только «к выплате»).
    const party = spec?.party && kind !== "request" ? PARTY_REGISTER[spec.party] : undefined;
    const registers = [main.register, ...(party ? [party.register] : [])];
    const found: Record<string, ODataEntity[]> = {};
    for (const r of [...new Set(registers)]) found[r] = await recordsOf(conn, r, entitySet, guid);
    const mismatches: string[] = [];
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
      if (party) {
        const p = sum(found[party.register]!, party.direction);
        if (p !== amount) mismatches.push(`${party.register}: ${p / 100}, в документе ${amount / 100}`);
      }
    }
    const matched = !mismatches.length;
    return {
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

/** Платёжное поручение Бухгалтерии в УТ — заявка на расходование и списание безналичных. */
export const PAYOUT_ORDER_IN_TRADE =
  "В Управлении торговлей платёжного поручения как документа Бухгалтерии нет: оплату оформляют заявкой на " +
  "расходование (write.money.create_payment_request) и списанием безналичных (write.money.create_bank_writeoff). " +
  "Ничего не записано.";
