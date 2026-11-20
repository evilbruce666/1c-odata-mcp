import type { CallToolResult } from "@modelcontextprotocol/server";
import type { Connection } from "../context.js";
import type { ODataEntity } from "../types/odata.js";
import { ODataError } from "../odata/errors.js";
import { EMPTY_REF } from "../odata/pricing.js";
import { buildQuery } from "../odata/query.js";
import { fail } from "./_shared.js";
import { createCore, resolveCatalogItem, resolveOrgRef } from "./write.js";
import { previewDay } from "./trade-counterparty.js";

/**
 * Договор и соглашение Управления торговлей. Значения перечислений в OData не опубликованы — взяты с копии рабочей
 * УТ 11.5 чтением (20.11.2026): у договоров и соглашений, созданных в 1С, эти поля заполнены одинаково почти везде.
 * Через OData 1С их не заполняет (нет «заполнения по умолчанию»), поэтому коннектор ставит те же значения.
 */

const CONTRACT_KIND: Record<string, { type: string; operation: string } | undefined> = {
  СПокупателем: { type: "СПокупателем", operation: "РеализацияКлиенту" },
  СПоставщиком: { type: "СПоставщиком", operation: "ЗакупкаУПоставщика" },
  СКомитентом: { type: "СКомитентом", operation: "ПриемНаКомиссию" },
  СКомиссионером: { type: "СКомиссионером", operation: "ПередачаНаКомиссию" },
};

export const SETTLEMENT_ORDER = {
  contracts: "ПоДоговорамКонтрагентов",
  invoices: "ПоНакладным",
  orders: "ПоЗаказам",
} as const;

/** Партнёр контрагента: в УТ договор и соглашение ссылаются на обоих. */
async function counterpartyPartner(
  conn: Connection,
  counterpartyRef: string,
): Promise<{ partner: string; name: string } | undefined> {
  const guid = counterpartyRef.replace(/[{}']/g, "");
  try {
    const c = await conn.client.getEntity<ODataEntity>(
      `Catalog_Контрагенты(guid'${guid}')${buildQuery({ select: ["Партнер_Key", "Description", "DeletionMark"] })}`,
    );
    const partner = String(c["Партнер_Key"] ?? "");
    if (!partner || partner === EMPTY_REF || c["DeletionMark"] === true) return undefined;
    return { partner, name: String(c["Description"] ?? "") };
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found") return undefined;
    throw e;
  }
}

const rub = (conn: Connection, currency: string | undefined) =>
  resolveCatalogItem(conn, ["Catalog_Валюты"], "Справочник «Валюты»", currency ?? "643");

export interface TradeContractInput {
  organization?: string | undefined;
  counterpartyRef: string;
  kind: string;
  number?: string | undefined;
  date?: string | undefined;
  name?: string | undefined;
  currency?: string | undefined;
  settlementOrder?: keyof typeof SETTLEMENT_ORDER | undefined;
  ignored: string[];
}

export async function createTradeContract(
  conn: Connection,
  set: string,
  input: TradeContractInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const kind = CONTRACT_KIND[input.kind];
  if (!kind)
    return fail(
      `Вид договора «${input.kind}» в Управлении торговлей этим инструментом не создаётся: доступны СПокупателем, ` +
        "СПоставщиком, СКомитентом, СКомиссионером. Ничего не создано.",
    );
  const owner = await counterpartyPartner(conn, input.counterpartyRef);
  if (!owner)
    return fail(
      `Контрагент ${input.counterpartyRef} не найден, помечен на удаление или без партнёра. Ничего не создано.`,
    );
  const org = await resolveOrgRef(conn, input.organization);
  const cur = await rub(conn, input.currency);
  const num = input.number ?? "б/н";
  const notes = input.ignored.length
    ? [`Не записано — в договоре Управления торговлей таких полей нет: ${input.ignored.join(", ")}.`]
    : [];
  const payload = {
    Description: input.name ?? `Договор ${num}`,
    Номер: num,
    Дата: input.date ? `${input.date}T00:00:00` : await previewDay(conn),
    Партнер_Key: owner.partner,
    Контрагент_Key: input.counterpartyRef.replace(/[{}']/g, ""),
    Организация_Key: org,
    ВалютаВзаиморасчетов_Key: cur.ref,
    ТипДоговора: kind.type,
    ХозяйственнаяОперация: kind.operation,
    Статус: "Действует",
    Согласован: true,
    ПорядокРасчетов: SETTLEMENT_ORDER[input.settlementOrder ?? "contracts"],
    НалогообложениеНДСОпределяетсяВДокументе: true,
    ЗакупкаПодДеятельностьОпределяетсяВДокументе: true,
    ВариантПриемкиТоваров: "РазделенаТолькоПоНакладным",
    СпособДоставки: "ОпределяетсяВРаспоряжении",
    ВариантКурсаДоговора: "Переменный",
    ВариантОформленияЗакупок: "НеРазделять",
    ПорядокОтраженияНеотфактурованныхПоставокНУ: "ОтражатьПриПоступлении",
    ВариантОпределенияГрафикаИсполненияВДоговоре: "ЗаказыНакладные",
  };
  return createCore(conn, set, payload, confirm, notes, () => undefined);
}

export const AGREEMENT_SETS = {
  client: "Catalog_СоглашенияСКлиентами",
  supplier: "Catalog_СоглашенияСПоставщиками",
} as const;

export interface AgreementInput {
  side: keyof typeof AGREEMENT_SETS;
  organization?: string | undefined;
  counterpartyRef: string;
  name: string;
  number?: string | undefined;
  date?: string | undefined;
  currency?: string | undefined;
  priceKindRef?: string | undefined;
  pricesIncludeVat?: boolean | undefined;
  useContracts?: boolean | undefined;
}

export async function createAgreement(
  conn: Connection,
  input: AgreementInput,
  confirm: boolean,
): Promise<CallToolResult> {
  const set = AGREEMENT_SETS[input.side];
  if (!(await conn.available()).has(set))
    return fail(
      `Справочник ${set} не опубликован в OData: соглашения есть в Управлении торговлей, КА и ERP. Ничего не создано.`,
    );
  const owner = await counterpartyPartner(conn, input.counterpartyRef);
  if (!owner)
    return fail(
      `Контрагент ${input.counterpartyRef} не найден, помечен на удаление или без партнёра. Ничего не создано.`,
    );
  if (input.side === "supplier" && input.priceKindRef)
    return fail("priceKindRef — только для соглашения с клиентом (вид цен продажи). Ничего не создано.");
  const org = await resolveOrgRef(conn, input.organization);
  const cur = await rub(conn, input.currency);
  const day = input.date ? `${input.date}T00:00:00` : await previewDay(conn);
  const common = {
    Description: input.name,
    ...(input.number ? { Номер: input.number } : {}),
    Партнер_Key: owner.partner,
    Контрагент_Key: input.counterpartyRef.replace(/[{}']/g, ""),
    Организация_Key: org,
    Валюта_Key: cur.ref,
    Статус: "Действует",
    Согласован: true,
    ЦенаВключаетНДС: input.pricesIncludeVat ?? false,
    ИспользуютсяДоговорыКонтрагентов: input.useContracts ?? true,
  };
  const payload =
    input.side === "client"
      ? {
          ...common,
          Дата: day,
          ХозяйственнаяОперация: "РеализацияКлиенту",
          Типовое: false,
          ВидСоглашенияДляОграниченияЧтения: "Индивидуальные",
          ВидСоглашенияДляОграниченияИзменения: "Индивидуальные",
          ...(input.priceKindRef
            ? { ВидЦен_Key: input.priceKindRef.replace(/[{}']/g, ""), ВариантРасчетаЦен: "ВидЦенВШапке" }
            : { ВариантРасчетаЦен: "БезЦен" }),
        }
      : {
          ...common,
          ДатаНачалаДействия: day,
          ХозяйственнаяОперация: "ЗакупкаУПоставщика",
          ВариантПриемкиТоваров: "РазделенаТолькоПоНакладным",
          СпособДоставки: "СиламиПоставщикаДоНашегоСклада",
        };
  return createCore(conn, set, payload, confirm, [], () => undefined);
}
