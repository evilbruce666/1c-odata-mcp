import { registerExistingBankSettlementTools } from "../tools/existing-bank-settlement.js";
import { registerCounterpartyGroupTools } from "../tools/counterparty-groups.js";
import { registerSettlementSettingsTools } from "../tools/settlement-settings.js";
import { readFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/server";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { createResultSchema } from "../schemas/output.js";
import { fingerprintWriteInput, withWriteOperation } from "../odata/write-operation-context.js";
import { callAbortError, withToolCall } from "../odata/call-context.js";
import { registerMetaTools } from "../tools/meta.js";
import { registerCounterpartyTools } from "../tools/counterparties.js";
import { registerDocumentTools } from "../tools/documents.js";
import { registerRegisterTools } from "../tools/registers.js";
import { registerStockTools } from "../tools/stock.js";
import { registerItemSalesTools } from "../tools/item-sales.js";
import { registerPartnerSettlementTools } from "../tools/partner-settlements.js";
import { registerItemTools } from "../tools/items.js";
import { registerUserTools } from "../tools/users.js";
import { registerPricingTools } from "../tools/pricing.js";
import { registerWarehouseTools } from "../tools/warehouses.js";
import { tradeWithoutAccounting, withoutAccounting } from "../odata/configuration.js";
import { PRICE_DOC, priceMechanisms } from "../odata/pricing.js";
import { PARTNERS, tradeCounterpartySchema } from "../tools/trade-counterparty.js";
import { tradeNomenclatureSchema } from "../tools/trade-nomenclature.js";
import { STOCK_DOCS, tradeStockSchema } from "../tools/trade-warehouse.js";
import { MONEY_DOCS, tradeMoneySet, type MoneyKind } from "../tools/trade-money.js";
import { SALES_DOCS, tradeSalesSet, type SalesKind } from "../tools/trade-sales.js";
import { registerDebtTools } from "../tools/debts.js";
import { registerMonthCloseTools } from "../tools/month-close.js";
import { registerCardTools } from "../tools/card.js";
import { registerPrompts } from "./prompts.js";
import { APPS_EXTENSION_ID, appMetaFor, registerApps } from "./apps.js";
import { registerCashflowTools } from "../tools/cashflow.js";
import { registerSalesTools } from "../tools/sales.js";
import { registerOrganizationTools } from "../tools/organization.js";
import { registerWriteTools } from "../tools/write.js";
import { registerAuditTools } from "../tools/audit.js";
import { registerFreshnessTools } from "../tools/freshness.js";
import {
  READ_HINTS,
  WRITE_HINTS,
  DESTRUCTIVE_HINTS,
  MAX_RESPONSE_CHARS,
  RESPONSE_CHARS_NORM,
  elementChars,
  fail,
  fitToChars,
  guard,
  offsetField,
  ok,
  pageCharsLimit,
  renderJson,
  responseCharsLimit,
  toolListed,
} from "../tools/_shared.js";

/** Версия берётся из package.json (в собранном пакете он на два уровня выше dist/mcp/). */
function readVersion(): string {
  try {
    const url = new URL("../../package.json", import.meta.url);
    const pkg = JSON.parse(readFileSync(url, "utf8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

const INSTRUCTIONS =
  "Доступ к данным 1С:Предприятие через OData. По умолчанию только чтение " +
  "(аналитика: дебиторка, остатки, продажи, движение денег; справочники и документы). " +
  "Запись включается отдельно и по умолчанию работает в режиме предпросмотра (dry-run): " +
  "сначала показывайте пользователю, что будет создано/изменено, и выполняйте запись " +
  "только после явного согласия (confirm=true). У инструментов есть параметр database " +
  "(см. read.system.list_databases) и organization (см. read.system.list_organizations). " +
  "Для создания объекта при подтверждении передавайте operationId из предпросмотра; " +
  "при таймауте повторяйте подтверждение с тем же id и теми же аргументами, а не с новым предпросмотром. " +
  "Это защищает от дубликата, если ответ 1С потерялся. " +
  "Если настроено несколько баз и из вопроса неясно какая, спросите пользователя или проверьте " +
  "read.system.data_freshness, а не полагайтесь на базу по умолчанию: при её подстановке в ответе чтения есть " +
  "databaseDefaulted, а запись без database при нескольких базах отказывает. " +
  "Ref_Key действуют только в своей базе: чужая ссылка даёт пустой результат, а не ошибку. " +
  "Длинный ответ приходит частями: при truncated=true и nextOffset читайте дальше с offset=nextOffset; truncated без nextOffset — читать дальше нечего, уточните отбор. " +
  "Вызывайте инструменты по одному: 1С обрабатывает запросы одной учётной записи по очереди, и при параллельных " +
  "вызовах тяжёлые (check_month_close, data_freshness, акт сверки, карточка) упираются в таймаут клиента, а соседние " +
  "получают Connection closed. После такой ошибки повторите вызов один, это не значит, что данных нет.";

const LISTED_TOOLS_NOTE =
  "Список инструментов на этом сервере ограничен настройкой TOOLS_ALLOW: если ответ или описание советует " +
  "инструмент, которого нет в вашем списке, он недоступен, скажите об этом пользователю.";

/**
 * Правки документа, которые нельзя слепо повторять: инструмент перечитывает документ, и повтор
 * добавил бы ещё строку или удалил соседнюю. Как и создание, защищены operationId и журналом.
 */
/** Правки существующего банковского документа: итог несёт readback, повтор — квитанцию журнала. */
const BANK_DOCUMENT_TOOLS = new Set([
  "write.document.apply_bank_settlement_settings",
  "write.document.align_bank_payment_amounts",
]);
const JOURNALED_UPDATE_TOOLS = new Set([
  "write.document.apply_bank_settlement_settings",
  "write.document.align_bank_payment_amounts",
  "write.document.post_document",
  "write.accounting.set_contract_settlement_settings",
  "write.accounting.set_settlement_settings",
  "write.counterparty.create_group",
  "write.counterparty.set_group",
  "write.document.add_document_line",
  "write.document.remove_document_line",
]);

/** Подсказки клиенту о характере инструмента — по имени (одна точка вместо правок 55 конфигов). */
function annotationsFor(name: string) {
  if (name === "write.document.post_document" || name === "write.entity.mark_for_deletion")
    return DESTRUCTIVE_HINTS;
  if (name.startsWith("read.")) return READ_HINTS;
  return WRITE_HINTS;
}

/**
 * Ответ чтения всегда называет базу, а подставленная по умолчанию при нескольких настроенных базах
 * помечена: пустой результат в чужой базе выглядит как «данных нет». Опрашивать другие базы здесь
 * нельзя: это лишние запросы, разные юрлица и ложные подсказки там, где пусто законно.
 */
export function withDatabaseInfo(
  ctx: ServerContext,
  args: Record<string, unknown>,
  result: CallToolResult,
): CallToolResult {
  const data = result.structuredContent;
  if (result.isError || !data || typeof data !== "object") return result;
  const requested = typeof args.database === "string" ? args.database : undefined;
  const name = ctx.db(requested).cfg.name;
  const others = ctx
    .databases()
    .map((d) => d.name)
    .filter((n) => n !== name);
  const structuredContent = {
    database: name,
    ...data,
    ...(!requested && others.length
      ? {
          databaseDefaulted: true,
          databaseNote: `Использована база по умолчанию ${name}. Также доступны: ${others.join(", ")}.`,
        }
      : {}),
  };
  const text = renderJson(data);
  const content = result.content.map((item) =>
    item.type === "text" && item.text === text ? { ...item, text: renderJson(structuredContent) } : item,
  );
  return { ...result, content, structuredContent };
}

/** Список страницы, если в ответе есть другие списки; countIsTotal — count означает «всего», а не «в этом ответе». */
type PageList = { list: string; countIsTotal?: boolean };

/** Угадывать по длине нельзя: список счетов в ответе остатков бывает длиннее самих остатков. */
const PAGE_LISTS: Record<string, PageList> = {
  "read.analytics.get_inventory": { list: "items", countIsTotal: true },
  "read.analytics.get_warehouse_stock": { list: "items", countIsTotal: true },
  "read.analytics.get_sales_register": { list: "items", countIsTotal: true },
  "read.analytics.get_partner_settlements": { list: "partners", countIsTotal: true },
  "read.catalog.list_warehouses": { list: "items", countIsTotal: true },
  "read.catalog.list_items": { list: "items", countIsTotal: true },
  "read.analytics.get_prices": { list: "items", countIsTotal: true },
  "read.analytics.get_debtors": { list: "debtors", countIsTotal: true },
  "read.analytics.get_creditors": { list: "creditors", countIsTotal: true },
  "read.analytics.get_unpaid_invoices": { list: "invoices", countIsTotal: true },
};

type NestedList = { path: string; holder: unknown[]; chars: number };

/**
 * Самый длинный вложенный список: обрезка верхнего уровня его не видит, а элемент с большим списком внутри
 * (проверка закрытия месяца, табличная часть документа) один превышает предел. skip — список со страницами.
 */
function largestList(value: unknown, path: string, skip?: string): NestedList | undefined {
  if (!value || typeof value !== "object") return undefined;
  const own: NestedList | undefined =
    Array.isArray(value) && value.length > 1 && path !== skip
      ? { path, holder: value, chars: renderJson(value).length }
      : undefined;
  const children = Array.isArray(value)
    ? value.map((v, i) => largestList(v, `${path}[${i}]`, skip))
    : Object.entries(value).map(([k, v]) => largestList(v, path ? `${path}.${k}` : k, skip));
  return [own, ...children].reduce<NestedList | undefined>(
    (best, c) => (c && (!best || c.chars > best.chars) ? c : best),
    undefined,
  );
}

type ListCut = { path: string; total: number; returned: number };

/** Срезов за один ответ: защита от квадратичной работы на отчёте из тысяч мелких списков. */
const MAX_LIST_CUTS = 200;

/**
 * Режет вложенные списки, пока ответ не уложится; продолжения у них нет, поэтому каждый срез назван в ответе, и
 * его описание входит в подсчёт. minChars — не трогать мелкие списки: нужно, пока срез верхнего уровня ещё впереди
 * и может решить задачу без потери вложенных строк.
 */
function fitNestedLists(data: Record<string, unknown>, maxChars: number, skip?: string, minChars = 0) {
  const fitted = structuredClone(data);
  const cuts: ListCut[] = [];
  const sizeOf = () => renderJson(fitted).length + (cuts.length ? 2 * renderJson(cuts).length + 200 : 0);
  while (sizeOf() > maxChars && cuts.length < MAX_LIST_CUTS) {
    const list = largestList(fitted, "", skip);
    if (!list || list.chars < minChars) break;
    const before = sizeOf();
    const room = maxChars - (before - list.chars) - 300;
    const kept = fitToChars(list.holder, 0, Math.max(room, 0)).length;
    if (kept >= list.holder.length) break;
    const known = cuts.find((c) => c.path === list.path);
    const removed = list.holder.slice(kept);
    const undo = known ? { ...known } : undefined;
    if (known) known.returned = kept;
    else cuts.push({ path: list.path, total: list.holder.length, returned: kept });
    list.holder.length = kept;
    // Описание среза длиннее убранного: данные потеряны, а ответ вырос. Такой срез не нужен.
    if (sizeOf() >= before) {
      list.holder.push(...removed);
      if (known && undo) known.returned = undo.returned;
      else cuts.pop();
      break;
    }
  }
  return { fitted, cuts };
}

/**
 * Последняя линия защиты размера. Обычная страница сюда не доходит: инструменты считают её с запасом.
 *
 * Инструмент со своими страницами отдаёт offset и returned: срез его страницы обязан сдвинуть nextOffset, иначе
 * срезанные строки клиент не прочитает никогда. Инструмент без своих страниц считает всё заново, и список режется
 * с продолжением по offset (continuable). Без продолжения сначала режутся большие вложенные списки: иначе элемент
 * с таким списком (проверка закрытия месяца) пропал бы целиком. Срезы вложенных списков — в truncatedLists.
 */
export function limitResponseSize(
  result: CallToolResult,
  o: { offset?: number; continuable?: boolean; limit?: number | undefined; page?: PageList | undefined } = {},
): CallToolResult {
  const data = result.structuredContent as Record<string, unknown> | undefined;
  const first = result.content[0];
  if (result.isError || !data || first?.type !== "text") return result;
  const ownPage =
    !o.continuable && typeof data["offset"] === "number" && typeof data["returned"] === "number";
  const offset = ownPage ? (data["offset"] as number) : (o.offset ?? 0);
  // Без настройки срабатывает на норме, а страницы считаются с запасом к ней; настроенный предел строгий.
  const limit = responseCharsLimit();
  const trigger = limit >= MAX_RESPONSE_CHARS ? RESPONSE_CHARS_NORM : limit;
  const lists = Object.keys(data)
    .filter((k) => Array.isArray(data[k]))
    .sort((a, b) => elementChars(data[b]) - elementChars(data[a]));
  // Инструмент считает с начала и отдаёт limit+offset строк. Заполненное окно не значит «всё»: дальше могут быть ещё.
  const windowSize = o.continuable && o.limit !== undefined ? o.limit + offset : undefined;
  const windowList =
    windowSize === undefined
      ? undefined
      : o.page
        ? (data[o.page.list] as unknown[] | undefined)?.length === windowSize
          ? o.page.list
          : undefined
        : lists.find((k) => (data[k] as unknown[]).length === windowSize);
  // Сколько строк всего, если инструмент это знает: тогда заполненное окно не даёт ложного продолжения.
  const knownTotal = o.page?.countIsTotal && typeof data["count"] === "number" ? data["count"] : undefined;
  const moreAfterWindow = !!windowList && (knownTotal === undefined || knownTotal > windowSize!);
  if ((ownPage || offset === 0) && first.text.length <= trigger && !moreAfterWindow) return result;
  // Своя страница — тот список, длину которого инструмент назвал в returned.
  const field = ownPage
    ? (lists.find((k) => (data[k] as unknown[]).length === data["returned"]) ?? undefined)
    : o.page && Array.isArray(data[o.page.list])
      ? o.page.list
      : (windowList ?? lists[0]);
  // С запасом к бюджету страницы ниже: иначе элемент с урезанным списком всё равно не войдёт в неё.
  const before =
    field && !o.continuable && !ownPage
      ? fitNestedLists(data, pageCharsLimit() - 1_500, field, limit / 4)
      : undefined;
  const source = before?.fitted ?? data;
  let structured = source;
  let hint = "";
  let staleNote = false;
  if (field) {
    const list = source[field] as unknown[];
    const { [field]: _list, ...rest } = source;
    const kept = fitToChars(ownPage ? list : list.slice(offset), renderJson(rest).length + 1_000);
    const next = offset + kept.length;
    if (ownPage) {
      const cut = kept.length < list.length;
      staleNote = cut;
      hint = cut ? `Показаны строки ${offset + 1}–${next}. Следующая страница: offset=${next}.` : "";
      structured = {
        ...rest,
        [field]: kept,
        returned: kept.length,
        // count у rows — строки страницы; у остальных это «всего», и срез страницы его не меняет.
        ...(field === "rows" && data["count"] === list.length ? { count: kept.length } : {}),
        ...(cut ? { truncated: true, nextOffset: next } : {}),
      };
    } else {
      const windowFull = field === windowList && kept.length > 0 && knownTotal === undefined;
      const total = knownTotal ?? (windowFull || data["truncated"] === true ? undefined : list.length);
      const more = kept.length > 0 && (next < (knownTotal ?? list.length) || windowFull);
      staleNote = more && !!o.continuable;
      const range = `${offset + 1}–${next}${total === undefined ? "" : ` из ${total}`} списка ${field}`;
      hint = !more
        ? ""
        : o.continuable
          ? `Показаны строки ${range}. Следующая страница: offset=${next}.`
          : `Показаны строки ${range}, больше не помещается. Уточните отбор или период.`;
      structured = {
        ...rest,
        [field]: kept,
        listField: field,
        // Заполненное окно или упор в предел выборки не знают, сколько строк всего.
        ...(total === undefined ? {} : { listTotal: total }),
        ...(field === "rows" && data["count"] === list.length ? { count: kept.length } : {}),
        offset,
        returned: kept.length,
        truncated: more || data["truncated"] === true,
        ...(more && o.continuable ? { nextOffset: next } : {}),
      };
    }
  }
  const after = fitNestedLists(structured, trigger - hint.length - 100, field);
  // Индекс в пути — от начала списка, а не страницы: «rows[120].lines» при offset=120 и есть строка 120.
  const shift = (cut: ListCut): ListCut =>
    field && offset && cut.path.startsWith(`${field}[`)
      ? {
          ...cut,
          path: cut.path.replace(
            /^([^[]+)\[(\d+)\]/,
            (_m, k: string, i: string) => `${k}[${Number(i) + offset}]`,
          ),
        }
      : cut;
  const cuts = [...(before?.cuts ?? []), ...after.cuts.map(shift)];
  if (!field && !cuts.length && renderJson(after.fitted).length <= trigger) return result;
  const cutHint = cuts.length
    ? `Списки показаны не полностью, продолжения нет: ${cuts
        .map((c) => `${c.path} — ${c.returned} из ${c.total}`)
        .join(", ")}. Уточните отбор или период.`
    : "";
  // Текст инструмента называет страницу, которой больше нет, или говорит, что продолжения нет: рядом с новым
  // диапазоном это сбило бы клиента.
  const baseNote = staleNote
    ? String(data["note"] ?? "")
        .replace(/Показан[оы] [^.]*\.(?: [^.]*offset=\d+\.)?\s*/, "")
        .replace(
          /Выборка упёрлась в лимит строк, продолжения нет: показаны не все\. Уточните фильтр или период\.\s*/,
          "",
        )
        .replace(/Результат усечён до \d+ строк\. Уточните фильтр или период\.\s*/, "")
    : data["note"];
  const note = [baseNote, hint, cutHint].filter(Boolean).join(" ");
  const fittedContent = {
    ...after.fitted,
    ...(cuts.length ? { truncated: true, truncatedLists: cuts } : {}),
    ...(note ? { note } : {}),
  };
  const overLimit = renderJson(fittedContent).length > trigger;
  const structuredContent = overLimit
    ? {
        ...fittedContent,
        overLimit: true,
        note: [note, `Ответ больше предела ${trigger} символов: его нечем сократить.`]
          .filter(Boolean)
          .join(" "),
      }
    : fittedContent;
  return {
    ...result,
    content: [{ type: "text", text: renderJson(structuredContent) }, ...result.content.slice(1)],
    structuredContent,
  };
}

/**
 * Простые поля раньше объектов и списков, самый большой — последним: клиент, обрезающий длинный ответ сам,
 * не теряет итоги, предупреждения и признак неполноты. Содержание ответа не меняется.
 */
export function summaryFirst(result: CallToolResult): CallToolResult {
  const data = result.structuredContent;
  if (result.isError || !data || typeof data !== "object") return result;
  const weight = (v: unknown) => (v && typeof v === "object" ? renderJson(v).length : 0);
  const structuredContent = Object.fromEntries(
    Object.entries(data).sort(([, a], [, b]) => weight(a) - weight(b)),
  );
  const text = renderJson(data);
  const content = result.content.map((item) =>
    item.type === "text" && item.text === text ? { ...item, text: renderJson(structuredContent) } : item,
  );
  return { ...result, content, structuredContent };
}

/**
 * Запись при нескольких настроенных базах — только с явной базой, и предпросмотр тоже: иначе
 * документ можно создать не в той организации, а предпросмотр и подтверждение разойдутся по аргументам.
 */
export function requireDatabase(
  ctx: ServerContext,
  cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
): (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult> {
  return async (args, extra) => {
    const names = ctx.databases().map((d) => d.name);
    if (names.length > 1 && typeof args.database !== "string") {
      return fail(
        `Настроено несколько баз (${names.join(", ")}): для записи укажите database. Ничего не записано. ` +
          "Если неясно какая база, спросите пользователя.",
      );
    }
    return cb(args, extra);
  };
}

/** Инструменты write.*, которые ничего не пишут: статус операции читает журнал. */
const WRITE_CHECKED_ELSEWHERE = new Set(["write.operation.status"]);

/**
 * Запись, проверенная для торговой базы (Управление торговлей): инструмент и вид объекта. Общие инструменты —
 * только для документа цен: проведение и пометка любого другого документа УТ по-прежнему отказывают. Положительный
 * признак — цены в схеме устроены как в УТ (priceMechanisms); сбой чтения схемы — отказ.
 */
interface TradeWrite {
  /** Вызов из поддержанного в УТ набора (например, документ цен, а не любой документ). */
  applies: (args: Record<string, unknown>) => boolean;
  /** Схема базы позволяет: у цен — опубликован механизм цен, у контрагента — партнёры. */
  ready: (conn: Connection, args: Record<string, unknown>) => Promise<boolean>;
}
/** Справочники, которые запись в УТ создаёт. */
const TRADE_CATALOGS = new Set([
  "Catalog_Контрагенты",
  PARTNERS,
  "Catalog_БанковскиеСчетаКонтрагентов",
  "Catalog_КонтактныеЛицаПартнеров",
  "Catalog_ДоговорыКонтрагентов",
  "Catalog_СоглашенияСКлиентами",
  "Catalog_СоглашенияСПоставщиками",
  "Catalog_Номенклатура",
]);
const prices = async (conn: Connection) => (await priceMechanisms(conn)).length > 0;
const partners = (conn: Connection) => tradeCounterpartySchema(conn);
const NOMENCLATURE = "Catalog_Номенклатура";
const nomenclature = (conn: Connection) => tradeNomenclatureSchema(conn, NOMENCLATURE);
/** Складские документы УТ: перемещение, излишки, недостачи, пересчёт — устроены как в УТ, а не как в Бухгалтерии. */
const STOCK_SETS = new Set<string>(Object.values(STOCK_DOCS));
const stock = (set: string) => async (conn: Connection) => (await tradeStockSchema(conn, set)) !== undefined;
const stockOf = (conn: Connection, a: Record<string, unknown>) => stock(String(a.entitySet))(conn);
/** Денежные документы УТ: ПКО, РКО, безналичные, заявка — по устройству документа (хозяйственная операция, касса). */
const MONEY_SETS = new Map<string, MoneyKind>(
  (Object.entries(MONEY_DOCS) as Array<[MoneyKind, string]>).map(([k, set]) => [set, k]),
);
const money = (kind: MoneyKind) => async (conn: Connection) =>
  (await tradeMoneySet(conn, kind)) !== undefined;
const moneyOf = (conn: Connection, a: Record<string, unknown>) =>
  money(MONEY_SETS.get(String(a.entitySet))!)(conn);
/** Продажи и закупки УТ: реализация, приобретение, возвраты, счёт клиенту, заказы, счета-фактуры. */
const SALES_SETS = new Map<string, SalesKind>(
  (Object.entries(SALES_DOCS) as Array<[SalesKind, string]>).map(([k, set]) => [set, k]),
);
const sales = (kind: SalesKind) => async (conn: Connection) =>
  (await tradeSalesSet(conn, kind)) !== undefined;
const salesOf = (conn: Connection, a: Record<string, unknown>) =>
  sales(SALES_SETS.get(String(a.entitySet))!)(conn);
const TRADE_WRITE: Record<string, TradeWrite> = {
  "write.pricing.create_price_setting": { applies: () => true, ready: prices },
  "write.document.post_document": {
    applies: (a) =>
      a.entitySet === PRICE_DOC ||
      STOCK_SETS.has(String(a.entitySet)) ||
      MONEY_SETS.has(String(a.entitySet)) ||
      SALES_SETS.has(String(a.entitySet)),
    ready: (conn, a) =>
      a.entitySet === PRICE_DOC
        ? prices(conn)
        : MONEY_SETS.has(String(a.entitySet))
          ? moneyOf(conn, a)
          : SALES_SETS.has(String(a.entitySet))
            ? salesOf(conn, a)
            : stock(String(a.entitySet))(conn),
  },
  // Пометка — и для справочников, которые в УТ создаются (ошибочно созданного контрагента и партнёра убрать).
  "write.entity.mark_for_deletion": {
    applies: (a) =>
      a.entitySet === PRICE_DOC ||
      TRADE_CATALOGS.has(String(a.entitySet)) ||
      STOCK_SETS.has(String(a.entitySet)) ||
      MONEY_SETS.has(String(a.entitySet)) ||
      SALES_SETS.has(String(a.entitySet)),
    ready: (conn, a) =>
      a.entitySet === PRICE_DOC
        ? prices(conn)
        : MONEY_SETS.has(String(a.entitySet))
          ? moneyOf(conn, a)
          : SALES_SETS.has(String(a.entitySet))
            ? salesOf(conn, a)
            : STOCK_SETS.has(String(a.entitySet))
              ? stock(String(a.entitySet))(conn)
              : a.entitySet === NOMENCLATURE
                ? nomenclature(conn)
                : partners(conn),
  },
  "write.counterparty.create_counterparty": { applies: () => true, ready: partners },
  "write.counterparty.create_partner": { applies: () => true, ready: partners },
  "write.counterparty.update_counterparty": { applies: () => true, ready: partners },
  "write.counterparty.create_bank_account": { applies: () => true, ready: partners },
  "write.counterparty.create_contact_person": { applies: () => true, ready: partners },
  "write.catalog.create_contract": { applies: () => true, ready: partners },
  "write.catalog.create_agreement": { applies: () => true, ready: partners },
  "write.catalog.create_nomenclature": { applies: () => true, ready: nomenclature },
  "write.catalog.update_nomenclature": { applies: () => true, ready: nomenclature },
  // Правка, копия и поля складских документов УТ — только этих четырёх документов (белый список полей — в trade-edit).
  ...Object.fromEntries(
    [
      "write.document.update_document_lines",
      "write.document.add_document_line",
      "write.document.remove_document_line",
      "write.document.copy_document",
      "write.entity.update_entity",
    ].map((name) => [
      name,
      {
        applies: (a: Record<string, unknown>) =>
          STOCK_SETS.has(String(a.entitySet)) || SALES_SETS.has(String(a.entitySet)),
        ready: (conn: Connection, a: Record<string, unknown>) =>
          SALES_SETS.has(String(a.entitySet)) ? salesOf(conn, a) : stockOf(conn, a),
      },
    ]),
  ),
  // Копия и поля денежных документов УТ (строк товаров у них нет — инструменты строк отказывают).
  ...Object.fromEntries(
    ["write.document.copy_document", "write.entity.update_entity"].map((name) => [
      name,
      {
        applies: (a: Record<string, unknown>) =>
          STOCK_SETS.has(String(a.entitySet)) ||
          MONEY_SETS.has(String(a.entitySet)) ||
          SALES_SETS.has(String(a.entitySet)),
        ready: (conn: Connection, a: Record<string, unknown>) =>
          MONEY_SETS.has(String(a.entitySet))
            ? moneyOf(conn, a)
            : SALES_SETS.has(String(a.entitySet))
              ? salesOf(conn, a)
              : stockOf(conn, a),
      },
    ]),
  ),
  "write.sales.create_shipment": { applies: () => true, ready: sales("shipment") },
  "write.sales.create_act": { applies: () => true, ready: sales("shipment") },
  // Акт услуг Бухгалтерии и счёт поставщика в УТ — отказ с указанием замены (объясняет сам инструмент).
  "write.sales.create_services_act": { applies: () => true, ready: sales("shipment") },
  "write.purchase.create_supplier_invoice": { applies: () => true, ready: sales("purchase") },
  "write.purchase.create_purchase": { applies: () => true, ready: sales("purchase") },
  "write.warehouse.create_return_from_customer": { applies: () => true, ready: sales("returnFromCustomer") },
  "write.warehouse.create_return_to_supplier": { applies: () => true, ready: sales("returnToSupplier") },
  "write.sales.create_invoice": { applies: () => true, ready: sales("invoice") },
  "write.sales.create_customer_order": { applies: () => true, ready: sales("customerOrder") },
  "write.purchase.create_supplier_order": { applies: () => true, ready: sales("supplierOrder") },
  "write.sales.create_issued_invoice": { applies: () => true, ready: sales("issuedInvoice") },
  "write.purchase.create_received_invoice": { applies: () => true, ready: sales("receivedInvoice") },
  "write.money.create_cash_receipt": { applies: () => true, ready: money("cashIn") },
  "write.money.create_cash_payment": { applies: () => true, ready: money("cashOut") },
  "write.money.create_payment": { applies: () => true, ready: money("bankIn") },
  "write.money.create_bank_writeoff": { applies: () => true, ready: money("bankOut") },
  "write.money.create_payment_request": { applies: () => true, ready: money("request") },
  // Платёжное поручение в УТ — отказ с указанием заявки и списания (сам инструмент объясняет).
  "write.money.create_payout_order": { applies: () => true, ready: money("request") },
  "write.warehouse.create_transfer": { applies: () => true, ready: stock(STOCK_DOCS.transfer) },
  "write.warehouse.create_surplus": { applies: () => true, ready: stock(STOCK_DOCS.surplus) },
  "write.warehouse.create_writeoff": { applies: () => true, ready: stock(STOCK_DOCS.writeoff) },
  "write.warehouse.create_inventory": { applies: () => true, ready: stock(STOCK_DOCS.inventory) },
  // Папки — только номенклатуры: у контрагентов УТ иерархии нет (решение владельца — отказ).
  "write.entity.create_folder": { applies: (a) => a.entitySet === NOMENCLATURE, ready: nomenclature },
  "write.entity.move_to_folder": { applies: (a) => a.entitySet === NOMENCLATURE, ready: nomenclature },
};

/** Что в торговой базе можно записать: подсказка в отказе. */
const TRADE_WRITE_HINT =
  "В Управлении торговлей поддержаны установка цен (write.pricing.create_price_setting, затем " +
  `write.document.post_document и write.entity.mark_for_deletion для ${PRICE_DOC}) и контрагенты с партнёрами ` +
  "(write.counterparty.create_partner, create_counterparty, update_counterparty, create_bank_account, create_contact_person), " +
  "договоры и соглашения (write.catalog.create_contract, create_agreement), номенклатура и её папки " +
  "(create_nomenclature, update_nomenclature, write.entity.create_folder и move_to_folder для Catalog_Номенклатура), " +
  "складские документы (write.warehouse.create_transfer, create_surplus, create_writeoff, create_inventory), " +
  "деньги (write.money.create_cash_receipt, create_cash_payment, create_payment, create_bank_writeoff, " +
  "create_payment_request), продажи и закупки (write.sales.create_customer_order, create_invoice, create_shipment, " +
  "create_act, create_issued_invoice; write.purchase.create_purchase, create_supplier_order, create_received_invoice; " +
  "write.warehouse.create_return_to_supplier, create_return_from_customer); для этих документов — правка строк, " +
  "copy_document, update_entity, post_document и mark_for_deletion. Розница и кассы ККМ — нет.";

/**
 * Запись собирает документы и справочники Бухгалтерии. В базе без регистра бухгалтерии (Управление торговлей, другие
 * небухгалтерские конфигурации) поля и объекты другие, и запись не проверена ни на одной такой базе: отказ до
 * предпросмотра, ничего не отправлено. Признак — только регистр бухгалтерии: по регистрам склада и продаж УТ с
 * неполным составом OData не узнавалась, и запись в неё проходила (ревью).
 * Если база не определена (несколько баз без database), решает следующая проверка — она откажет сама.
 */
function refuseWriteWithoutAccounting(
  ctx: ServerContext,
  name: string,
  cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
): (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult> {
  return async (args, extra) => {
    const several = ctx.databases().length > 1;
    if (typeof args.database === "string" || !several) {
      let conn;
      try {
        conn = ctx.db(typeof args.database === "string" ? args.database : undefined);
      } catch {
        // Неизвестная база — сообщит сам инструмент, как раньше.
        return cb(args, extra);
      }
      let noAccounting: boolean;
      let trade = false;
      let tradeAllowed = false;
      try {
        noAccounting = await withoutAccounting(conn);
        if (noAccounting) trade = await tradeWithoutAccounting(conn);
        const tradeWrite = TRADE_WRITE[name];
        if (noAccounting && tradeWrite?.applies(args)) tradeAllowed = await tradeWrite.ready(conn, args);
      } catch (e) {
        // Защита закрывается, а не открывается: при сбое схемы инструмент сам перечитал бы её и записал (ревью:
        // первый $metadata 503 — и PATCH уходил в базу УТ).
        return fail(
          `Не удалось проверить конфигурацию базы перед записью (${e instanceof Error ? e.message : String(e)}). ` +
            "Ничего не записано; повторите позже.",
          "refused",
        );
      }
      if (noAccounting && tradeAllowed) return cb(args, extra);
      if (noAccounting) {
        // Регистр из настройки учёта называется как есть: совет «опубликуйте Хозрасчётный» базе со своим регистром
        // ведёт не туда (ревью).
        const configured = conn.cfg.accounting?.register;
        const register = configured ? `«${configured}» (из настройки учёта)` : "«Хозрасчётный»";
        return fail(
          `Этот инструмент записи поддерживается только для Бухгалтерии: в этой базе не опубликован регистр бухгалтерии ${register}` +
            `${trade ? " (похоже, это Управление торговлей)" : ""}. Документы и справочники в ней устроены иначе, ` +
            "запись на такой базе не проверена. Ничего не записано. Если это Бухгалтерия, КА или ERP, опубликуйте " +
            `${register} в составе OData.${trade ? ` ${TRADE_WRITE_HINT}` : ""}`,
          "refused",
        );
      }
    }
    return cb(args, extra);
  };
}

/** Создаёт MCP-сервер и регистрирует все инструменты (чтение + гейтованная запись). */
export function createServer(ctx: ServerContext): McpServer {
  // Какие инструменты видит клиент. Это про объём контекста модели, запись защищают READ_ONLY и WRITABLE.
  const { toolsProfile = "all", toolsAllow } = ctx.behavior;
  const allowed = toolsAllow && new Set(toolsAllow);
  const listed = (name: string) => toolListed(ctx.behavior, name);
  const known = new Set<string>();
  const server = new McpServer(
    { name: "1c-odata-mcp", version: readVersion() },
    {
      instructions: allowed
        ? `${INSTRUCTIONS} ${LISTED_TOOLS_NOTE}`
        : toolsProfile === "read"
          ? `${INSTRUCTIONS} Инструменты записи на этом сервере не подключены (TOOLS_PROFILE=read).`
          : INSTRUCTIONS,
      capabilities: { extensions: { [APPS_EXTENSION_ID]: {} } },
    },
  );

  // Оборачиваем registerTool, чтобы каждому инструменту проставить аннотации по имени.
  const original = server.registerTool.bind(server) as unknown as (
    name: string,
    config: Record<string, unknown>,
    cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
  ) => unknown;
  const registerTool = (
    name: string,
    config: Record<string, unknown>,
    handler: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
  ) => {
    known.add(name);
    if (!listed(name)) return undefined;
    const cb =
      WRITE_CHECKED_ELSEWHERE.has(name) || !name.startsWith("write.")
        ? handler
        : refuseWriteWithoutAccounting(ctx, name, handler);
    const appMeta = appMetaFor(name);
    const annotatedConfig = {
      ...config,
      annotations: {
        ...annotationsFor(name),
        ...(config.annotations as Record<string, unknown> | undefined),
      },
      ...(appMeta ? { _meta: { ...(config._meta as Record<string, unknown> | undefined), ...appMeta } } : {}),
    };
    const isCreateTool = name.startsWith("write.") && config.outputSchema === createResultSchema;
    if (name.startsWith("read.") && name !== "read.system.list_databases") {
      const input = config.inputSchema;
      const shape = input instanceof z.ZodObject ? (input.shape as Record<string, unknown>) : {};
      // Свои страницы есть — обёртка только страхует. У инструмента со списком по limit появляется offset.
      const continuable = "limit" in shape && !("offset" in shape);
      const offsetSchema = continuable
        ? {
            inputSchema: (input as z.ZodObject<z.ZodRawShape>).extend({
              offset: offsetField,
            }),
          }
        : {};
      return original(name, { ...annotatedConfig, ...offsetSchema }, async (args, extra) => {
        const offset = continuable && typeof args.offset === "number" ? args.offset : 0;
        // Инструмент считает с начала: чтобы дойти до этой страницы, ему нужно offset + limit строк.
        const toolArgs =
          offset > 0 && typeof args.limit === "number" ? { ...args, limit: args.limit + offset } : args;
        const signal = (extra as { mcpReq?: { signal?: AbortSignal } } | undefined)?.mcpReq?.signal;
        const { toolDeadlineMs: deadlineMs = 0, maxResponseChars } = ctx.behavior;
        return withToolCall({ signal, deadlineMs, maxResponseChars, pageStart: offset }, async () => {
          const result = await cb(toolArgs, extra);
          // Прерванный вызов не отдаёт результат: часть инструментов переживает сбой отдельного запроса
          // и показала бы неполные данные как обычный ответ.
          const aborted = callAbortError();
          if (aborted) return fail(`[${aborted.kind}] ${aborted.message}`, aborted.kind);
          return summaryFirst(
            limitResponseSize(withDatabaseInfo(ctx, args, result), {
              offset,
              continuable,
              limit: continuable && typeof args.limit === "number" ? args.limit : undefined,
              page: PAGE_LISTS[name],
            }),
          );
        });
      });
    }
    if (!isCreateTool && !JOURNALED_UPDATE_TOOLS.has(name)) {
      return original(name, annotatedConfig, name.startsWith("write.") ? requireDatabase(ctx, cb) : cb);
    }

    // v2: inputSchema — z.object(...); добавляем operationId через extend (разворот zod-объекта
    // как простого объекта даёт не схему, а его внутренние поля).
    const operationIdField = z
      .string()
      .uuid()
      .optional()
      .describe("UUID из предпросмотра. Передайте его при confirm=true, чтобы повтор не создал дубликат.");
    const base = config.inputSchema;
    const inputSchema =
      base instanceof z.ZodObject
        ? base.extend({ operationId: operationIdField })
        : { ...(base as Record<string, unknown>), operationId: operationIdField };
    const wrapped = async (args: Record<string, unknown>, extra: unknown): Promise<CallToolResult> => {
      if (name === "write.document.post_document" && args.expectedDataVersion === undefined)
        return cb(args, extra);
      const operationId = typeof args.operationId === "string" ? args.operationId : undefined;
      if (args.confirm === true && !operationId) {
        return fail("Сначала выполните dry-run. При confirm=true передайте operationId из его результата.");
      }
      const token = operationId ?? randomUUID();
      const requestHash = fingerprintWriteInput(name, args as Record<string, unknown>);
      // Повтор уже подтверждённой операции: ответ из журнала — до того, как инструмент заново
      // прочитает 1С (иначе правка строк применилась бы к уже изменённому документу).
      if (args.confirm === true && operationId) {
        let replay: CallToolResult | undefined;
        const checked = await guard(name, async () => {
          const conn = ctx.db(typeof args.database === "string" ? args.database : undefined);
          const settled = await conn.client.operationSettled(operationId, requestHash);
          if (settled) {
            replay = ok({
              ...(isCreateTool ? { created: true } : { updated: true }),
              ...(name === "write.document.post_document"
                ? {
                    done: true,
                    action: args.post ? "Post" : "Unpost",
                    verified: false,
                    verificationScope: "journal_receipt",
                    readbackFetched: false,
                    accountingMovementsVerified: false,
                    readbackNote:
                      "Журнал хранит только квитанцию; полный документ и проводки повторно не читались.",
                  }
                : {}),
              // Правка банковского документа отдаёт документ из контрольного чтения (#50); повтор — нет.
              ...(BANK_DOCUMENT_TOOLS.has(name)
                ? {
                    verified: false,
                    verificationScope: "journal_receipt",
                    readbackFetched: false,
                    readbackNote: "Журнал хранит только квитанцию; документ повторно не читался.",
                  }
                : {}),
              replayed: true,
              note: "Операция с этим operationId уже выполнена; повторный запрос в 1С не отправлялся.",
              database: conn.cfg.name,
              entitySet: settled.entitySet,
              ref: settled.result["Ref_Key"],
              ...(settled.result["RecordPath"] ? { recordPath: settled.result["RecordPath"] } : {}),
              ...(settled.result["Code"] ? { code: settled.result["Code"] } : {}),
              operationId,
            });
          }
          return ok({});
        });
        if (checked.isError) return checked;
        if (replay) return replay;
        // Срок предпросмотра (#67): интеграция может подтвердить давний предпросмотр, а документ получил бы дату и
        // данные того дня. Выполненные операции выше уже отданы из журнала; сюда доходит только неподтверждённая.
        const maxAge = ctx.behavior.previewMaxAgeMs ?? 0;
        if (maxAge > 0) {
          let stale: CallToolResult | undefined;
          const aged = await guard(name, async () => {
            const conn = ctx.db(typeof args.database === "string" ? args.database : undefined);
            const entry = await conn.client.operationEntry(operationId);
            // Нечитаемая или будущая дата предпросмотра (журнал правили, часы ушли назад) — как устаревшая.
            const created = entry?.state === "prepared" ? Date.parse(entry.createdAt) : Date.now();
            const age =
              Number.isFinite(created) && created <= Date.now() + 60_000 ? Date.now() - created : Infinity;
            if (age > maxAge)
              stale = fail(
                `${Number.isFinite(age) ? `Предпросмотр сделан ${Math.floor(age / 60_000)} мин назад` : "Дата предпросмотра в журнале нечитаема или в будущем"} — дольше срока ODATA_WRITE_PREVIEW_MAX_AGE ` +
                  `(${Math.round(maxAge / 60_000)} мин): данные в 1С могли измениться. Сделайте новый предпросмотр и ` +
                  "подтвердите его operationId. Ничего не записано.",
              );
            return ok({});
          });
          if (aged.isError) return aged;
          if (stale) return stale;
        }
      }
      return withWriteOperation(token, requestHash, async () => {
        const result = await cb(args, extra);
        if (result?.isError || !result?.structuredContent || typeof result.structuredContent !== "object")
          return result;
        const structuredContent = { ...result.structuredContent, operationId: token };
        const content = result.content.map((item) =>
          item.type === "text" ? { ...item, text: renderJson(structuredContent) } : item,
        );
        return { ...result, content, structuredContent };
      });
    };
    return original(name, { ...annotatedConfig, inputSchema }, requireDatabase(ctx, wrapped));
  };
  Reflect.set(server, "registerTool", (name: string, config: unknown, cb: unknown) =>
    registerTool(
      name,
      config as Record<string, unknown>,
      cb as (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
    ),
  );

  registerMetaTools(server, ctx);
  registerCounterpartyTools(server, ctx);
  registerDocumentTools(server, ctx);
  registerRegisterTools(server, ctx);
  registerStockTools(server, ctx);
  registerItemSalesTools(server, ctx);
  registerPartnerSettlementTools(server, ctx);
  registerWarehouseTools(server, ctx);
  registerItemTools(server, ctx);
  registerUserTools(server, ctx);
  registerPricingTools(server, ctx);
  registerDebtTools(server, ctx);
  registerMonthCloseTools(server, ctx);
  registerCardTools(server, ctx);
  registerCashflowTools(server, ctx);
  registerSalesTools(server, ctx);
  registerOrganizationTools(server, ctx);
  registerAuditTools(server, ctx);
  registerFreshnessTools(server, ctx);
  registerWriteTools(server, ctx);
  registerExistingBankSettlementTools(server, ctx);
  registerSettlementSettingsTools(server, ctx);
  registerCounterpartyGroupTools(server, ctx);
  registerPrompts(server, listed);
  registerApps(server);

  // Опечатка в списке молча оставила бы клиента без инструмента: отказ при старте с точным именем.
  const unknown = (toolsAllow ?? []).filter((name) => !known.has(name));
  if (unknown.length) {
    throw new Error(
      `TOOLS_ALLOW: таких инструментов нет: ${unknown.join(", ")}. Имена — как в tools/list, например read.analytics.get_debtors.`,
    );
  }

  return server;
}
