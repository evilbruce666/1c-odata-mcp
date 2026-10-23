import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";

/**
 * Готовые сценарии (MCP prompts): клиент показывает их пользователю как команды, а текст
 * говорит модели, какие инструменты вызвать и в каком виде отдать ответ. Один источник для
 * сервера и manifest.json (npm run manifest): в тексте — шаблоны ${arguments.<имя>}, как в
 * формате MCPB. Сценарии только читают: ничего не создают и не меняют в 1С.
 */

export interface PromptArg {
  name: string;
  description: string;
  required?: boolean;
  /** Что подставить в текст, если необязательный аргумент не задан. */
  ifEmpty?: string;
}

export interface PromptDef {
  name: string;
  title: string;
  description: string;
  arguments: PromptArg[];
  text: string;
}

const DB_ARGS: PromptArg[] = [
  {
    name: "database",
    description: "База 1С из list_databases (без указания — основная)",
    ifEmpty: "основная",
  },
  {
    name: "organization",
    description: "Организация (без указания — все организации базы)",
    ifEmpty: "все организации базы",
  },
];

const SCOPE = "База: ${arguments.database}. Организация: ${arguments.organization}.";

export const PROMPTS: PromptDef[] = [
  {
    name: "morning_brief",
    title: "Утренняя сводка руководителя",
    description:
      "Деньги на счетах, кто просрочил оплату, неоплаченные счета и что пора платить — одной сводкой",
    arguments: DB_ARGS,
    text: [
      `Подготовь утреннюю сводку руководителя по данным 1С. ${SCOPE}`,
      "",
      "1. read.analytics.get_cash_balance — сколько денег доступно сейчас (available); спецсчета и деньги в пути — отдельно; отрицательные остатки назови.",
      "2. read.analytics.get_debtors с overdueOnly=true и limit=10 — кто просрочил оплату, сколько и на сколько дней. Если в ответе aging.fallback = byAge (в договорах не задан срок оплаты), это долги старше 60 дней — скажи об этом одной строкой. Возраст и просрочка считаются до вычета авансов: итог по должнику — amount.",
      "3. read.analytics.get_unpaid_invoices — неоплаченные счета покупателям за последние 90 дней; выдели самые старые и те, где customerBalance ≤ 0 (их стоит уточнить у бухгалтера).",
      "4. read.analytics.get_creditors с aging=true и limit=10 — кому мы должны и что пора платить.",
      "",
      "Формат: четыре коротких блока — «Деньги», «Нам должны», «Счета без оплаты», «Мы должны»; в каждом итог и до пяти строк. В конце — одно-три действия на сегодня. Суммы в рублях с разделителями разрядов. Ничего не создавай и не изменяй в 1С.",
    ].join("\n"),
  },
  {
    name: "month_close_review",
    title: "Разбор закрытия месяца",
    description:
      "Проверка месяца и разбор каждой найденной проблемы до конкретных документов — чек-лист для бухгалтера",
    arguments: [{ name: "month", description: "Месяц в формате YYYY-MM", required: true }, ...DB_ARGS],
    text: [
      `Проверь закрытие месяца \${arguments.month} в 1С. ${SCOPE}`,
      "",
      "1. read.accounting.check_month_close за этот месяц.",
      "2. По каждой проверке со статусом error или warning найди причину до документа. Остаток на закрываемом счёте или красное сальдо — read.accounting.get_account_card по этому счёту за месяц (ссылка из analyticsRefs: при analyticsTypes Catalog_Контрагенты — counterpartyRef, иначе dimensionRef); для быстрого обзора субсчетов — read.accounting.get_account_turnover. Если остаток сложился раньше месяца — read.accounting.get_balance_origin по этому счёту и аналитике покажет, с какой даты и из каких документов он висит. Проверка lateDocuments уже называет документы, созданные после закрытия, — их суммы в закрытие не попали. Непроведённый документ — read.document.get_document.",
      "3. Итог — чек-лист для бухгалтера: что исправить, в каком документе и в каком порядке (сначала незакрытые предыдущие месяцы, затем документы, затем перезакрыть месяц).",
      "",
      "Проверки со статусом ok не пересказывай. Ничего не проводи и не изменяй — только предложи действия.",
    ].join("\n"),
  },
  {
    name: "reconciliation_act",
    title: "Акт сверки с контрагентом",
    description: "Данные для акта сверки взаиморасчётов с контрагентом за период — по документам и договорам",
    arguments: [
      { name: "counterparty", description: "Название или ИНН контрагента", required: true },
      { name: "from", description: "Начало периода, YYYY-MM-DD", required: true },
      { name: "to", description: "Конец периода, YYYY-MM-DD", required: true },
      ...DB_ARGS,
    ],
    text: [
      "Подготовь данные для акта сверки с контрагентом «${arguments.counterparty}» за период с ${arguments.from} по ${arguments.to}. " +
        SCOPE,
      "",
      "1. read.counterparty.find_counterparty — найди контрагента; если подходит несколько — спроси, какой.",
      "2. read.counterparty.get_reconciliation за этот период.",
      "3. Покажи как акт сверки: сальдо на начало, таблица (дата, документ, номер, дебет, кредит), обороты, сальдо на конец «в пользу организации / контрагента». Если договоров несколько — итоги по каждому. Если consistent=false — предупреди и не выдавай итог за окончательный.",
      "",
      "Напомни одной строкой, что печатная форма акта формируется в 1С. Ничего не создавай и не изменяй в 1С.",
    ].join("\n"),
  },
];

/** Подставляет аргументы в шаблон; не заданный необязательный аргумент — его ifEmpty. */
export function renderPrompt(def: PromptDef, args: Record<string, string | undefined>): string {
  return def.text.replace(/\$\{arguments\.([A-Za-z_]+)\}/g, (_, name: string) => {
    const v = args[name]?.trim();
    return v || def.arguments.find((a) => a.name === name)?.ifEmpty || "не указано";
  });
}

export function registerPrompts(server: McpServer): void {
  for (const def of PROMPTS) {
    const argsSchema = Object.fromEntries(
      def.arguments.map((a) => [
        a.name,
        a.required ? z.string().describe(a.description) : z.string().optional().describe(a.description),
      ]),
    );
    server.registerPrompt(
      def.name,
      { title: def.title, description: def.description, argsSchema: z.object(argsSchema) },
      (args) => ({
        messages: [
          {
            role: "user",
            content: { type: "text", text: renderPrompt(def, args as Record<string, string | undefined>) },
          },
        ],
      }),
    );
  }
}
