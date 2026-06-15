import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { ok, fail, guard, databaseField } from "./_shared.js";
import { listOrganizations } from "../odata/orgs.js";
import type { EntityClass } from "../types/odata.js";

const CLASS_LABEL: Record<EntityClass, string> = {
  catalog: "Справочники",
  document: "Документы",
  documentJournal: "Журналы документов",
  accumulationRegister: "Регистры накопления",
  accountingRegister: "Регистры бухгалтерии",
  informationRegister: "Регистры сведений",
  calculationRegister: "Регистры расчёта",
  enum: "Перечисления",
  chartOfAccounts: "Планы счетов",
  chartOfCharacteristicTypes: "Планы видов характеристик",
  constant: "Константы",
  businessProcess: "Бизнес-процессы",
  task: "Задачи",
  exchangePlan: "Планы обмена",
  other: "Прочее",
};

export function registerMetaTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "list_databases",
    {
      title: "Список баз 1С",
      description:
        "Показывает настроенные базы 1С, к которым подключён сервер, и какая из них основная. " +
        "Имя базы передаётся в параметр database остальных инструментов. " +
        "Вызывайте, если нужно понять, какие базы доступны или сравнить данные нескольких баз.",
      inputSchema: {},
    },
    () =>
      guard("list_databases", async () => ok({ default: ctx.defaultName, databases: ctx.databases() })),
  );

  server.registerTool(
    "health_check",
    {
      title: "Проверка соединения",
      description:
        "Проверяет доступность OData-сервиса 1С и корректность авторизации для выбранной базы. " +
        "Возвращает версию OData и число опубликованных объектов. " +
        "Вызывайте первым, если другие инструменты выдают ошибки.",
      inputSchema: { database: databaseField },
    },
    ({ database }) =>
      guard("health_check", async () => {
        const conn = ctx.db(database);
        const meta = await conn.getMetadata();
        return ok({
          status: "ok",
          database: conn.cfg.name,
          label: conn.cfg.label,
          odataVersion: meta.odataVersion,
          entityCount: meta.entities.size,
          baseUrl: conn.cfg.baseUrl,
          readOnly: conn.behavior.readOnly,
        });
      }),
  );

  server.registerTool(
    "list_organizations",
    {
      title: "Организации базы",
      description:
        "Список организаций (юрлиц) внутри выбранной базы 1С. " +
        "Название организации передаётся в параметр organization аналитических инструментов " +
        "(get_sales, get_debtors, get_cashflow, get_inventory и др.), чтобы получить данные по одному юрлицу.",
      inputSchema: { database: databaseField },
    },
    ({ database }) =>
      guard("list_organizations", async () => {
        const orgs = await listOrganizations(ctx.db(database));
        return ok({ count: orgs.length, organizations: orgs });
      }),
  );

  server.registerTool(
    "list_entities",
    {
      title: "Карта объектов базы",
      description:
        "Возвращает карту бизнес-объектов 1С, сгруппированную по классам: " +
        "справочники, документы, регистры, перечисления и т.д. " +
        "Используйте, чтобы понять, какие данные доступны в базе. " +
        "Можно отфильтровать по классу и/или подстроке имени.",
      inputSchema: {
        database: databaseField,
        class: z
          .enum([
            "catalog",
            "document",
            "accumulationRegister",
            "accountingRegister",
            "informationRegister",
            "enum",
            "chartOfAccounts",
          ])
          .optional()
          .describe("Ограничить одним классом объектов"),
        search: z.string().optional().describe("Подстрока в имени объекта (без учёта регистра)"),
      },
    },
    ({ database, class: cls, search }) =>
      guard("list_entities", async () => {
        const meta = await ctx.db(database).getMetadata();
        const needle = search?.toLowerCase();
        const grouped: Record<string, Array<{ entitySet: string; name: string }>> = {};

        for (const e of meta.entities.values()) {
          if (cls && e.class !== cls) continue;
          if (needle && !e.entitySet.toLowerCase().includes(needle)) continue;
          const label = CLASS_LABEL[e.class];
          (grouped[label] ??= []).push({ entitySet: e.entitySet, name: e.shortName });
        }

        return ok({ database: ctx.db(database).cfg.name, odataVersion: meta.odataVersion, groups: grouped });
      }),
  );

  server.registerTool(
    "describe_entity",
    {
      title: "Описание объекта",
      description:
        "Показывает поля (с типами) и связи конкретного объекта 1С из $metadata. " +
        "Принимает техническое имя EntitySet, напр. 'Catalog_Контрагенты' или " +
        "'Document_РеализацияТоваровУслуг'. Используйте перед поиском, чтобы узнать доступные поля.",
      inputSchema: {
        database: databaseField,
        entitySet: z.string().describe("Техническое имя объекта, напр. Catalog_Контрагенты"),
      },
    },
    ({ database, entitySet }) =>
      guard("describe_entity", async () => {
        const meta = await ctx.db(database).getMetadata();
        const e = meta.entities.get(entitySet);
        if (!e) {
          return fail(
            `Объект '${entitySet}' не найден в базе. ` +
              `Проверьте имя через list_entities и что он включён в «Состав OData».`,
          );
        }
        return ok({
          entitySet: e.entitySet,
          class: CLASS_LABEL[e.class],
          keys: e.keys,
          fields: e.properties.map((p) => ({ name: p.name, type: p.type, nullable: p.nullable })),
          relations: e.navigations.map((n) => n.name),
        });
      }),
  );
}
