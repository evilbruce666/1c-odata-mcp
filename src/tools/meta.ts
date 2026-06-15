import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { ok, fail, guard } from "./_shared.js";
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
    "health_check",
    {
      title: "Проверка соединения",
      description:
        "Проверяет доступность OData-сервиса 1С и корректность авторизации. " +
        "Возвращает версию OData и число опубликованных объектов. " +
        "Вызывайте первым, если другие инструменты выдают ошибки.",
      inputSchema: {},
    },
    () =>
      guard("health_check", async () => {
        const meta = await ctx.getMetadata();
        return ok({
          status: "ok",
          odataVersion: meta.odataVersion,
          entityCount: meta.entities.size,
          baseUrl: ctx.cfg.ODATA_BASE_URL,
          readOnly: ctx.cfg.READ_ONLY,
        });
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
    ({ class: cls, search }) =>
      guard("list_entities", async () => {
        const meta = await ctx.getMetadata();
        const needle = search?.toLowerCase();
        const grouped: Record<string, Array<{ entitySet: string; name: string }>> = {};

        for (const e of meta.entities.values()) {
          if (cls && e.class !== cls) continue;
          if (needle && !e.entitySet.toLowerCase().includes(needle)) continue;
          const label = CLASS_LABEL[e.class];
          (grouped[label] ??= []).push({ entitySet: e.entitySet, name: e.shortName });
        }

        return ok({ odataVersion: meta.odataVersion, groups: grouped });
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
        entitySet: z.string().describe("Техническое имя объекта, напр. Catalog_Контрагенты"),
      },
    },
    ({ entitySet }) =>
      guard("describe_entity", async () => {
        const meta = await ctx.getMetadata();
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
