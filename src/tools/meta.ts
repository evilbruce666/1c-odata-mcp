import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { ok, fail, guard, databaseField, offsetField, pageOf, snapshotField, toolListed } from "./_shared.js";
import { listOrganizations } from "../odata/orgs.js";
import type { EntityClass } from "../types/odata.js";
import {
  listDatabasesResultSchema,
  healthCheckResultSchema,
  listOrganizationsResultSchema,
  listEntitiesResultSchema,
  describeEntityResultSchema,
} from "../schemas/output.js";

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
    "read.system.list_databases",
    {
      title: "Список баз 1С",
      description:
        "Показывает настроенные базы 1С, к которым подключён сервер, и какая из них основная. " +
        "Имя базы передаётся в параметр database остальных инструментов. " +
        "Вызывайте, если нужно понять, какие базы доступны или сравнить данные нескольких баз.",
      inputSchema: z.object({}),
      outputSchema: listDatabasesResultSchema,
    },
    () =>
      guard("read.system.list_databases", async () =>
        ok({ default: ctx.defaultName, databases: ctx.databases() }),
      ),
  );

  server.registerTool(
    "read.system.health_check",
    {
      title: "Проверка соединения",
      description:
        "Проверяет доступность OData-сервиса 1С и корректность авторизации для выбранной базы. " +
        "Перечитывает $metadata и обновляет кеш схемы; возвращает версию OData и число видимых объектов. " +
        "Вызывайте первым, если другие инструменты выдают ошибки.",
      inputSchema: z.object({ database: databaseField }),
      outputSchema: healthCheckResultSchema,
    },
    ({ database }) =>
      guard("read.system.health_check", async () => {
        const conn = ctx.db(database);
        const meta = await conn.client.refreshMetadata();
        const journal = await conn.client.journalSummary();
        return ok({
          status: "ok",
          database: conn.cfg.name,
          label: conn.cfg.label,
          odataVersion: meta.odataVersion,
          entityCount: meta.entities.size,
          // Адрес базы нужен при настройке на своём компьютере; держателю токена и модели в сети он ни к чему.
          ...(conn.behavior.networkTransport ? {} : { baseUrl: conn.cfg.baseUrl }),
          readOnly: conn.behavior.readOnly,
          regulatedCurrency: {
            code: conn.behavior.regulatedCurrencyCode,
            source: conn.regulatedCurrencySource,
          },
          ...(journal
            ? {
                writeJournal: {
                  dir: journal.dir,
                  writable: journal.writable,
                  uncertainOperations: journal.uncertain.length,
                  ...(journal.uncertain.length ? { uncertainIds: journal.uncertain.slice(0, 20) } : {}),
                  ...(journal.error ? { error: journal.error } : {}),
                  note: !journal.writable
                    ? "Каталог журнала недоступен для записи — подтверждение создания и правки строк не пройдёт. Проверьте ODATA_WRITE_JOURNAL_DIR (в Docker — том)."
                    : journal.uncertain.length
                      ? toolListed(conn.behavior, "write.operation.status")
                        ? "Есть операции с неизвестным исходом: проверьте их write.operation.status — повтор по ним заблокирован до сверки."
                        : "Есть операции с неизвестным исходом, повтор по ним заблокирован до сверки. Сверить их нечем: write.operation.status на этом сервере скрыт настройкой TOOLS_PROFILE или TOOLS_ALLOW."
                      : "Журнал в порядке.",
                },
              }
            : {}),
        });
      }),
  );

  server.registerTool(
    "read.system.list_organizations",
    {
      title: "Организации базы",
      description:
        "Список организаций (юрлиц) внутри выбранной базы 1С. " +
        "Название организации передаётся в параметр organization аналитических инструментов " +
        "(get_sales, get_debtors, get_cashflow, get_inventory и др.), чтобы получить данные по одному юрлицу.",
      inputSchema: z.object({ database: databaseField }),
      outputSchema: listOrganizationsResultSchema,
    },
    ({ database }) =>
      guard("read.system.list_organizations", async () => {
        const orgs = await listOrganizations(ctx.db(database));
        return ok({ count: orgs.length, organizations: orgs });
      }),
  );

  server.registerTool(
    "read.schema.list_entities",
    {
      title: "Карта объектов базы",
      description:
        "Возвращает карту бизнес-объектов 1С, сгруппированную по классам: " +
        "справочники, документы, регистры, перечисления и т.д. " +
        "Используйте, чтобы понять, какие данные доступны в базе. " +
        "В базе тысячи объектов: укажите class или search, countsByClass показывает, сколько их; ответ постраничный (offset, nextOffset). Можно отфильтровать по классу и/или подстроке имени.",
      inputSchema: z.object({
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
        offset: offsetField,
        snapshot: snapshotField,
      }),
      outputSchema: listEntitiesResultSchema,
    },
    ({ database, class: cls, search, offset, snapshot }) =>
      guard("read.schema.list_entities", async () => {
        const conn = ctx.db(database);
        const meta = await conn.getMetadata();
        const needle = search?.toLowerCase();
        const matched: Array<{ label: string; entitySet: string; name: string }> = [];
        const countsByClass: Record<string, number> = {};
        for (const e of meta.entities.values()) {
          if (cls && e.class !== cls) continue;
          if (needle && !e.entitySet.toLowerCase().includes(needle)) continue;
          const label = CLASS_LABEL[e.class];
          countsByClass[label] = (countsByClass[label] ?? 0) + 1;
          matched.push({ label, entitySet: e.entitySet, name: e.shortName });
        }
        // Страницы — по классам подряд, чтобы группы в ответе не перемешивались.
        const ordered = Object.keys(countsByClass).flatMap((l) => matched.filter((m) => m.label === l));
        const head = {
          database: conn.cfg.name,
          odataVersion: meta.odataVersion,
          total: matched.length,
          countsByClass,
        };
        const { items, page, note } = pageOf(ordered, {
          offset,
          head,
          unit: "объекты",
          snapshot,
          tracksChanges: true,
        });
        const groups: Record<string, Array<{ entitySet: string; name: string }>> = {};
        for (const { label, ...e } of items) (groups[label] ??= []).push(e);
        return ok({
          ...head,
          ...page,
          groups,
          ...(note ? { note: `${note} Короче: укажите class или search.` } : {}),
        });
      }),
  );

  server.registerTool(
    "read.schema.describe_entity",
    {
      title: "Описание объекта",
      description:
        "Показывает поля (с типами) и связи конкретного объекта 1С из $metadata. " +
        "Принимает техническое имя EntitySet, напр. 'Catalog_Контрагенты' или " +
        "'Document_РеализацияТоваровУслуг'. Используйте перед поиском, чтобы узнать доступные поля.",
      inputSchema: z.object({
        database: databaseField,
        entitySet: z.string().describe("Техническое имя объекта, напр. Catalog_Контрагенты"),
      }),
      outputSchema: describeEntityResultSchema,
    },
    ({ database, entitySet }) =>
      guard("read.schema.describe_entity", async () => {
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
