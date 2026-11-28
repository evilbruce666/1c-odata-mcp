import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { InputError } from "../errors.js";
import { resolveOrganization } from "../odata/orgs.js";
import { FILTER_OPERATORS, readRegisterRows } from "../odata/register-rows.js";
import { readRegisterResultSchema } from "../schemas/output.js";
import {
  databaseField,
  fitToChars,
  guard,
  offsetField,
  ok,
  organizationField,
  pageOf,
  renderJson,
  snapshotField,
  snapshotFields,
} from "./_shared.js";
import { compactDocument } from "./document-names.js";

const NOT_TOTALS =
  "Это строки регистра, а не итоги: в движениях бывают сторно и корректировки, не складывайте их самостоятельно.";

const day = z.string().regex(/^\d{4}-\d{2}-\d{2}$/, "Дата в формате YYYY-MM-DD");

export function registerReadRegisterTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.schema.read_register",
    {
      title: "Строки регистра",
      description:
        "Строки любого опубликованного регистра 1С (сведений, накопления, бухгалтерии) — только когда нет " +
        "профильного инструмента: долги, остатки, деньги, проводки, налоги и цены берите ими, они учитывают " +
        "сторно и особенности учёта. Сначала list_entities (class=informationRegister и т. п.) и describe_entity: " +
        "select обязателен, имена полей — из схемы. Возвращает строки, а не итоги: сами их не складывайте. У " +
        "регистра с регистратором читается таблица записей «…_RecordType» (её имя — в register). Даты в filters — " +
        "«2026-03-31» (весь день) или «2026-03-31T00:00:00»; ссылка составного поля — с type. Двоичные данные " +
        "и поля с «пароль», «секрет», «токен» в имени не читаются. Права — у пользователя 1С. Если truncated без nextOffset — прочитана часть: сузьте " +
        "отбор.",
      inputSchema: z.object({
        database: databaseField,
        register: z.string().describe("Имя регистра, напр. InformationRegister_КурсыВалют"),
        select: z.array(z.string()).min(1).max(30).describe("Поля, которые вернуть (из describe_entity)"),
        filters: z
          .array(
            z.object({
              field: z.string().describe("Имя поля"),
              op: z.enum(FILTER_OPERATORS).describe("Сравнение; contains — подстрока в строковом поле"),
              value: z.union([z.string(), z.number(), z.boolean()]).describe("Значение"),
              type: z
                .string()
                .optional()
                .describe("Для ссылки в составном поле — её тип, напр. Catalog_Контрагенты"),
            }),
          )
          .max(10)
          .optional()
          .describe("Отборы, все одновременно"),
        organization: organizationField,
        from: day.optional().describe("С даты по полю Period"),
        to: day.optional().describe("По дату по полю Period"),
        orderBy: z
          .string()
          .optional()
          .describe("Поле сортировки; без него порядок — по значениям выбранных полей"),
        descending: z.boolean().optional().describe("true — по убыванию orderBy (напр. последние записи)"),
        compact: z
          .boolean()
          .optional()
          .describe("true — без пустых значений, с названиями рядом со ссылками, где их можно определить"),
        limit: z.number().int().positive().max(500).default(100).describe("Сколько строк вернуть"),
        offset: offsetField,
        snapshot: snapshotField,
      }),
      outputSchema: readRegisterResultSchema,
    },
    ({ database, organization, compact, limit, offset, snapshot, ...request }) =>
      guard("read.schema.read_register", async () => {
        const conn = ctx.db(database);
        const orgRef = organization ? (await resolveOrganization(conn, organization)).ref : undefined;
        const found = await readRegisterRows(conn, { ...request, orgRef });
        const head = {
          database: conn.cfg.name,
          register: found.register,
          // Прочитано строк под отбором; при capped их в 1С больше.
          matched: found.rows.length,
        };
        // Метка — по всей выборке: изменилась любая строка, а не только их число.
        const changes = snapshotFields({ head, rows: found.rows }, snapshot);
        if (found.capped && offset > 0)
          throw new InputError(
            `Под отбором больше ${found.limit} строк, продолжения нет: сузьте отбор и читайте с offset=0.`,
          );
        let items: Array<Record<string, unknown>>;
        let page: Record<string, unknown>;
        let note: string | undefined;
        if (found.capped) {
          // Какие строки вошли в выборку, 1С не обещает: следующая страница могла бы потерять или повторить их.
          items = fitToChars(found.rows.slice(0, limit), renderJson(head).length);
          page = { offset: 0, returned: items.length, truncated: true };
          note =
            `Под отбором больше ${found.limit} строк: показаны ${items.length}, продолжения нет. ` +
            (request.orderBy ? "Это первые по orderBy. " : "Это произвольная часть. ") +
            "Сузьте отбор.";
        } else ({ items, page, note } = pageOf(found.rows, { offset, limit, head }));
        const shown = compact
          ? ((await compactDocument(conn, { rows: items })) as { rows?: Array<Record<string, unknown>> }).rows
          : items;
        return ok({
          ...head,
          ...page,
          ...changes.fields,
          note: [changes.note, note, NOT_TOTALS].filter(Boolean).join(" "),
          rows: shown ?? [],
        });
      }),
  );
}
