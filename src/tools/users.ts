import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { foldText, strictText } from "../odata/inventory.js";
import { departmentNames, readUsers, type UserRow } from "../odata/users.js";
import { listUsersResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok } from "./_shared.js";

const matches = (u: UserRow, query: string, fold: (t: string) => string) => {
  const text = fold(u.name);
  return fold(query)
    .split(/\s+/)
    .filter(Boolean)
    .every((w) => text.includes(w));
};

export function registerUserTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.catalog.list_users",
    {
      title: "Пользователи 1С",
      description:
        "Пользователи 1С (справочник «Пользователи», Бухгалтерия и Управление торговлей) — чтобы выбрать " +
        "ответственного документа (responsibleRef в инструментах создания документов и copy_document). Поля: ref (Ref_Key), " +
        "name, inactive (недействителен), service (служебный), deletionMark, department — подразделение, если задано. " +
        "Контактов, фото и кадровых данных нет. query — части слов имени (все, без учёта регистра, ё как е). По " +
        "умолчанию только действующие: без помеченных на удаление, недействительных и служебных — includeInactive=true, " +
        "чтобы показать всех. refs — проверить привязки: состояние названных пользователей, ненайденные — в missing. count — всего; длинный список читайте с offset=nextOffset, пока он есть.",
      inputSchema: z.object({
        database: databaseField,
        query: z.string().trim().min(1).max(200).optional().describe("Части слов имени: «иван петр»"),
        refs: z
          .array(
            z
              .string()
              .trim()
              .regex(
                /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/,
                "Ref_Key пользователя (GUID)",
              ),
          )
          .min(1)
          .max(200)
          .optional()
          .describe(
            "Проверить привязки: состояние этих пользователей по Ref_Key — действует, недействителен, служебный, помечен на удаление или не найден (missing). Показываются все, независимо от includeInactive",
          ),
        includeInactive: z
          .boolean()
          .default(false)
          .describe("Показывать помеченных на удаление, недействительных и служебных"),
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько строк вернуть"),
      }),
      outputSchema: listUsersResultSchema,
    },
    ({ database, query, refs, includeInactive, limit }) =>
      guard("read.catalog.list_users", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const { users, truncated } = await readUsers(conn);
        // Проверка привязок: названные пользователи — все, в каком бы состоянии ни были; ненайденные — в missing.
        const wanted = refs ? new Set(refs.map((r) => r.replace(/[{}]/g, "").toLowerCase())) : undefined;
        const missing = wanted
          ? [...wanted].filter((r) => !users.some((u) => u.ref.toLowerCase() === r))
          : [];
        const visible = users.filter((u) =>
          wanted
            ? wanted.has(u.ref.toLowerCase())
            : includeInactive || (!u.deletionMark && !u.inactive && !u.service),
        );
        // Как у номенклатуры: сначала буквально, похожие буквы кириллицы и латиницы — только если ничего нет.
        const literal = query ? visible.filter((u) => matches(u, query, strictText)) : visible;
        const found = query && !literal.length ? visible.filter((u) => matches(u, query, foldText)) : literal;
        // Скрытые — среди подходящих под запрос, а не во всём справочнике (ревью 1.1).
        const fits = (u: UserRow) => !query || matches(u, query, strictText) || matches(u, query, foldText);
        const shown = new Set(visible);
        const hidden = users.filter((u) => fits(u) && !shown.has(u)).length;
        // Вторичный ключ обязателен: каждая страница считается заново.
        const sorted = found.sort((a, b) => a.name.localeCompare(b.name, "ru") || a.ref.localeCompare(b.ref));
        const page = sorted.slice(0, limit);
        const names = await departmentNames(
          conn,
          page.flatMap((u) => (u.departmentRef ? [u.departmentRef] : [])),
        );
        const notes: string[] = [];
        if (query && !literal.length && found.length)
          notes.push(
            "Точных совпадений нет: найдено с учётом похожих букв кириллицы и латиницы, проверьте имена.",
          );
        if (!wanted && !includeInactive && hidden)
          notes.push(
            `Скрыто ${hidden} — помеченные на удаление, недействительные и служебные; includeInactive=true, чтобы показать.`,
          );
        if (truncated) notes.push("Справочник прочитан не целиком: уточните query.");
        return ok({
          database: conn.cfg.name,
          ...(query ? { query } : {}),
          ...(missing.length ? { missing } : {}),
          count: sorted.length,
          ...(notes.length ? { note: notes.join(" ") } : {}),
          items: page.map((u) => {
            const department = u.departmentRef ? names.get(u.departmentRef) : undefined;
            return {
              ref: u.ref,
              name: u.name,
              inactive: u.inactive,
              service: u.service,
              deletionMark: u.deletionMark,
              ...(department ? { department } : {}),
              ...(u.departmentRef ? { departmentRef: u.departmentRef } : {}),
            };
          }),
          scan: { rowsScanned: users.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
