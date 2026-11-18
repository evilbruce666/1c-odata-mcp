import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { CATALOGS } from "../config/mapping.js";
import { InputError } from "../errors.js";
import {
  CHARACTERISTICS_USAGE,
  MAX_ITEM_GROUPS,
  articlePublished,
  foldText,
  groupOwners,
  groupPath,
  loadItemGroups,
  matchesItemQuery,
  readAllItems,
  readItemsInGroups,
  resolveItemGroup,
  strictText,
  type ItemAttributes,
  type ItemGroup,
} from "../odata/inventory.js";
import { requireEntity } from "../odata/publication.js";
import { listItemsResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok } from "./_shared.js";

interface Row {
  ref: string;
  name: string;
  article?: string;
  code?: string;
  groupRef?: string;
  path?: string;
  isFolder: boolean;
  deletionMark: boolean;
  characteristicsUsage?: string;
  matchedGroupRef?: string;
}

/**
 * Позиции и папки справочника номенклатуры — все, в том числе без остатка и продаж: остатки и продажи показывают только
 * позиции с движением, и ноль там нельзя отличить от позиции, которой в базе нет. Отбор по группам и поиск — те же,
 * что у остатков и продаж (одни функции), поэтому список и остатки по одной группе сопоставляются по ref.
 */
export function registerItemTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.catalog.list_items",
    {
      title: "Номенклатура",
      description:
        "Список номенклатуры справочника — все позиции, в том числе без остатка и продаж (get_warehouse_stock, " +
        "get_sales_register и get_inventory показывают только позиции с движением). Поля: ref (Ref_Key — как ref в " +
        "остатках и продажах), name, article, code, groupRef и path — папка позиции и её путь (у папки — её собственный " +
        "путь), isFolder, deletionMark, characteristicsUsage — реквизит ИспользованиеХарактеристик как есть (УТ, КА, " +
        "ERP; «ОбщиеДляВидаНоменклатуры» — характеристики задаются видом номенклатуры). Отбор: group или groups — группа " +
        "или несколько групп с подгруппами, как в get_warehouse_stock (у позиции matchedGroupRef); query — части слов " +
        "названия или артикул; kind — items (по умолчанию), folders или all; помеченные на удаление не показываются без " +
        "includeDeleted=true. count — всего; длинный список читайте с offset=nextOffset, пока он есть.",
      inputSchema: z.object({
        database: databaseField,
        group: z
          .string()
          .trim()
          .min(1)
          .max(500)
          .optional()
          .describe("Группа номенклатуры: Ref_Key, название или путь через «/». Подгруппы входят"),
        groups: z
          .array(z.string().trim().min(1).max(500))
          .min(1)
          .max(MAX_ITEM_GROUPS)
          .optional()
          .describe(
            `Несколько групп (до ${MAX_ITEM_GROUPS}), каждая как group, с подгруппами; вместо group. ` +
              "У вложенных групп позиция относится к самой глубокой (matchedGroupRef)",
          ),
        query: z
          .string()
          .trim()
          .min(1)
          .max(200)
          .optional()
          .describe("Части слов из названия или артикул; у папок — из названия"),
        kind: z
          .enum(["items", "folders", "all"])
          .default("items")
          .describe("items — позиции (по умолчанию), folders — папки (при group — подгруппы), all — всё"),
        includeDeleted: z.boolean().default(false).describe("Показывать помеченные на удаление"),
        limit: z.number().int().positive().max(1000).default(500).describe("Сколько строк вернуть"),
      }),
      outputSchema: listItemsResultSchema,
    },
    ({ database, group, groups: groupList, query, kind, includeDeleted, limit }) =>
      guard("read.catalog.list_items", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const set = await requireEntity(conn, CATALOGS.nomenclature, "Справочник «Номенклатура»");
        if (group && groupList?.length) throw new InputError("Укажите group или groups, не оба сразу.");
        const groups = await loadItemGroups(conn, set);
        const folderOf = new Map(groups.map((g) => [g.ref, g]));
        const pathOf = (ref: string | undefined) => {
          const folder = ref ? folderOf.get(ref) : undefined;
          return folder ? groupPath(folder, groups) : undefined;
        };

        const wanted = groupList?.length ? groupList : group ? [group] : [];
        const selected = [
          ...new Map(wanted.map((q) => resolveItemGroup(q, groups)).map((g) => [g.ref, g])).values(),
        ];
        const owner = selected.length ? groupOwners(selected, groups) : undefined;
        const items: Map<string, ItemAttributes> =
          kind === "folders"
            ? new Map()
            : owner
              ? await readItemsInGroups(conn, set, owner.keys())
              : await readAllItems(conn, set);
        const roots = new Set(selected.map((g) => g.ref));
        const folders: ItemGroup[] =
          kind === "items" ? [] : groups.filter((g) => !owner || (owner.has(g.ref) && !roots.has(g.ref)));

        const rows: Row[] = [
          ...[...items].map(([ref, a]): Row => {
            const path = pathOf(a.groupRef);
            const matched = owner && a.groupRef ? owner.get(a.groupRef) : undefined;
            return {
              ref,
              name: a.name ?? "",
              ...(a.article ? { article: a.article } : {}),
              ...(a.code ? { code: a.code } : {}),
              ...(a.groupRef ? { groupRef: a.groupRef } : {}),
              ...(path ? { path } : {}),
              isFolder: false,
              deletionMark: a.deletionMark === true,
              ...(a.characteristicsUsage ? { characteristicsUsage: a.characteristicsUsage } : {}),
              ...(groupList?.length && matched ? { matchedGroupRef: matched } : {}),
            };
          }),
          ...folders.map((g): Row => {
            const matched = owner?.get(g.ref);
            return {
              ref: g.ref,
              name: g.name,
              ...(g.code ? { code: g.code } : {}),
              ...(g.parent ? { groupRef: g.parent } : {}),
              path: groupPath(g, groups),
              isFolder: true,
              deletionMark: g.deletionMark === true,
              ...(groupList?.length && matched ? { matchedGroupRef: matched } : {}),
            };
          }),
        ].filter((r) => includeDeleted || !r.deletionMark);

        // Как у остатков: сначала буквально, похожие буквы кириллицы и латиницы — только если ничего нет.
        const matching = (fold: (text: string) => string) =>
          rows.filter((r) =>
            matchesItemQuery({ name: r.name, ...(r.article ? { article: r.article } : {}) }, query!, fold),
          );
        const literal = query ? matching(strictText) : rows;
        const found = query && !literal.length ? matching(foldText) : literal;
        const lookalike = !!query && !literal.length && found.length > 0;

        // Вторичные ключи обязательны: каждая страница считается заново.
        const sorted = found.sort(
          (a, b) =>
            (a.path ?? "").localeCompare(b.path ?? "") ||
            Number(b.isFolder) - Number(a.isFolder) ||
            a.name.localeCompare(b.name) ||
            a.ref.localeCompare(b.ref),
        );

        const notes: string[] = [];
        if (lookalike)
          notes.push(
            "Точных совпадений нет: найдено с учётом похожих букв кириллицы и латиницы, проверьте позиции.",
          );
        if (query && kind !== "folders" && !(await articlePublished(conn, set)))
          notes.push("Артикул не опубликован в OData: поиск шёл только по названию.");
        const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));
        if (kind !== "folders" && !props.has(CHARACTERISTICS_USAGE))
          notes.push(
            "В справочнике этой базы нет реквизита ИспользованиеХарактеристик: characteristicsUsage не указан.",
          );
        if (selected.length && !sorted.length)
          notes.push(
            query
              ? "В выбранных группах ничего не найдено по запросу; вне групп поиск не проверял."
              : "В выбранных группах нет позиций (с учётом includeDeleted и kind).",
          );

        return ok({
          database: conn.cfg.name,
          ...(selected.length
            ? {
                itemGroups: selected.map((g) => ({
                  name: g.name,
                  ref: g.ref,
                  path: groupPath(g, groups),
                  includesSubgroups: true,
                })),
              }
            : {}),
          ...(query ? { query } : {}),
          // Всего, а не на странице: сколько в этом ответе — длина списка и returned.
          count: sorted.length,
          ...(notes.length ? { note: notes.join(" ") } : {}),
          items: sorted.slice(0, limit),
          scan: { rowsScanned: items.size + groups.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
