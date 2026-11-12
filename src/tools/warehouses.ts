import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { CATALOGS } from "../config/mapping.js";
import { fetchAllForAggregation } from "../odata/aggregate.js";
import { groupPath, strictText, type ItemGroup } from "../odata/inventory.js";
import { requireEntity } from "../odata/publication.js";
import { refOf } from "../odata/settlements-aging.js";
import { listWarehousesResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok } from "./_shared.js";

/** Страница чтения справочника складов: короткие строки, обычно сотни записей — один запрос. */
const PAGE = 5000;

/**
 * Склады справочника для собственного отбора: Ref_Key, название, тип, путь папки. Набор складов (например, только
 * основные склады магазинов) собирает вызывающий по этим фактам: правило отбора в ядре было бы эвристикой с дорогой
 * ошибкой — новый склад молча выпадал бы из итога или попадал в него.
 */
export function registerWarehouseTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.catalog.list_warehouses",
    {
      title: "Склады",
      description:
        "Список складов справочника: ref (Ref_Key — для warehouses в get_warehouse_stock и get_sales_register), name, " +
        "type (ТипСклада: РозничныйМагазин, ОптовыйСклад… — если поле есть в базе), path — путь папки, isFolder, " +
        "deletionMark. Отбор: query — часть названия (без учёта регистра, ё как е), type — точный тип, kind — " +
        "warehouses (по умолчанию), folders или all; помеченные на удаление не показываются без includeDeleted=true. " +
        "Здесь все склады, в том числе без остатка и продаж. Набор складов (например, основные склады магазинов) " +
        "собирайте по этим полям сами. count — всего; длинный список читайте с offset=nextOffset.",
      inputSchema: z.object({
        database: databaseField,
        query: z.string().trim().min(1).max(200).optional().describe("Часть названия склада"),
        type: z
          .string()
          .trim()
          .min(1)
          .max(100)
          .optional()
          .describe("Тип склада, точно: РозничныйМагазин, ОптовыйСклад…"),
        kind: z
          .enum(["warehouses", "folders", "all"])
          .default("warehouses")
          .describe("warehouses — склады (по умолчанию), folders — папки, all — всё"),
        includeDeleted: z.boolean().default(false).describe("Показывать помеченные на удаление"),
        limit: z.number().int().positive().max(1000).default(500).describe("Сколько строк вернуть"),
      }),
      outputSchema: listWarehousesResultSchema,
    },
    ({ database, query, type, kind, includeDeleted, limit }) =>
      guard("read.catalog.list_warehouses", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const set = await requireEntity(conn, CATALOGS.warehouses, "Справочник «Склады»");
        const props = new Set((await conn.getMetadata()).entities.get(set)?.properties.map((p) => p.name));
        const optional = ["ТипСклада", "IsFolder", "Parent_Key", "DeletionMark"].filter((f) => props.has(f));
        const { rows } = await fetchAllForAggregation(
          conn,
          set,
          { select: ["Ref_Key", "Description", ...optional] },
          "справочник складов",
          PAGE,
        );
        const all = rows.map((r) => {
          const parent = refOf(r["Parent_Key"]);
          const typeValue = typeof r["ТипСклада"] === "string" && r["ТипСклада"] ? r["ТипСклада"] : undefined;
          return {
            ref: String(r["Ref_Key"]),
            name: String(r["Description"] ?? ""),
            ...(typeValue ? { type: typeValue } : {}),
            ...(parent ? { parent } : {}),
            isFolder: r["IsFolder"] === true,
            deletionMark: r["DeletionMark"] === true,
          };
        });
        const folders: ItemGroup[] = all
          .filter((w) => w.isFolder)
          .map((w) => ({ ref: w.ref, name: w.name, ...(w.parent ? { parent: w.parent } : {}) }));
        const folderOf = new Map(folders.map((f) => [f.ref, f]));
        const pathOf = (w: { parent?: string }) => {
          const parent = w.parent ? folderOf.get(w.parent) : undefined;
          return parent ? groupPath(parent, folders) : undefined;
        };
        const needle = query ? strictText(query) : undefined;
        const selected = all
          .filter((w) => kind === "all" || (kind === "folders" ? w.isFolder : !w.isFolder))
          .filter((w) => includeDeleted || !w.deletionMark)
          .filter((w) => !type || w.type === type)
          .filter((w) => !needle || strictText(w.name).includes(needle))
          .map((w) => {
            const path = pathOf(w);
            return {
              ref: w.ref,
              name: w.name,
              ...(w.type ? { type: w.type } : {}),
              ...(path ? { path } : {}),
              isFolder: w.isFolder,
              deletionMark: w.deletionMark,
            };
          })
          // Вторичный ключ обязателен: каждая страница считается заново.
          .sort(
            (a, b) =>
              (a.path ?? "").localeCompare(b.path ?? "") ||
              a.name.localeCompare(b.name) ||
              a.ref.localeCompare(b.ref),
          );

        const notes: string[] = [];
        if (type && !props.has("ТипСклада"))
          notes.push("В справочнике складов этой базы нет поля ТипСклада: отбор по типу ничего не нашёл.");
        if (!props.has("IsFolder")) notes.push("Справочник складов без папок.");
        return ok({
          database: conn.cfg.name,
          // Всего, а не на странице: сколько в этом ответе — длина списка и returned.
          count: selected.length,
          ...(notes.length ? { note: notes.join(" ") } : {}),
          items: selected.slice(0, limit),
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
