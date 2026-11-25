import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { resolveEntity } from "../config/mapping.js";
import { resolveNames } from "../odata/accounting.js";
import { callPageStart } from "../odata/call-context.js";
import {
  MAX_ITEM_GROUPS,
  groupPath,
  nomenclatureSet,
  pageGroupNames,
  readPageAttributes,
  selectByItem,
  selectionNotes,
} from "../odata/inventory.js";
import {
  CHARACTERISTICS,
  MAX_WAREHOUSES,
  aggregateWarehouseStock,
  describeWarehouses,
  qty,
  readStockBalance,
  readWarehouses,
  type Warehouse,
  type WarehouseStockLine,
} from "../odata/stock.js";
import { getStockResultSchema } from "../schemas/output.js";
import { databaseField, dateField, guard, ok, snapshotField, snapshotFields } from "./_shared.js";

const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;

type Line = WarehouseStockLine & { itemRef: string };

const sum = (lines: readonly WarehouseStockLine[], field: "inStock" | "toShip") =>
  qty(lines.reduce((t, l) => t + l[field], 0));

export function registerStockTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.analytics.get_warehouse_stock",
    {
      title: "Остатки на складах (УТ 11, КА 2, ERP 2)",
      description:
        "Остатки товаров по регистру «Товары на складах» (оперативный учёт по складам: УТ 11, КА 2, ERP 2; проверено на " +
        "УТ 11.5). В УТ нет Хозрасчётного: количество по складам — здесь, стоимость по регистру «Себестоимость товаров» — " +
        "get_inventory; в Бухгалтерии, КА и ERP остатки по счетам с суммами — get_inventory. Только количество: inStock — в наличии, toShip — к отгрузке " +
        "(зарезервировано под отгрузку, ещё на складе); сумм в этом регистре нет. Назначение, помещение, серия и партия " +
        "сворачиваются. По умолчанию — текущие остатки, asOf=YYYY-MM-DD — на конец даты. warehouses — Ref_Key складов " +
        "(ответ перечисляет их с типом в warehouseFilter); Ref_Key и типы складов с остатком — в byWarehouse при " +
        "groupBy=warehouse на первой странице (у сети в сотни складов он может не поместиться — см. truncatedLists, сузьте " +
        "warehouses). group — группа номенклатуры с подгруппами: Ref_Key, название или путь «Телефоны/Смартфоны/Apple» " +
        "(несколько совпадений — отказ со списком путей); groups — несколько групп за один вызов, итоги каждой в byGroup, " +
        "у позиции matchedGroupRef. query — часть названия или артикул, как в get_inventory. " +
        "Характеристика — characteristicRef, название — если справочник характеристик опубликован; пустая характеристика " +
        "обычна (товар без характеристик). Товары в пути между складами в остаток не входят. count — всего позиций; " +
        "длинный список читайте с offset=nextOffset, пока он есть: следующие страницы в течение минуты берут строки " +
        "первой, регистр заново не читается.",
      inputSchema: z.object({
        database: databaseField,
        asOf: dateField("Дата остатков — на конец этой даты (без параметра — текущие)").optional(),
        warehouses: z
          .array(z.string().trim().regex(GUID, "Ref_Key склада (GUID)"))
          .min(1)
          .max(MAX_WAREHOUSES)
          .optional()
          .describe(
            `Ref_Key складов (до ${MAX_WAREHOUSES}); без параметра — все склады. Группа складов не принимается`,
          ),
        groupBy: z
          .enum(["item", "warehouse"])
          .default("item")
          .describe("item — по номенклатуре (по умолчанию); warehouse — номенклатура в разрезе складов"),
        group: z
          .string()
          .trim()
          .min(1)
          .max(500)
          .optional()
          .describe(
            "Группа номенклатуры: Ref_Key, название или путь через «/» («Смартфоны/Apple»). Подгруппы входят",
          ),
        groups: z
          .array(z.string().trim().min(1).max(500))
          .min(1)
          .max(MAX_ITEM_GROUPS)
          .optional()
          .describe(
            `Несколько групп номенклатуры за один вызов (до ${MAX_ITEM_GROUPS}), каждая как group, с подгруппами; вместо ` +
              "group. Позиция входит один раз; у вложенных групп — в самую глубокую. matchedGroupRef у позиции и byGroup " +
              "с итогами по каждой группе",
          ),
        query: z
          .string()
          .trim()
          .min(1)
          .max(200)
          .optional()
          .describe("Поиск товара: части слов из названия или артикул. Можно вместе с group или groups"),
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько позиций вернуть"),
        snapshot: snapshotField,
      }),
      outputSchema: getStockResultSchema,
    },
    ({ database, asOf, warehouses, groupBy, group, groups: groupList, query, limit, snapshot }) =>
      guard("read.analytics.get_warehouse_stock", async () => {
        const t0 = Date.now();
        const conn = ctx.db(database);
        const byWarehouse = groupBy === "warehouse";
        // Склады проверяются до остатков: неизвестный или группа дали бы тихий ноль.
        const filter: Warehouse[] | undefined = warehouses
          ? await readWarehouses(conn, warehouses)
          : undefined;
        const rows = await readStockBalance(conn, { asOf, warehouses: filter?.map((w) => w.ref) });
        const nomSet = await nomenclatureSet(conn);
        const sel = await selectByItem(conn, nomSet, aggregateWarehouseStock(rows, byWarehouse), {
          group,
          groups: groupList,
          query,
        });
        const { attributes, selectedGroup, selectedGroups, matchedGroup, groups, otherSimilar } = sel;
        const pageStart = callPageStart();

        const withItem = sel.lines
          .filter((l): l is Line => !!l.itemRef)
          // Вторичные ключи обязательны: каждая страница считается заново, а порядок равных остатков не гарантирован.
          .sort(
            (a, b) =>
              b.inStock - a.inStock ||
              b.toShip - a.toShip ||
              a.itemRef.localeCompare(b.itemRef) ||
              (a.characteristicRef ?? "").localeCompare(b.characteristicRef ?? "") ||
              (a.warehouseRef ?? "").localeCompare(b.warehouseRef ?? ""),
          );
        const shown = withItem.slice(0, limit);
        // Строки до начала страницы обёртка отбросит: сведения для них у 1С не запрашиваем.
        const page = shown.slice(pageStart);
        if (nomSet) {
          const unread = page.map((l) => l.itemRef).filter((ref) => !attributes.has(ref));
          for (const [ref, a] of await readPageAttributes(conn, nomSet, unread)) attributes.set(ref, a);
        }
        const parentRefs = [...new Set(page.flatMap((l) => attributes.get(l.itemRef)?.groupRef ?? []))];
        const groupNames = groups
          ? new Map(groups.map((g) => [g.ref, g.name]))
          : nomSet
            ? await pageGroupNames(conn, nomSet, parentRefs)
            : new Map<string, string>();
        const charSet = resolveEntity(CHARACTERISTICS, await conn.available());
        const charRefs = page.flatMap((l) => l.characteristicRef ?? []);
        const charNames =
          charSet && charRefs.length
            ? await resolveNames(conn, charSet, charRefs)
            : new Map<string, string>();
        // Итоги по складам — на первой странице, по всем строкам; на остальных — только склады страницы.
        const totalsShown = byWarehouse && pageStart === 0;
        const warehouseRefs = new Set(
          (totalsShown ? sel.lines : page).flatMap((l) => (l.warehouseRef ? [l.warehouseRef] : [])),
        );
        const known = new Map((filter ?? []).map((w) => [w.ref, w] as const));
        const described = await describeWarehouses(
          conn,
          [...warehouseRefs].filter((r) => !known.has(r)),
        );
        const warehouseOf = (ref: string) => known.get(ref) ?? described.get(ref);

        const items = shown.map((l) => {
          const a = attributes.get(l.itemRef);
          const groupName = a?.groupRef ? groupNames.get(a.groupRef) : undefined;
          const characteristic = l.characteristicRef ? charNames.get(l.characteristicRef) : undefined;
          return {
            item: a?.name ?? l.itemRef,
            ref: l.itemRef,
            ...(a?.article ? { article: a.article } : {}),
            ...(groupName ? { group: groupName } : {}),
            ...(a?.groupRef ? { groupRef: a.groupRef } : {}),
            ...(characteristic ? { characteristic } : {}),
            ...(l.characteristicRef ? { characteristicRef: l.characteristicRef } : {}),
            ...(matchedGroup?.has(l.itemRef) ? { matchedGroupRef: matchedGroup.get(l.itemRef)! } : {}),
            ...(l.warehouseRef
              ? {
                  warehouse: warehouseOf(l.warehouseRef)?.name ?? l.warehouseRef,
                  warehouseRef: l.warehouseRef,
                }
              : {}),
            inStock: l.inStock,
            toShip: l.toShip,
          };
        });

        // Итоги по каждой группе из groups — на первой странице, по всем строкам отбора: сумма равна общим итогам.
        const byGroup =
          selectedGroups && groups && pageStart === 0
            ? selectedGroups.map((g) => {
                const own = withItem.filter((l) => matchedGroup?.get(l.itemRef) === g.ref);
                return {
                  name: g.name,
                  ref: g.ref,
                  path: groupPath(g, groups),
                  positions: own.length,
                  inStock: sum(own, "inStock"),
                  toShip: sum(own, "toShip"),
                };
              })
            : undefined;
        const notes = await selectionNotes(
          conn,
          nomSet,
          sel,
          { group, groups: groupList, query },
          "остатка на складах",
        );
        if (!charSet && page.some((l) => l.characteristicRef)) {
          notes.push(
            "Справочник характеристик не опубликован в OData: у строк с характеристикой только characteristicRef.",
          );
        }
        // Метка итогов: страница через минуту читает регистр заново, и если в 1С провели документы, склеивать нельзя.
        const changes = snapshotFields(
          { count: withItem.length, inStock: sum(sel.lines, "inStock"), toShip: sum(sel.lines, "toShip") },
          snapshot,
        );
        if (changes.note) notes.unshift(changes.note);
        const noItem = sel.lines.filter((l) => !l.itemRef);
        if (noItem.length) notes.push("Строки без номенклатуры не показаны, но входят в итоги.");

        const byWarehouseTotals = new Map<string, { positions: number; inStock: number; toShip: number }>();
        for (const l of sel.lines) {
          const t = byWarehouseTotals.get(l.warehouseRef ?? "") ?? { positions: 0, inStock: 0, toShip: 0 };
          if (l.itemRef) t.positions += 1;
          t.inStock += l.inStock;
          t.toShip += l.toShip;
          byWarehouseTotals.set(l.warehouseRef ?? "", t);
        }

        return ok({
          database: conn.cfg.name,
          ...(asOf ? { asOf } : {}),
          ...(byWarehouse ? { groupBy } : {}),
          ...(filter ? { warehouseFilter: filter } : {}),
          ...(selectedGroup
            ? { itemGroup: { name: selectedGroup.name, ref: selectedGroup.ref, includesSubgroups: true } }
            : {}),
          ...(selectedGroups && groups
            ? {
                itemGroups: selectedGroups.map((g) => ({
                  name: g.name,
                  ref: g.ref,
                  path: groupPath(g, groups),
                  includesSubgroups: true,
                })),
              }
            : {}),
          ...(byGroup ? { byGroup } : {}),
          ...(query ? { query } : {}),
          ...(otherSimilar.length
            ? {
                otherSimilar: {
                  // Позиции, а не строки: при groupBy=warehouse одна позиция на трёх складах — одна.
                  count: new Set(otherSimilar.map((l) => l.itemRef)).size,
                  inStock: sum(otherSimilar, "inStock"),
                  examples: [...new Set(otherSimilar.map((l) => attributes.get(l.itemRef ?? "")?.name ?? ""))]
                    .filter(Boolean)
                    .slice(0, 5),
                },
              }
            : {}),
          totalInStock: sum(sel.lines, "inStock"),
          totalToShip: sum(sel.lines, "toShip"),
          // Всего, а не на странице: сколько в этом ответе — длина списка и returned.
          count: withItem.length,
          ...changes.fields,
          ...(totalsShown
            ? {
                byWarehouse: [...byWarehouseTotals]
                  .map(([ref, t]) => {
                    const w = ref ? warehouseOf(ref) : undefined;
                    return {
                      warehouse: ref ? (w?.name ?? ref) : null,
                      ...(ref ? { warehouseRef: ref } : {}),
                      ...(w?.type ? { type: w.type } : {}),
                      positions: t.positions,
                      inStock: qty(t.inStock),
                      toShip: qty(t.toShip),
                    };
                  })
                  .sort(
                    (a, b) =>
                      b.inStock - a.inStock || String(a.warehouseRef).localeCompare(String(b.warehouseRef)),
                  ),
              }
            : {}),
          ...(notes.length ? { note: notes.join(" ") } : {}),
          items,
          scan: { rowsScanned: rows.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
