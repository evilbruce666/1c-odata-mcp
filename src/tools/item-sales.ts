import type { McpServer } from "@modelcontextprotocol/server";
import { USERS } from "../odata/users.js";
/** Подразделение строки продаж: в УТ — структура предприятия. */
const DEPARTMENT_SETS = ["Catalog_СтруктураПредприятия", "Catalog_ПодразделенияОрганизаций"] as const;
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
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
  aggregateSales,
  checkPeriod,
  isOwnBuyer,
  lastRetailReports,
  ownBuyers,
  readItemSales,
  toSalesRows,
  type SalesLine,
  type SalesRow,
} from "../odata/item-sales.js";
import {
  CHARACTERISTICS,
  MAX_WAREHOUSES,
  describeWarehouses,
  qty,
  readWarehouses,
  type Warehouse,
} from "../odata/stock.js";
import { getItemSalesResultSchema } from "../schemas/output.js";
import { databaseField, dateField, guard, ok, snapshotField, snapshotFields } from "./_shared.js";

const GUID = /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/;

/** Операция продажи в розницу: она попадает в регистр только отчётом о розничных продажах после закрытия смены. */
const RETAIL_OPERATION = "РеализацияВРозницу";

type Line = SalesLine & { itemRef: string };

const total = (rows: readonly SalesRow[], pick: (q: number) => number) =>
  qty(rows.reduce((t, r) => t + pick(r.quantity), 0));

export function registerItemSalesTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.analytics.get_sales_register",
    {
      title: "Продажи по номенклатуре (УТ 11, КА 2, ERP 2)",
      description:
        "Продажи товаров в штуках за период по регистру «Выручка и себестоимость продаж» (УТ 11, КА 2, ERP 2; проверено " +
        "на УТ 11.5), включая розницу (в Бухгалтерии продажи по документам — get_sales_breakdown). quantity — нетто: sold (продано) минус returns (возвраты и сторно " +
        "продаж, в регистре они с минусом); byOperation — то же по хозяйственным операциям (РеализацияВРозницу, " +
        "РеализацияКлиенту, ВозвратТоваровОтКлиентаПрошлыхПериодов, СторноРеализации…). Розница попадает в регистр только " +
        "отчётом о розничных продажах после закрытия смены: retailReports.completeThrough — дата, по которую розница " +
        "видна по всем складам выборки, latest — самый свежий отчёт; какой склад отстаёт — отбор warehouses. Отбор: warehouses — Ref_Key складов (как в get_warehouse_stock), group — группа " +
        "номенклатуры с подгруппами (Ref_Key, название или путь «Телефоны/Смартфоны/Apple»), groups — несколько групп за " +
        "один вызов (итоги каждой в byGroup, у позиции matchedGroupRef), query — часть названия или " +
        "артикул. managers — только продажи этих менеджеров (Ref_Key пользователей 1С, read.catalog.list_users); byManager=true — итоги по менеджерам в byManager; departments и byDepartment — то же по подразделениям. groupBy=warehouse — по складам. excludeOwnOrganizations=true — без продаж своим организациям: в УТ это " +
        "обычная продажа клиенту, и покупатель узнаётся по ИНН, совпавшему с ИНН организации базы; исключённые штуки и " +
        "выручка и все контрагенты, признанные своими по ИНН, — в excludedOwnOrganizations (организация без ИНН так не узнаётся, их число там же). revenue и revenueExVat — выручка с НДС и без НДС, нетто (возвраты и сторно " +
        "вычитаются), в валюте управленческого учёта базы (обычно рубли): розница с комиссионным товаром сверена с отчётами " +
        "о розничных продажах до рубля. count — всего позиций; длинный список " +
        "читайте с offset=nextOffset, пока он есть: следующие страницы в течение минуты берут строки первой, регистр " +
        "заново не читается.",
      inputSchema: z.object({
        database: databaseField,
        from: dateField("Начало периода (включительно)"),
        to: dateField("Конец периода (включительно)"),
        warehouses: z
          .array(z.string().trim().regex(GUID, "Ref_Key склада (GUID)"))
          .min(1)
          .max(MAX_WAREHOUSES)
          .optional()
          .describe(
            `Ref_Key складов (до ${MAX_WAREHOUSES}); без параметра — все, включая продажи без склада. Группа складов не принимается`,
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
        managers: z
          .array(z.string().trim().regex(GUID, "Ref_Key пользователя (GUID)"))
          .min(1)
          .max(50)
          .optional()
          .describe(
            "Только продажи этих менеджеров — Ref_Key пользователей 1С (read.catalog.list_users); менеджер — измерение регистра продаж",
          ),
        byManager: z
          .boolean()
          .default(false)
          .describe(
            "Итоги по менеджерам (byManager) на первой странице: штуки и выручка каждого, без менеджера — отдельно",
          ),
        departments: z
          .array(z.string().trim().regex(GUID, "Ref_Key подразделения (GUID)"))
          .min(1)
          .max(50)
          .optional()
          .describe("Только продажи этих подразделений — Ref_Key (их ссылки и названия — в byDepartment)"),
        byDepartment: z
          .boolean()
          .default(false)
          .describe("Итоги по подразделениям (byDepartment) на первой странице: штуки и выручка каждого"),
        excludeOwnOrganizations: z
          .boolean()
          .default(false)
          .describe(
            "Не считать продажи своим организациям: покупатель-контрагент с ИНН одной из организаций базы. Исключённое — в excludedOwnOrganizations",
          ),
        limit: z.number().int().positive().max(1000).default(200).describe("Сколько позиций вернуть"),
        snapshot: snapshotField,
      }),
      outputSchema: getItemSalesResultSchema,
    },
    ({
      database,
      from,
      to,
      warehouses,
      groupBy,
      group,
      groups: groupList,
      query,
      excludeOwnOrganizations,
      managers,
      byManager,
      departments,
      byDepartment,
      limit,
      snapshot,
    }) =>
      guard("read.analytics.get_sales_register", async () => {
        const t0 = Date.now();
        checkPeriod(from, to);
        const conn = ctx.db(database);
        const byWarehouse = groupBy === "warehouse";
        // Склады проверяются до оборотов: неизвестный или группа дали бы тихий ноль.
        const filter: Warehouse[] | undefined = warehouses
          ? await readWarehouses(conn, warehouses)
          : undefined;
        const warehouseSet = resolveEntity(CATALOGS.warehouses, await conn.available());
        const raw = await readItemSales(conn, {
          from,
          to,
          warehouses: filter?.map((w) => w.ref),
          warehouseSet,
          withBuyer: excludeOwnOrganizations,
          withManager: byManager || !!managers?.length,
          withDepartment: byDepartment || !!departments?.length,
        });
        const nomSet = await nomenclatureSet(conn);
        // Отбор по товару — по строкам до свёртки: из них же считаются итоги по операциям.
        const wanted = managers?.length
          ? new Set(managers.map((m) => m.replace(/[{}]/g, "").toLowerCase()))
          : undefined;
        const wantedDeps = departments?.length
          ? new Set(departments.map((d) => d.replace(/[{}]/g, "").toLowerCase()))
          : undefined;
        const salesRows = toSalesRows(raw, CATALOGS.warehouses).filter(
          (r) =>
            (!wanted || (!!r.managerRef && wanted.has(r.managerRef.toLowerCase()))) &&
            (!wantedDeps || (!!r.departmentRef && wantedDeps.has(r.departmentRef.toLowerCase()))),
        );
        const sel = await selectByItem(conn, nomSet, salesRows, {
          group,
          groups: groupList,
          query,
        });
        // Свои организации-покупатели узнаются по ИНН; исключённое считается после отбора по товару и складу.
        const own = excludeOwnOrganizations ? await ownBuyers(conn) : undefined;
        const orgSet = own ? resolveEntity(CATALOGS.organizations, await conn.available()) : undefined;
        const kept = own ? sel.lines.filter((r) => !isOwnBuyer(r, own, orgSet)) : sel.lines;
        const excluded = own ? sel.lines.filter((r) => isOwnBuyer(r, own, orgSet)) : [];
        const { attributes, selectedGroup, selectedGroups, matchedGroup, groups, otherSimilar } = sel;
        const lines = aggregateSales(kept, byWarehouse);
        const pageStart = callPageStart();

        const withItem = lines
          .filter((l): l is Line => !!l.itemRef)
          // Вторичные ключи обязательны: каждая страница считается заново, а порядок равных продаж не гарантирован.
          .sort(
            (a, b) =>
              b.quantity - a.quantity ||
              b.sold - a.sold ||
              a.itemRef.localeCompare(b.itemRef) ||
              (a.characteristicRef ?? "").localeCompare(b.characteristicRef ?? "") ||
              (a.warehouseRef ?? "").localeCompare(b.warehouseRef ?? ""),
          );
        const shown = withItem.slice(0, limit);
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

        // Сводки — на первой странице, по всем строкам отбора; на остальных они вытесняли бы сами строки.
        const first = pageStart === 0;
        const retailWarehouses = new Set(
          kept.filter((r) => r.operation === RETAIL_OPERATION && r.warehouseRef).map((r) => r.warehouseRef!),
        );
        // Отчёты о рознице — только по складам с розничными продажами: розница в регистре и есть их отчёты.
        const reports =
          first && retailWarehouses.size
            ? await lastRetailReports(conn, from, to, retailWarehouses)
            : undefined;
        // Только факты, без списка «отстающих»: любое правило отставания ошибалось — вечером, когда часть смен уже
        // закрыта сегодняшним отчётом, отстающими выглядели почти все склады (три находки ревью подряд).
        const lastDates = reports ? [...retailWarehouses].map((w) => reports.last.get(w) ?? "") : [];
        const latest = lastDates.length ? [...lastDates].sort().at(-1) || undefined : undefined;
        const scope = new Set([
          ...(first && byWarehouse ? kept.flatMap((r) => (r.warehouseRef ? [r.warehouseRef] : [])) : []),
          ...page.flatMap((l) => (l.warehouseRef ? [l.warehouseRef] : [])),
        ]);
        const known = new Map((filter ?? []).map((w) => [w.ref, w] as const));
        const described = await describeWarehouses(
          conn,
          [...scope].filter((r) => !known.has(r)),
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
            quantity: l.quantity,
            sold: l.sold,
            returns: l.returns,
            revenue: l.revenue,
            revenueExVat: l.revenueExVat,
          };
        });

        // Итоги по каждой группе из groups — на первой странице, по всем строкам отбора: сумма равна общим итогам.
        const byGroup =
          selectedGroups && groups && pageStart === 0
            ? selectedGroups.map((g) => {
                const own = withItem.filter((l) => matchedGroup?.get(l.itemRef) === g.ref);
                const add = (pick: (l: Line) => number) => own.reduce((t, l) => t + pick(l), 0);
                return {
                  name: g.name,
                  ref: g.ref,
                  path: groupPath(g, groups),
                  positions: own.length,
                  quantity: qty(add((l) => l.quantity)),
                  sold: qty(add((l) => l.sold)),
                  returns: qty(add((l) => l.returns)),
                  // В копейках: строки уже округлены до копейки, сумма float накопила бы хвост.
                  revenue: add((l) => Math.round(l.revenue * 100)) / 100,
                  revenueExVat: add((l) => Math.round(l.revenueExVat * 100)) / 100,
                };
              })
            : undefined;
        const notes = await selectionNotes(
          conn,
          nomSet,
          sel,
          { group, groups: groupList, query },
          "продаж за период",
        );
        if (!charSet && page.some((l) => l.characteristicRef)) {
          notes.push(
            "Справочник характеристик не опубликован в OData: у строк с характеристикой только characteristicRef.",
          );
        }
        // Метка итогов: страница через минуту читает регистр заново, и если в 1С провели документы, склеивать нельзя.
        const changes = snapshotFields(
          {
            count: withItem.length,
            quantity: total(kept, (q) => q),
            revenue: kept.reduce((t, r) => t + r.revenueCents, 0),
          },
          snapshot,
        );
        if (changes.note) notes.unshift(changes.note);
        const noItem = lines.filter((l) => !l.itemRef);
        if (noItem.length) notes.push("Строки без номенклатуры не показаны, но входят в итоги.");
        if (byWarehouse && first && kept.some((r) => !r.warehouseRef)) {
          notes.push("Часть продаж без склада: они в byWarehouse со значением warehouse: null.");
        }
        if (reports?.unknown.length) {
          notes.push(
            `По ${reports.unknown.length} складам дату последнего отчёта о рознице узнать не удалось (слишком много складов или отчётов): retailReports.completeThrough не указан.`,
          );
        }
        if (first && retailWarehouses.size && !reports) {
          notes.push(
            "Отчёты о розничных продажах не опубликованы: по какую дату видна розница, ответ сказать не может.",
          );
        }

        const byOperation = new Map<
          string,
          { quantity: number; sold: number; returns: number; revenueCents: number; revenueExVatCents: number }
        >();
        for (const r of kept) {
          const t = byOperation.get(r.operation) ?? {
            quantity: 0,
            sold: 0,
            returns: 0,
            revenueCents: 0,
            revenueExVatCents: 0,
          };
          t.quantity += r.quantity;
          t.revenueCents += r.revenueCents;
          t.revenueExVatCents += r.revenueExVatCents;
          if (r.quantity >= 0) t.sold += r.quantity;
          else t.returns -= r.quantity;
          byOperation.set(r.operation, t);
        }
        const byWarehouseTotals = new Map<
          string,
          { positions: Set<string>; quantity: number; revenueCents: number }
        >();
        for (const r of kept) {
          const t = byWarehouseTotals.get(r.warehouseRef ?? "") ?? {
            positions: new Set<string>(),
            quantity: 0,
            revenueCents: 0,
          };
          if (r.itemRef) t.positions.add(`${r.itemRef}|${r.characteristicRef ?? ""}`);
          t.quantity += r.quantity;
          t.revenueCents += r.revenueCents;
          byWarehouseTotals.set(r.warehouseRef ?? "", t);
        }

        // Итоги по менеджерам (пользователи 1С): на первой странице, по всем строкам отбора.
        const managerTotals = new Map<
          string,
          { quantity: number; revenueCents: number; revenueExVatCents: number }
        >();
        if (byManager && first)
          for (const r of kept) {
            const t = managerTotals.get(r.managerRef ?? "") ?? {
              quantity: 0,
              revenueCents: 0,
              revenueExVatCents: 0,
            };
            t.quantity += r.quantity;
            t.revenueCents += r.revenueCents;
            t.revenueExVatCents += r.revenueExVatCents;
            managerTotals.set(r.managerRef ?? "", t);
          }
        const departmentTotals = new Map<
          string,
          { quantity: number; revenueCents: number; revenueExVatCents: number }
        >();
        if (byDepartment && first)
          for (const r of kept) {
            const t = departmentTotals.get(r.departmentRef ?? "") ?? {
              quantity: 0,
              revenueCents: 0,
              revenueExVatCents: 0,
            };
            t.quantity += r.quantity;
            t.revenueCents += r.revenueCents;
            t.revenueExVatCents += r.revenueExVatCents;
            departmentTotals.set(r.departmentRef ?? "", t);
          }
        const depSet = resolveEntity(DEPARTMENT_SETS, await conn.available());
        const depRefs = [
          ...departmentTotals.keys(),
          ...(departments ?? []).map((d) => d.replace(/[{}]/g, "").toLowerCase()),
        ].filter(Boolean);
        const departmentNames =
          depSet && depRefs.length
            ? await resolveNames(conn, depSet, depRefs).catch(() => new Map<string, string>())
            : new Map<string, string>();
        const managerRefs = [
          ...managerTotals.keys(),
          ...(managers ?? []).map((m) => m.replace(/[{}]/g, "").toLowerCase()),
        ].filter(Boolean);
        const managerNames = managerRefs.length
          ? await resolveNames(conn, USERS, managerRefs).catch(() => new Map<string, string>())
          : new Map<string, string>();

        return ok({
          database: conn.cfg.name,
          from,
          to,
          ...(byWarehouse ? { groupBy } : {}),
          ...(managers?.length
            ? {
                managerFilter: [...new Set(managers.map((m) => m.replace(/[{}]/g, "").toLowerCase()))].map(
                  (m) => ({
                    ref: m,
                    name: managerNames.get(m) ?? null,
                  }),
                ),
              }
            : {}),
          ...(departments?.length
            ? {
                departmentFilter: [
                  ...new Set(departments.map((d) => d.replace(/[{}]/g, "").toLowerCase())),
                ].map((d) => ({ ref: d, name: departmentNames.get(d) ?? null })),
              }
            : {}),
          ...(byDepartment && first
            ? {
                byDepartment: [...departmentTotals]
                  .map(([ref, t]) => ({
                    department: ref ? (departmentNames.get(ref) ?? ref) : null,
                    ...(ref ? { departmentRef: ref } : {}),
                    quantity: qty(t.quantity),
                    revenue: t.revenueCents / 100,
                    revenueExVat: t.revenueExVatCents / 100,
                  }))
                  .sort(
                    (a, b) =>
                      b.revenue - a.revenue || String(a.departmentRef).localeCompare(String(b.departmentRef)),
                  ),
              }
            : {}),
          ...(byManager && first
            ? {
                byManager: [...managerTotals]
                  .map(([ref, t]) => ({
                    manager: ref ? (managerNames.get(ref) ?? ref) : null,
                    ...(ref ? { managerRef: ref } : {}),
                    quantity: qty(t.quantity),
                    revenue: t.revenueCents / 100,
                    revenueExVat: t.revenueExVatCents / 100,
                  }))
                  .sort(
                    (a, b) =>
                      b.revenue - a.revenue || String(a.managerRef).localeCompare(String(b.managerRef)),
                  ),
              }
            : {}),
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
                  count: new Set(otherSimilar.map((r) => r.itemRef)).size,
                  quantity: total(otherSimilar, (q) => q),
                  examples: [...new Set(otherSimilar.map((r) => attributes.get(r.itemRef ?? "")?.name ?? ""))]
                    .filter(Boolean)
                    .slice(0, 5),
                },
              }
            : {}),
          ...(own
            ? {
                excludedOwnOrganizations: {
                  quantity: total(excluded, (q) => q),
                  revenue: excluded.reduce((t, r) => t + r.revenueCents, 0) / 100,
                  revenueExVat: excluded.reduce((t, r) => t + r.revenueExVatCents, 0) / 100,
                  counterparties: own.counterparties,
                  ...(own.organizationsWithoutInn
                    ? { organizationsWithoutInn: own.organizationsWithoutInn }
                    : {}),
                },
              }
            : {}),
          totalQuantity: total(kept, (q) => q),
          sold: total(kept, (q) => Math.max(q, 0)),
          returns: total(kept, (q) => Math.max(-q, 0)),
          revenue: kept.reduce((t, r) => t + r.revenueCents, 0) / 100,
          revenueExVat: kept.reduce((t, r) => t + r.revenueExVatCents, 0) / 100,
          // Всего, а не на странице: сколько в этом ответе — длина списка и returned (его добавляет обёртка страниц).
          count: withItem.length,
          ...changes.fields,
          ...(first
            ? {
                byOperation: [...byOperation]
                  .map(([operation, t]) => ({
                    operation,
                    quantity: qty(t.quantity),
                    sold: qty(t.sold),
                    returns: qty(t.returns),
                    revenue: t.revenueCents / 100,
                    revenueExVat: t.revenueExVatCents / 100,
                  }))
                  .sort((a, b) => b.quantity - a.quantity || a.operation.localeCompare(b.operation)),
              }
            : {}),
          ...(first && byWarehouse
            ? {
                byWarehouse: [...byWarehouseTotals]
                  .map(([ref, t]) => {
                    const w = ref ? warehouseOf(ref) : undefined;
                    return {
                      warehouse: ref ? (w?.name ?? ref) : null,
                      ...(ref ? { warehouseRef: ref } : {}),
                      ...(w?.type ? { type: w.type } : {}),
                      positions: t.positions.size,
                      quantity: qty(t.quantity),
                      revenue: t.revenueCents / 100,
                    };
                  })
                  .sort(
                    (a, b) =>
                      b.quantity - a.quantity || String(a.warehouseRef).localeCompare(String(b.warehouseRef)),
                  ),
              }
            : {}),
          ...(reports
            ? {
                retailReports: {
                  warehouses: retailWarehouses.size,
                  latest: latest ?? null,
                  // Розница видна по всем складам выборки по эту дату включительно; null — у кого-то дата неизвестна.
                  completeThrough: reports.unknown.length ? null : [...lastDates].sort()[0] || null,
                  ...(reports.unknown.length ? { unknown: reports.unknown.length } : {}),
                },
              }
            : {}),
          ...(notes.length ? { note: notes.join(" ") } : {}),
          items,
          scan: { rowsScanned: raw.length, elapsedMs: Date.now() - t0 },
        });
      }),
  );
}
