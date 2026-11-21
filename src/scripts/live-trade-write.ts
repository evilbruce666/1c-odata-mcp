/**
 * Живой прогон записи Управления торговлей против НАСТОЯЩЕЙ, лучше тестовой базы (копии рабочей): тестовые объекты
 * ZZ-TEST и документы склада, денег, продаж и закупок через инструменты — создание (предпросмотр → подтверждение),
 * проведение со сверкой движений, отмена проведения, пометка удаления. Данные базы в скрипте не зашиты: организация
 * и склады — параметрами, статьи, вид цен и подразделение — из последних проведённых документов того же вида.
 *
 *   npm run test:live:ut -- --db <база> --org <организация> --warehouse <склад> --yes [--warehouse2 <склад>]
 *                            [--cash <касса>] [--bank <банковский счёт организации>]
 *
 * Тестовые справочники (партнёр, контрагент, договоры, соглашения, номенклатура «ZZ-TEST live …») находятся по
 * наименованию и создаются, только если их нет; их не помечает (номенклатуру УТ 11.5 через OData пометить нельзя).
 * Документы — организации из --org с комментарием «ZZ-TEST live …»; все помечаются на удаление в конце.
 * Журнал операций — во временном каталоге. Код возврата 1 — если хоть одна проверка не прошла.
 */
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

process.env["ODATA_WRITE_JOURNAL_DIR"] = mkdtempSync(join(tmpdir(), "1c-odata-live-ut-"));

type Json = Record<string, unknown>;
type Result = { err: boolean; sc: Json; text: string };
type Tool = {
  inputSchema: { safeParse: (v: unknown) => { success: boolean; data?: Json; error?: unknown } };
  handler: (
    a: Json,
    e: Json,
  ) => Promise<{ isError?: boolean; structuredContent?: Json; content?: Array<{ text?: string }> }>;
};

const arg = (name: string): string | undefined => {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const line = (s = "") => process.stdout.write(`${s}\n`);
const EMPTY = /^0{8}-/;
const q = (s: string) => s.replace(/'/g, "''");

async function main(): Promise<number> {
  const db = arg("db");
  const orgName = arg("org");
  const warehouse = arg("warehouse");
  if (!db || !orgName || !warehouse || !process.argv.includes("--yes")) {
    line(
      "Пишет в базу Управления торговлей. Запуск: npm run test:live:ut -- --db <база> --org <организация> --warehouse <склад> --yes",
    );
    return 2;
  }
  const { loadConfig } = await import("../config/env.js");
  const { ServerContext } = await import("../context.js");
  const { createServer } = await import("../mcp/server.js");
  const cfg = loadConfig();
  const ctx = new ServerContext(cfg);
  const conn = ctx.db(db);
  if (cfg.behavior.readOnly || !conn.cfg.writable) {
    line(`База "${db}" недоступна для записи: нужны READ_ONLY=false и WRITABLE для базы.`);
    return 2;
  }
  const tools = (createServer(ctx) as unknown as { _registeredTools: Record<string, Tool> })._registeredTools;
  const get = async (path: string) => (await conn.client.getCollection(path)).value as Json[];
  const run = async (name: string, a: Json): Promise<Result> => {
    const t = tools[name];
    if (!t) return { err: true, sc: {}, text: `нет инструмента ${name}` };
    const parsed = t.inputSchema.safeParse({ database: db, ...a });
    if (!parsed.success)
      return { err: true, sc: {}, text: `входные данные: ${JSON.stringify(parsed.error)}` };
    const r = await t.handler(parsed.data!, {});
    return { err: !!r.isError, sc: r.structuredContent ?? {}, text: r.content?.[0]?.text ?? "" };
  };
  /** Предпросмотр → подтверждение; «Код/Номер не уникально» (нумератор копии отстаёт) — новый предпросмотр. */
  const apply = async (name: string, a: Json): Promise<Result> => {
    for (let i = 0; i < 30; i++) {
      const p = await run(name, a);
      if (p.err) return p;
      const c = await run(name, {
        ...a,
        confirm: true,
        ...(p.sc["operationId"] ? { operationId: p.sc["operationId"] } : {}),
      });
      if (!c.err || !/не уникально/.test(c.text)) return c;
    }
    return { err: true, sc: {}, text: "30 попыток: номер или код занят" };
  };

  const rows: Array<[string, boolean, string]> = [];
  const check = (name: string, ok: boolean, detail = "") => rows.push([name, ok, detail]);
  const tag = `ZZ-TEST live ${new Date().toISOString().slice(0, 16)}`;
  const org = (
    await get(`Catalog_Организации?$format=json&$filter=Description eq '${q(orgName)}'&$select=Ref_Key`)
  )[0];
  if (!org) {
    line(`Организация «${orgName}» не найдена.`);
    return 2;
  }
  const orgRef = String(org["Ref_Key"]);
  const first = async (set: string, select: string, filter = "") =>
    (
      await get(
        `${set}?$format=json&$top=1&$orderby=Date desc&$filter=Posted eq true${filter}&$select=${select}`,
      ).catch(() => [] as Json[])
    )[0];

  // --- тестовые справочники: найти по наименованию или создать
  const find = async (set: string, name: string) =>
    (
      await get(
        `${set}?$format=json&$filter=Description eq '${q(name)}' and DeletionMark eq false&$select=Ref_Key`,
      )
    )[0]?.["Ref_Key"] as string | undefined;
  const ensure = async (label: string, set: string, name: string, tool: string, a: Json) => {
    const have = await find(set, name);
    if (have) return have;
    const r = await apply(tool, a);
    check(`справочник: ${label}`, !r.err, r.err ? r.text.slice(0, 160) : "");
    return r.err ? undefined : String(r.sc["ref"]);
  };
  const NAME = "ZZ-TEST live контрагент коннектора";
  const partner = await ensure("партнёр", "Catalog_Партнеры", NAME, "write.counterparty.create_partner", {
    name: NAME,
    relations: ["client", "supplier"],
  });
  const cp = partner
    ? await ensure("контрагент", "Catalog_Контрагенты", NAME, "write.counterparty.create_counterparty", {
        name: NAME,
        legalType: "ЮридическоеЛицо",
        partnerRef: partner,
      })
    : undefined;
  const contract = async (kind: string, name: string) =>
    cp
      ? ensure(name, "Catalog_ДоговорыКонтрагентов", name, "write.catalog.create_contract", {
          organization: orgName,
          counterpartyRef: cp,
          kind,
          name,
        })
      : undefined;
  const ctrClient = await contract("СПокупателем", "ZZ-TEST live договор с покупателем");
  const ctrSupplier = await contract("СПоставщиком", "ZZ-TEST live договор с поставщиком");
  const agreement = async (side: string, set: string, name: string) =>
    cp
      ? ensure(name, set, name, "write.catalog.create_agreement", {
          organization: orgName,
          side,
          counterpartyRef: cp,
          name,
        })
      : undefined;
  const agrClient = await agreement(
    "client",
    "Catalog_СоглашенияСКлиентами",
    "ZZ-TEST live соглашение с клиентом",
  );
  await agreement("supplier", "Catalog_СоглашенияСПоставщиками", "ZZ-TEST live соглашение с поставщиком");
  // Вид номенклатуры — как у позиции последнего приобретения (обычный товар этой базы).
  const lastPurchase = await first("Document_ПриобретениеТоваровУслуг", "Товары");
  const sampleItem = ((lastPurchase?.["Товары"] as Json[] | undefined) ?? [])[0]?.["Номенклатура_Key"];
  const kind = sampleItem
    ? (
        await conn.client.getEntity(
          `Catalog_Номенклатура(guid'${String(sampleItem)}')?$format=json&$select=ВидНоменклатуры_Key`,
        )
      )["ВидНоменклатуры_Key"]
    : undefined;
  const item = kind
    ? await ensure(
        "номенклатура",
        "Catalog_Номенклатура",
        "ZZ-TEST live товар коннектора",
        "write.catalog.create_nomenclature",
        {
          name: "ZZ-TEST live товар коннектора",
          kind,
        },
      )
    : undefined;
  if (!cp || !ctrClient || !ctrSupplier || !item) {
    for (const [name, ok, detail] of rows) line(`${ok ? "✓" : "✗"} ${name.padEnd(52)} ${detail}`);
    line("Тестовые справочники не готовы — документы не создаются.");
    return 1;
  }

  // --- документы: создать → провести со сверкой → (в конце) отменить → пометить
  const made: Array<[string, string, string]> = [];
  const doc = async (label: string, set: string, tool: string, a: Json) => {
    const r = await apply(tool, { comment: tag, ...a });
    if (r.err) return void check(`${label}: создание`, false, r.text.slice(0, 160));
    const ref = String(r.sc["ref"]);
    made.push([label, set, ref]);
    const p = await apply("write.document.post_document", { entitySet: set, ref, post: true });
    check(
      `${label}: проведение со сверкой`,
      !p.err && p.sc["verified"] !== false && p.sc["movementsVerified"] !== false,
      p.err ? p.text.slice(0, 160) : String(p.sc["movementsNote"] ?? "").slice(0, 100),
    );
    return ref;
  };
  const L = (n: number, price: number) => [{ nomenclatureRef: item, quantity: n, price }];
  const party = { organization: orgName, counterpartyRef: cp };
  const pur = await doc(
    "приобретение",
    "Document_ПриобретениеТоваровУслуг",
    "write.purchase.create_purchase",
    {
      ...party,
      contractRef: ctrSupplier,
      warehouse,
      lines: L(3, 10),
    },
  );
  const shipment = await doc("реализация", "Document_РеализацияТоваровУслуг", "write.sales.create_shipment", {
    ...party,
    contractRef: ctrClient,
    ...(agrClient ? { agreementRef: agrClient } : {}),
    warehouse,
    lines: L(1, 100),
  });
  if (pur)
    await doc(
      "возврат поставщику",
      "Document_ВозвратТоваровПоставщику",
      "write.warehouse.create_return_to_supplier",
      {
        ...party,
        contractRef: ctrSupplier,
        warehouse,
        basisRef: pur,
        lines: L(1, 10),
      },
    );
  const order = await doc("заказ клиента", "Document_ЗаказКлиента", "write.sales.create_customer_order", {
    ...party,
    contractRef: ctrClient,
    ...(agrClient ? { agreementRef: agrClient } : {}),
    warehouse,
    lines: L(1, 100),
  });
  const warehouse2 = arg("warehouse2");
  if (warehouse2)
    await doc("перемещение", "Document_ПеремещениеТоваров", "write.warehouse.create_transfer", {
      organization: orgName,
      fromWarehouse: warehouse,
      toWarehouse: warehouse2,
      lines: [{ nomenclatureRef: item, quantity: 1, price: 10 }],
    });
  const surplusSample = await first("Document_ОприходованиеИзлишковТоваров", "ВидЦены_Key,СтатьяДоходов_Key");
  if (surplusSample && !EMPTY.test(String(surplusSample["ВидЦены_Key"] ?? "0")))
    await doc(
      "оприходование излишков",
      "Document_ОприходованиеИзлишковТоваров",
      "write.warehouse.create_surplus",
      {
        organization: orgName,
        warehouse,
        priceKind: String(surplusSample["ВидЦены_Key"]),
        ...(surplusSample["СтатьяДоходов_Key"]
          ? { incomeItem: String(surplusSample["СтатьяДоходов_Key"]) }
          : {}),
        lines: L(1, 10),
      },
    );
  else
    check(
      "оприходование излишков",
      true,
      "пропущено: в базе нет проведённых оприходований — не из чего взять вид цен",
    );

  // Деньги: статьи ДДС — из последнего проведённого документа той же операции.
  const cashName =
    arg("cash") ??
    ((
      await get(
        `Catalog_Кассы?$format=json&$filter=Owner_Key eq guid'${orgRef}' and DeletionMark eq false&$select=Description`,
      )
    )[0]?.["Description"] as string | undefined);
  const bankName =
    arg("bank") ??
    ((
      await get(
        `Catalog_БанковскиеСчетаОрганизаций?$format=json&$filter=Owner_Key eq guid'${orgRef}' and DeletionMark eq false&$select=Description`,
      )
    )[0]?.["Description"] as string | undefined);
  const cf = async (set: string, op: string) => {
    const d = await first(
      set,
      "РасшифровкаПлатежа,СтатьяДвиженияДенежныхСредств_Key",
      ` and ХозяйственнаяОперация eq '${op}'`,
    );
    const r = ((d?.["РасшифровкаПлатежа"] as Json[] | undefined) ?? [])[0];
    const v = r?.["СтатьяДвиженияДенежныхСредств_Key"] ?? d?.["СтатьяДвиженияДенежныхСредств_Key"];
    return v && !EMPTY.test(String(v)) ? String(v) : undefined;
  };
  if (cashName) {
    const item = await cf("Document_ПриходныйКассовыйОрдер", "ПоступлениеОплатыОтКлиента");
    await doc("ПКО от клиента", "Document_ПриходныйКассовыйОрдер", "write.money.create_cash_receipt", {
      ...party,
      contractRef: ctrClient,
      operationKind: "ПоступлениеОплатыОтКлиента",
      amount: 10,
      cash: cashName,
      ...(item ? { cashflowItem: item } : {}),
    });
  } else check("ПКО от клиента", true, "пропущено: у организации нет кассы");
  if (bankName) {
    const item = await cf("Document_СписаниеБезналичныхДенежныхСредств", "ОплатаПоставщику");
    await doc(
      "списание безналичных поставщику",
      "Document_СписаниеБезналичныхДенежныхСредств",
      "write.money.create_bank_writeoff",
      {
        ...party,
        contractRef: ctrSupplier,
        operationKind: "ОплатаПоставщику",
        amount: 10,
        orgBankAccount: bankName,
        ...(item ? { cashflowItem: item } : {}),
      },
    );
    // Оплата несколькими строками (settlementLines): поставщику — по двум закупкам, от клиента — по реализации и заказу.
    // Проведение разносит расчёты по объектам расчётов документов-оснований.
    const byObject = async (set: string, ref: string, register: string, bases: string[]) => {
      const objs = await Promise.all(
        bases.map(async (b) => {
          for (const t of [
            "Document_ПриобретениеТоваровУслуг",
            "Document_РеализацияТоваровУслуг",
            "Document_ЗаказКлиента",
          ]) {
            const o = (
              await get(
                `Catalog_ОбъектыРасчетов?$format=json&$filter=Объект eq cast(guid'${b}', '${t}')&$select=Ref_Key`,
              )
            )[0];
            if (o) return String(o["Ref_Key"]);
          }
          return "";
        }),
      );
      const rs =
        (
          (await conn.client
            .getEntity(`${register}(Recorder='${ref}', Recorder_Type='StandardODATA.${set}')?$format=json`)
            .catch(() => ({ RecordSet: [] }))) as { RecordSet?: Json[] }
        ).RecordSet ?? [];
      return objs.map(
        (o) =>
          Math.round(
            rs.filter((r) => r["ОбъектРасчетов_Key"] === o).reduce((t, r) => t + Number(r["Сумма"] ?? 0), 0) *
              100,
          ) / 100,
      );
    };
    const pur2 = pur
      ? await doc("приобретение 2", "Document_ПриобретениеТоваровУслуг", "write.purchase.create_purchase", {
          ...party,
          contractRef: ctrSupplier,
          warehouse,
          lines: L(1, 20),
        })
      : undefined;
    if (pur && pur2) {
      const split = await doc(
        "оплата поставщику по двум закупкам",
        "Document_СписаниеБезналичныхДенежныхСредств",
        "write.money.create_bank_writeoff",
        {
          ...party,
          contractRef: ctrSupplier,
          operationKind: "ОплатаПоставщику",
          amount: 30,
          orgBankAccount: bankName,
          ...(item ? { cashflowItem: item } : {}),
          settlementLines: [
            { basisRef: pur, amount: 10 },
            { basisRef: pur2, amount: 20 },
          ],
        },
      );
      if (split) {
        const got = await byObject(
          "Document_СписаниеБезналичныхДенежныхСредств",
          split,
          "AccumulationRegister_РасчетыСПоставщиками",
          [pur, pur2],
        );
        check(
          "две строки поставщику: расчёты по каждой закупке",
          got[0] === 10 && got[1] === 20,
          JSON.stringify(got),
        );
      }
    }
    if (shipment && order) {
      const inItem = await cf("Document_ПоступлениеБезналичныхДенежныхСредств", "ПоступлениеОплатыОтКлиента");
      const split = await doc(
        "поступление от клиента по реализации и заказу",
        "Document_ПоступлениеБезналичныхДенежныхСредств",
        "write.money.create_payment",
        {
          ...party,
          contractRef: ctrClient,
          operationKind: "ПоступлениеОплатыОтКлиента",
          amount: 15,
          orgBankAccount: bankName,
          ...(inItem ? { cashflowItem: inItem } : {}),
          settlementLines: [
            { basisRef: shipment, amount: 10 },
            { basisRef: order, amount: 5 },
          ],
        },
      );
      if (split) {
        const got = await byObject(
          "Document_ПоступлениеБезналичныхДенежныхСредств",
          split,
          "AccumulationRegister_РасчетыСКлиентами",
          [shipment, order],
        );
        check(
          "две строки от клиента: расчёты по каждому документу",
          got[0] === 10 && got[1] === 5,
          JSON.stringify(got),
        );
      }
    }
    const reqItem = await cf("Document_ЗаявкаНаРасходованиеДенежныхСредств", "ОплатаПоставщику");
    await doc(
      "заявка на расходование",
      "Document_ЗаявкаНаРасходованиеДенежныхСредств",
      "write.money.create_payment_request",
      {
        ...party,
        contractRef: ctrSupplier,
        operationKind: "ОплатаПоставщику",
        payForm: "bank",
        amount: 10,
        orgBankAccount: bankName,
        ...(reqItem ? { cashflowItem: reqItem } : {}),
      },
    );
  } else check("деньги по банку", true, "пропущено: у организации нет банковского счёта");

  // --- отмена проведения (в обратном порядке) и пометка удаления
  for (const [label, set, ref] of [...made].reverse()) {
    const posted = await conn.client.getEntity(`${set}(guid'${ref}')?$format=json&$select=Posted`);
    if (posted["Posted"] === true) {
      const u = await apply("write.document.post_document", { entitySet: set, ref, post: false });
      check(
        `${label}: отмена проведения, движений нет`,
        !u.err && u.sc["movementsVerified"] !== false,
        u.err ? u.text.slice(0, 120) : "",
      );
    }
    const m = await run("write.entity.mark_for_deletion", { entitySet: set, ref, mark: true, confirm: true });
    check(`${label}: пометка удаления`, !m.err, m.err ? m.text.slice(0, 120) : "");
  }

  line();
  for (const [name, ok, detail] of rows) line(`${ok ? "✓" : "✗"} ${name.padEnd(52)} ${detail}`);
  const failed = rows.filter(([, ok]) => !ok).length;
  line(`\nИтог: ${rows.length - failed} из ${rows.length} прошли.`);
  return failed === 0 ? 0 : 1;
}

main()
  .then((code) => process.exit(code))
  .catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? (e.stack ?? e.message) : String(e)}\n`);
    process.exit(1);
  });
