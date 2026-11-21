/**
 * Живой прогон записи против НАСТОЯЩЕЙ базы 1С: все инструменты создания (предпросмотр →
 * подтверждение → повтор с тем же operationId), правка строк документа, потерянный ответ
 * (создание и строки) со сверкой через write.operation.status. Все созданные объекты в конце
 * помечаются на удаление (физически не удаляются — это делается в 1С «Удалением помеченных»).
 *
 * Пишет в боевую базу, поэтому запускается только явно:
 *   npm run test:live -- --db <имя базы> --yes [--slow] [--no-post]
 * --no-post — без проведения (рабочая база, где проводить не разрешено): сценарий «договор → акт проводится» пропущен.
 * Документы получают комментарий с меткой ZZ-TEST, если инструмент его принимает.
 * --slow добавляет сценарий «правка не дошла» (ждёт 2 минуты; статус сам операцию не закрывает,
 * закрытие — только явным markNotApplied=true, #52).
 * Журнал операций — во временном каталоге, рабочий журнал не трогается.
 * Код возврата 1 — если хоть один сценарий не прошёл.
 */
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

process.env["ODATA_WRITE_JOURNAL_DIR"] = mkdtempSync(join(tmpdir(), "1c-odata-live-"));

type Json = Record<string, unknown>;
type Result = { err: boolean; sc: Json; text: string };
type Handler = (
  args: Json,
  extra: Json,
) => Promise<{
  isError?: boolean;
  structuredContent?: Json;
  content?: Array<{ type: string; text?: string }>;
}>;
type Tool = {
  inputSchema: { safeParse: (v: unknown) => { success: boolean; data?: Json; error?: unknown } };
  outputSchema?: { safeParse: (v: unknown) => { success: boolean; error?: { issues?: unknown } } };
  handler: Handler;
};

function arg(name: string): string | undefined {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? process.argv[i + 1] : undefined;
}
const has = (name: string) => process.argv.includes(`--${name}`);
const line = (s = "") => process.stdout.write(`${s}\n`);

async function main(): Promise<number> {
  const db = arg("db");
  if (!db || !has("yes")) {
    line("Живой прогон пишет в базу 1С. Запуск: npm run test:live -- --db <имя базы> --yes [--slow]");
    return 2;
  }
  const { loadConfig } = await import("../config/env.js");
  const { ServerContext } = await import("../context.js");
  const { createServer } = await import("../mcp/server.js");
  const { ODataError } = await import("../odata/errors.js");
  const cfg = loadConfig();
  const ctx = new ServerContext(cfg);
  const conn = ctx.db(db);
  if (cfg.behavior.readOnly || !conn.cfg.writable) {
    line(`База "${db}" недоступна для записи: нужны READ_ONLY=false и WRITABLE для базы.`);
    return 2;
  }
  const tools = (createServer(ctx) as unknown as { _registeredTools: Record<string, Tool> })._registeredTools;
  const client = conn.client as unknown as {
    request: (path: string, method?: string, body?: unknown) => Promise<unknown>;
    getEntity: (path: string) => Promise<Json>;
    getCollection: (path: string) => Promise<{ value: Json[] }>;
  };

  const run = async (name: string, a: Json): Promise<Result> => {
    const tool = tools[name];
    if (!tool) return { err: true, sc: {}, text: `нет инструмента ${name}` };
    const parsed = tool.inputSchema.safeParse({ database: db, ...a });
    if (!parsed.success)
      return { err: true, sc: {}, text: `входные данные: ${JSON.stringify(parsed.error)}` };
    const r = await tool.handler(parsed.data!, {});
    // Настоящий MCP-клиент отвергает ответ, не прошедший outputSchema; прямой вызов обработчика — нет.
    if (!r.isError && tool.outputSchema) {
      const out = tool.outputSchema.safeParse(r.structuredContent);
      if (!out.success)
        return {
          err: true,
          sc: {},
          text: `ответ не проходит outputSchema: ${JSON.stringify(out.error?.issues).slice(0, 300)}`,
        };
    }
    return { err: !!r.isError, sc: r.structuredContent ?? {}, text: r.content?.[0]?.text ?? "" };
  };

  const made: Array<[string, string]> = [];
  const rows: Array<[string, boolean, string]> = [];
  const check = (name: string, ok: boolean, detail = "") => rows.push([name, ok, detail]);

  /** Комментарий с меткой — документ находится в 1С поиском по «ZZ-TEST», даже если прогон оборвался. */
  const tagged = (name: string, a: Json): Json => {
    const shape = (tools[name]?.inputSchema as unknown as { shape?: Record<string, unknown> }).shape;
    return shape && "comment" in shape && a["comment"] === undefined ? { ...a, comment: tag } : a;
  };

  /** Предпросмотр → подтверждение → повтор: одна запись, повтор отвечает из журнала той же ссылкой. */
  async function flow(name: string, args: Json, label = name): Promise<Json | undefined> {
    let a = tagged(name, args);
    let p = await run(name, { ...a, confirm: false });
    // Часть инструментов принимает комментарий только в Управлении торговлей — тогда без него.
    if (p.err && a !== args && p.text.includes("только для Управления торговлей")) {
      a = args;
      p = await run(name, { ...a, confirm: false });
    }
    if (p.err || typeof p.sc["operationId"] !== "string")
      return void check(label, false, `предпросмотр: ${p.text.slice(0, 160)}`);
    const operationId = p.sc["operationId"];
    const c1 = await run(name, { ...a, confirm: true, operationId });
    if (c1.err) return void check(label, false, `подтверждение: ${c1.text.slice(0, 160)}`);
    const ref = String(c1.sc["ref"]);
    made.push([String(c1.sc["entitySet"]), ref]);
    const c2 = await run(name, { ...a, confirm: true, operationId });
    const replayOk = c2.sc["replayed"] === true && c2.sc["ref"] === ref;
    const v1 = /^[0-9a-f]{8}-[0-9a-f]{4}-1[0-9a-f]{3}-/i.test(ref);
    check(
      label,
      replayOk && v1,
      `${c1.sc["entitySet"]} ${ref.slice(0, 8)}${replayOk ? "" : " ПОВТОР НЕ ИЗ ЖУРНАЛА"}${v1 ? "" : " Ref_Key не UUIDv1"}`,
    );
    return c1.sc;
  }

  /** Правка существующего объекта: предпросмотр → подтверждение → повтор из журнала. */
  async function edit(name: string, a: Json, label = name): Promise<Json | undefined> {
    const p = await run(name, { ...a, confirm: false });
    if (p.err) return void check(label, false, `предпросмотр: ${p.text.slice(0, 160)}`);
    const op = typeof p.sc["operationId"] === "string" ? { operationId: p.sc["operationId"] } : {};
    const c1 = await run(name, { ...a, confirm: true, ...op });
    if (c1.err) return void check(label, false, `подтверждение: ${c1.text.slice(0, 160)}`);
    const c2 = await run(name, { ...a, confirm: true, ...op });
    const replayOk = !c2.err && (c2.sc["replayed"] === undefined || c2.sc["replayed"] === true);
    check(label, replayOk, replayOk ? "" : `повтор: ${c2.err ? c2.text.slice(0, 120) : "не из журнала"}`);
    return c1.sc;
  }

  /** Инструмент другой конфигурации: предпросмотр обязан отказать, не дойдя до записи. */
  async function refused(name: string, a: Json): Promise<void> {
    const p = await run(name, { ...a, confirm: false });
    check(
      `${name} — отказ в этой конфигурации`,
      p.err && !p.text.startsWith("входные данные"),
      p.text.slice(0, 100),
    );
  }

  const dataVersion = async (set: string, ref: string) =>
    String((await client.getEntity(`${set}(guid'${ref}')?$format=json&$select=DataVersion`))["DataVersion"]);

  /** Подменяет HTTP-запросы метода: after — ответ 1С «потерян» после отправки, before — запрос не ушёл. */
  const origRequest = client.request.bind(client);
  const lose = (method: string, when: "after" | "before") => {
    client.request = async (path, m, body) => {
      if (m === method && when === "before")
        throw new ODataError({ kind: "timeout", message: "имитация: запрос не ушёл" });
      const r = await origRequest(path, m, body);
      if (m === method) throw new ODataError({ kind: "timeout", message: "имитация: ответ потерян" });
      return r;
    };
  };
  const restore = () => {
    client.request = origRequest;
  };
  const first = async (set: string, filter = "") =>
    (await client.getCollection(`${set}?$format=json&$top=1&$select=Ref_Key,Description${filter}`)).value[0];

  const tag = `ZZ-TEST live ${new Date().toISOString().slice(0, 16)} (удалить)`;
  line(`Живой прогон записи: база "${db}", метка объектов «${tag}»`);

  // --- фикстуры и все инструменты создания
  const cp = await flow("write.counterparty.create_counterparty", { name: `${tag} контрагент` });
  const goods = await flow("write.catalog.create_nomenclature", { name: `${tag} товар` });
  let pay: Json | undefined;
  let ctrCustomer: Json | undefined;
  if (!cp || !goods) {
    line("Не удалось создать контрагента/номенклатуру — дальше не идём.");
  } else {
    const cpRef = String(cp["ref"]);
    const gRef = String(goods["ref"]);
    const L = [{ nomenclatureRef: gRef, quantity: 1, price: 100, vatRate: "БезНДС" }];
    const W = [{ nomenclatureRef: gRef, quantity: 1, price: 100 }];
    const wh = await first("Catalog_Склады", "&$filter=IsFolder%20eq%20false").catch(() => undefined);
    const grp = await first("Catalog_НоменклатурныеГруппы").catch(() => undefined);

    ctrCustomer = await flow("write.catalog.create_contract", {
      counterpartyRef: cpRef,
      kind: "СПокупателем",
    });
    await flow(
      "write.catalog.create_contract",
      { counterpartyRef: cpRef, kind: "СПоставщиком" },
      "write.catalog.create_contract (поставщик)",
    );
    const inv = await flow("write.sales.create_invoice", { counterpartyRef: cpRef, lines: L });
    const pur = await flow("write.purchase.create_purchase", { counterpartyRef: cpRef, lines: L });
    await flow("write.purchase.create_supplier_invoice", { counterpartyRef: cpRef, lines: L });
    await flow("write.money.create_payout_order", {
      counterpartyRef: cpRef,
      amount: 1,
      number: `ZZ${Date.now() % 1_000_000}`,
      purposeText: "тест",
      recipientText: "тест",
    });
    const ship = await flow("write.sales.create_shipment", { counterpartyRef: cpRef, lines: L });
    await flow("write.warehouse.create_return_from_customer", { counterpartyRef: cpRef, lines: L });
    await flow("write.warehouse.create_return_to_supplier", { counterpartyRef: cpRef, lines: L });
    if (wh) {
      const name = String(wh["Description"]);
      await flow("write.warehouse.create_transfer", { fromWarehouse: name, toWarehouse: name, lines: W });
    } else check("write.warehouse.create_transfer", false, "нет склада в базе");
    await flow("write.warehouse.create_surplus", { lines: W });
    await flow("write.warehouse.create_writeoff", { lines: W });
    await flow("write.warehouse.create_inventory", {
      lines: [{ nomenclatureRef: gRef, accountingQuantity: 0, factQuantity: 1, price: 100 }],
    });
    await flow("write.sales.create_act", { counterpartyRef: cpRef, lines: L });
    if (grp) {
      await flow("write.sales.create_services_act", {
        counterpartyRef: cpRef,
        nomenclatureGroupRef: String(grp["Ref_Key"]),
        lines: [
          { name: "Тестовая услуга", quantity: 1, price: 100, nomenclatureRef: gRef, vatRate: "БезНДС" },
        ],
      });
    } else check("write.sales.create_services_act", false, "нет номенклатурной группы в базе");
    pay = ctrCustomer
      ? await flow("write.money.create_payment", {
          counterpartyRef: cpRef,
          contractRef: String(ctrCustomer["ref"]),
          amount: 1,
        })
      : undefined;
    await flow("write.money.create_bank_writeoff", { operationKind: "ПрочееСписание", amount: 1 });
    await flow("write.money.create_cash_receipt", { operationKind: "ПрочийПриход", amount: 1 });
    await flow("write.money.create_cash_payment", { operationKind: "ПрочийРасход", amount: 1 });
    if (ship) await flow("write.sales.create_issued_invoice", { baseDocumentRef: String(ship["ref"]) });
    if (pur) await flow("write.purchase.create_received_invoice", { baseDocumentRef: String(pur["ref"]) });
    await flow("write.counterparty.create_bank_account", {
      ownerRef: cpRef,
      accountNumber: "40702810000000000001",
      bik: "044525225",
    });
    await flow("write.counterparty.create_contact_person", { ownerRef: cpRef, name: `${tag} лицо` });
    await flow("write.entity.create_folder", { name: `${tag} папка` });
    if (inv)
      await flow("write.document.copy_document", {
        entitySet: String(inv["entitySet"]),
        ref: String(inv["ref"]),
      });

    // --- создание: ответ потерян → повтор заблокирован → статус по Ref_Key → повтор из журнала
    {
      const a = { name: `${tag} потерянный ответ` };
      const p = await run("write.counterparty.create_counterparty", { ...a, confirm: false });
      const operationId = p.sc["operationId"];
      lose("POST", "after");
      const c1 = await run("write.counterparty.create_counterparty", { ...a, confirm: true, operationId });
      restore();
      const blocked = (
        await run("write.counterparty.create_counterparty", { ...a, confirm: true, operationId })
      ).err;
      const st = await run("write.operation.status", { operationId });
      const c3 = await run("write.counterparty.create_counterparty", { ...a, confirm: true, operationId });
      if (typeof st.sc["ref"] === "string") made.push(["Catalog_Контрагенты", st.sc["ref"]]);
      const found = await client.getCollection(
        `Catalog_Контрагенты?$format=json&$select=Ref_Key&$filter=${encodeURIComponent(`Description eq '${a.name}'`)}`,
      );
      check(
        "создание: потерянный ответ → сверка по Ref_Key",
        c1.err &&
          blocked &&
          st.sc["status"] === "found_reconciled" &&
          c3.sc["replayed"] === true &&
          found.value.length === 1,
        `статус ${String(st.sc["status"])}, в 1С объектов: ${found.value.length}`,
      );
    }

    // --- договор по умолчанию и проведение (без договора БП 3.0 не проводит реализацию)
    if (has("no-post")) line("  проведение пропущено (--no-post)");
    else {
      const postAndUnpost = async (entitySet: string, ref: string) => {
        const p = await run("write.document.post_document", { entitySet, ref, post: true, confirm: true });
        if (!p.err) await run("write.document.post_document", { entitySet, ref, post: false, confirm: true });
        return p;
      };
      const svcLine = [
        { nomenclatureRef: gRef, quantity: 1, price: 100, vatRate: "БезНДС", content: "Тест" },
      ];
      const withOne = await flow("write.counterparty.create_counterparty", { name: `${tag} один договор` });
      if (withOne) {
        await flow("write.catalog.create_contract", {
          counterpartyRef: String(withOne["ref"]),
          kind: "СПокупателем",
        });
        const preview = await run("write.sales.create_act", {
          counterpartyRef: String(withOne["ref"]),
          lines: svcLine,
          confirm: false,
        });
        const notes = JSON.stringify(preview.sc["notes"] ?? []);
        const act = await flow(
          "write.sales.create_act",
          { counterpartyRef: String(withOne["ref"]), lines: svcLine },
          "акт без contractRef (договор подставлен)",
        );
        const posted = act ? await postAndUnpost(String(act["entitySet"]), String(act["ref"])) : undefined;
        check(
          "договор подставлен → акт проводится",
          notes.includes("подставлен единственный договор") && posted?.err === false,
          posted?.err ? posted.text.slice(0, 120) : "",
        );
      }
      const noContract = await flow("write.counterparty.create_counterparty", {
        name: `${tag} без договора`,
      });
      if (noContract) {
        const act = await flow(
          "write.sales.create_act",
          { counterpartyRef: String(noContract["ref"]), lines: svcLine },
          "акт без договора (создаётся)",
        );
        const posted = act ? await postAndUnpost(String(act["entitySet"]), String(act["ref"])) : undefined;
        check(
          "без договора → понятная ошибка проведения",
          posted?.err === true && posted.text.includes("не заполнен договор"),
          posted?.text.slice(0, 80) ?? "",
        );
      }
    }

    // --- строки документа
    if (inv) {
      const DOC = String(inv["entitySet"]);
      const ref = String(inv["ref"]);
      const count = async () =>
        ((await client.getEntity(`${DOC}(guid'${ref}')?$format=json&$select=Товары`))["Товары"] as unknown[])
          .length;
      const line1 = { nomenclatureRef: gRef, quantity: 2, price: 50, vatRate: "БезНДС" };
      const n0 = await count();

      const add = { entitySet: DOC, ref, line: line1 };
      let p = await run("write.document.add_document_line", { ...add, confirm: false });
      let operationId = p.sc["operationId"];
      await run("write.document.add_document_line", { ...add, confirm: true, operationId });
      const r2 = await run("write.document.add_document_line", { ...add, confirm: true, operationId });
      const n1 = await count();
      check(
        "строка: добавление + повтор",
        r2.sc["replayed"] === true && n1 === n0 + 1,
        `строк ${n0} → ${n1}`,
      );

      p = await run("write.document.add_document_line", { ...add, confirm: false });
      operationId = p.sc["operationId"];
      lose("PATCH", "after");
      await run("write.document.add_document_line", { ...add, confirm: true, operationId });
      restore();
      const blocked = (await run("write.document.add_document_line", { ...add, confirm: true, operationId }))
        .err;
      const st = await run("write.operation.status", { operationId });
      const r3 = await run("write.document.add_document_line", { ...add, confirm: true, operationId });
      const n2 = await count();
      check(
        "строка: потерянный ответ → сверка по числу строк",
        blocked && st.sc["status"] === "found_reconciled" && r3.sc["replayed"] === true && n2 === n1 + 1,
        `статус ${String(st.sc["status"])}, строк ${n1} → ${n2}`,
      );

      const rem = { entitySet: DOC, ref, lineNumber: 1 };
      p = await run("write.document.remove_document_line", { ...rem, confirm: false });
      operationId = p.sc["operationId"];
      lose("PATCH", "after");
      await run("write.document.remove_document_line", { ...rem, confirm: true, operationId });
      restore();
      const blockedRem = (
        await run("write.document.remove_document_line", { ...rem, confirm: true, operationId })
      ).err;
      const stRem = await run("write.operation.status", { operationId });
      const n3 = await count();
      check(
        "строка: удаление, потерянный ответ",
        blockedRem && stRem.sc["status"] === "found_reconciled" && n3 === n2 - 1,
        `статус ${String(stRem.sc["status"])}, строк ${n2} → ${n3}`,
      );

      if (has("slow")) {
        p = await run("write.document.remove_document_line", { ...rem, confirm: false });
        operationId = p.sc["operationId"];
        lose("PATCH", "before");
        await run("write.document.remove_document_line", { ...rem, confirm: true, operationId });
        restore();
        const early = await run("write.operation.status", { operationId });
        // Раньше 2 минут не закрывается даже по ручному подтверждению.
        const earlyMark = await run("write.operation.status", { operationId, markNotApplied: true });
        line("  … ждём 2 минуты для сценария «правка не дошла»");
        await new Promise((r) => setTimeout(r, 125_000));
        // Без ручного подтверждения время само ничего не закрывает (#52).
        const late = await run("write.operation.status", { operationId });
        const marked = await run("write.operation.status", { operationId, markNotApplied: true });
        const again = await run("write.document.remove_document_line", {
          ...rem,
          confirm: true,
          operationId,
        });
        check(
          "строка: правка не дошла → not_applied только по markNotApplied",
          early.sc["status"] === "not_found" &&
            earlyMark.sc["status"] === "not_found" &&
            late.sc["status"] === "not_found" &&
            marked.sc["status"] === "not_applied" &&
            again.err &&
            (await count()) === n3,
          `статус ${String(early.sc["status"])} → ${String(late.sc["status"])} → ${String(marked.sc["status"])}`,
        );
      }
    }
  }

  // --- правка справочников и документов, цены, банковские сервисные правки
  if (cp && goods) {
    const cpRef = String(cp["ref"]);
    const gRef = String(goods["ref"]);
    const EMPTY = "00000000-0000-0000-0000-000000000000";
    // Группа отвечает записью 1С (record), а не ref.
    const group = await edit("write.counterparty.create_group", { name: `${tag} группа`, parentRef: EMPTY });
    const groupRef = (group?.["record"] as Json | undefined)?.["Ref_Key"];
    if (typeof groupRef === "string") made.push(["Catalog_Контрагенты", groupRef]);
    if (typeof groupRef === "string")
      await edit("write.counterparty.set_group", {
        ref: cpRef,
        subjectKind: "counterparty",
        parentRef: groupRef,
        expectedDataVersion: await dataVersion("Catalog_Контрагенты", cpRef),
      });
    await edit("write.counterparty.update_counterparty", {
      ref: cpRef,
      fullName: `${tag} полное наименование`,
    });
    await edit("write.catalog.update_nomenclature", { ref: gRef, article: "ZZ-1" });
    const folder = await flow(
      "write.entity.create_folder",
      { entitySet: "Catalog_Номенклатура", name: `${tag} папка товаров` },
      "write.entity.create_folder (номенклатура)",
    );
    if (folder)
      await edit("write.entity.move_to_folder", {
        entitySet: "Catalog_Номенклатура",
        ref: gRef,
        folder: String(folder["ref"]),
      });
    const doc = await flow(
      "write.sales.create_invoice",
      {
        counterpartyRef: cpRef,
        lines: [{ nomenclatureRef: gRef, quantity: 1, price: 100, vatRate: "БезНДС" }],
      },
      "write.sales.create_invoice (для правки)",
    );
    if (doc) {
      const set = String(doc["entitySet"]);
      const ref = String(doc["ref"]);
      await edit("write.entity.update_entity", {
        entitySet: set,
        ref,
        fields: { Комментарий: `${tag} правка` },
      });
      await edit("write.document.update_document_lines", {
        entitySet: set,
        ref,
        lines: [{ nomenclatureRef: gRef, quantity: 3, price: 10, vatRate: "БезНДС" }],
      });
    }
    const kinds = ["Catalog_ТипыЦенНоменклатуры", "Catalog_ВидыЦен"];
    let kind: Json | undefined;
    for (const k of kinds)
      kind ??= await first(k, "&$filter=DeletionMark%20eq%20false").catch(() => undefined);
    if (kind)
      await flow("write.pricing.create_price_setting", {
        priceKind: String(kind["Ref_Key"]),
        lines: [{ itemRef: gRef, price: 100 }],
      });
    else check("write.pricing.create_price_setting", false, "нет вида цен в базе");
    if (pay && ctrCustomer) {
      const set = String(pay["entitySet"]);
      const ref = String(pay["ref"]);
      await edit("write.document.apply_bank_settlement_settings", {
        entitySet: set,
        ref,
        contractRef: String(ctrCustomer["ref"]),
        side: "customer",
        expectedDataVersion: await dataVersion(set, ref),
      });
      // Сумма в шапке изменилась, строка — нет: ровно тот случай, ради которого выравнивание.
      await edit(
        "write.entity.update_entity",
        { entitySet: set, ref, fields: { СуммаДокумента: 2 } },
        "write.entity.update_entity (сумма банковского документа)",
      );
      await edit("write.document.align_bank_payment_amounts", {
        entitySet: set,
        ref,
        expectedDataVersion: await dataVersion(set, ref),
      });
    }
    // Инструменты только Управления торговлей: в Бухгалтерии — отказ до записи.
    const accounting = (await conn.getMetadata()).entities.has("AccountingRegister_Хозрасчетный");
    if (accounting) {
      const L = [{ nomenclatureRef: gRef, quantity: 1, price: 100 }];
      await refused("write.counterparty.create_partner", { name: `${tag} партнёр`, relations: ["client"] });
      await refused("write.catalog.create_agreement", {
        side: "client",
        counterpartyRef: cpRef,
        name: `${tag} соглашение`,
      });
      await refused("write.money.create_payment_request", {
        operationKind: "ОплатаПоставщику",
        amount: 1,
        payForm: "bank",
        counterpartyRef: cpRef,
      });
      await refused("write.sales.create_customer_order", { counterpartyRef: cpRef, lines: L });
      await refused("write.purchase.create_supplier_order", { counterpartyRef: cpRef, lines: L });
    }
  }

  // --- уборка: сначала документы, потом справочники
  made.sort(([a], [b]) => Number(b.startsWith("Document_")) - Number(a.startsWith("Document_")));
  let cleaned = 0;
  for (const [entitySet, ref] of made) {
    const r = await run("write.entity.mark_for_deletion", { entitySet, ref, mark: true, confirm: true });
    if (!r.err) cleaned++;
  }

  line();
  for (const [name, ok, detail] of rows) line(`${ok ? "✓" : "✗"} ${name.padEnd(52)} ${detail}`);
  const failed = rows.filter(([, ok]) => !ok).length;
  line();
  line(
    `Итог: ${rows.length - failed} из ${rows.length} прошли; помечено на удаление ${cleaned} из ${made.length}.`,
  );
  return failed === 0 && cleaned === made.length ? 0 : 1;
}

main()
  .then((code) => process.exit(code))
  .catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? (e.stack ?? e.message) : String(e)}\n`);
    process.exit(1);
  });
