import { z } from "zod";
import type { Connection } from "../context.js";
import type { EntityMeta, ODataEntity } from "../types/odata.js";
import { InputError } from "../errors.js";
import { CATALOGS, DOCUMENTS, resolveEntity } from "../config/mapping.js";
import { accountingChart } from "../config/accounting.js";
import { counterpartyCatalog, EMPTY_GUID, exactRef, readCounterparty } from "./counterparty-hierarchy.js";
import { settlementAccountsFor, settlementSideOf } from "./settlement-accounts.js";
import { odataGuid } from "./query.js";
import { ODataError } from "./errors.js";
import { fingerprintWriteInput } from "./write-operation-context.js";

export const expectedDocumentVersion = z
  .string()
  .min(1)
  .max(200)
  .refine((v) => v !== "*" && !/[\r\n]/.test(v));
export const bankSettlementInput = {
  entitySet: z.string(),
  ref: exactRef,
  contractRef: exactRef,
  counterpartyRef: exactRef
    .optional()
    .describe(
      "Явно выбранный контрагент: заполняет только нулевую ссылку документа. Заполненный контрагент должен совпадать; замена запрещена. Без выбора пустой контрагент блокирует действие.",
    ),
  side: z.enum(["supplier", "customer"]),
  expectedDataVersion: expectedDocumentVersion,
};
const CONTRACT = "ДоговорКонтрагента_Key",
  ACCOUNT = "СчетУчетаРасчетовСКонтрагентом_Key",
  ADVANCE = "СчетУчетаРасчетовПоАвансам_Key",
  ROWS = "РасшифровкаПлатежа";
const bankSets: ReadonlySet<string> = new Set([...DOCUMENTS.bankIn, ...DOCUMENTS.bankOut]);
const same = (a: unknown, b: unknown) =>
  fingerprintWriteInput("compare", { a }) === fingerprintWriteInput("compare", { a: b });
function requireTypes(meta: EntityMeta | undefined, fields: Record<string, string>): void {
  const props = new Map(meta?.properties.map((p) => [p.name, p.type]));
  if (!meta || Object.entries(fields).some(([k, t]) => props.get(k) !== t))
    throw new InputError("Не опубликованы совместимые поля документа/ссылки; действие запрещено.");
}
function published(record: ODataEntity, meta: EntityMeta): ODataEntity {
  return Object.fromEntries(
    meta.properties
      .filter((p) => p.name in record && p.name !== "DataVersion")
      .map((p) => [p.name, record[p.name]]),
  );
}
export async function readVersionedDocument(
  conn: Connection,
  entitySet: string,
  ref: string,
  version: string,
): Promise<{ doc: ODataEntity; meta: EntityMeta; path: string }> {
  exactRef.parse(ref);
  expectedDocumentVersion.parse(version);
  const meta = (await conn.getMetadata()).entities.get(entitySet);
  if (!meta || meta.class !== "document" || meta.keys.length !== 1 || meta.keys[0] !== "Ref_Key")
    throw new InputError("Нужен точный опубликованный документ с ключом Ref_Key.");
  requireTypes(meta, {
    Ref_Key: "Edm.Guid",
    DataVersion: "Edm.String",
    Posted: "Edm.Boolean",
    DeletionMark: "Edm.Boolean",
  });
  const path = `${entitySet}(${odataGuid(ref)})?$format=json`;
  const doc = await conn.client.getEntity(path);
  if (
    String(doc.Ref_Key).toLowerCase() !== ref.toLowerCase() ||
    doc.DeletionMark !== false ||
    typeof doc.Posted !== "boolean" ||
    doc.DataVersion !== version
  )
    throw new InputError("Документ удалён, состояние неизвестно или expectedDataVersion устарел.");
  return { doc, meta: meta!, path };
}
async function verifyAfter(
  send: () => Promise<unknown>,
  read: () => Promise<ODataEntity>,
  expected: ODataEntity,
  project: (record: ODataEntity) => ODataEntity = (record) => record,
): Promise<ODataEntity> {
  await send();
  try {
    const actual = await read();
    const protectedActual = project(actual);
    for (const [field, value] of Object.entries(expected))
      if (!same(protectedActual[field], value)) throw new Error(`Поле ${field} не подтверждено`);
    if (typeof actual.DataVersion !== "string" || !actual.DataVersion)
      throw new Error("Версия результата неизвестна");
    return actual;
  } catch (cause) {
    throw new ODataError({
      kind: "unknown",
      message:
        "Запрос отправлен, контрольное чтение не подтвердило результат. Повтор запрещён; требуется ручная сверка.",
      cause,
    });
  }
}
// Contract-dependent rate/settlement totals may be recalculated by 1C. Payment amounts,
// tax, cash-flow and expense analytics remain protected, including every other row field.
const recalculatedRowFields = new Set([
  "КурсВзаиморасчетов",
  "КратностьВзаиморасчетов",
  "СуммаВзаиморасчетов",
]);
function protectedSettlementResult(record: ODataEntity): ODataEntity {
  // No header recalculation is established: preserve every observed published business field.
  // The caller projects metadata first; DataVersion alone is expected to change.
  const result = Object.fromEntries(Object.entries(record).filter(([field]) => field !== "DataVersion"));
  if (Array.isArray(record[ROWS]))
    result[ROWS] = (record[ROWS] as ODataEntity[]).map((row) =>
      Object.fromEntries(Object.entries(row).filter(([field]) => !recalculatedRowFields.has(field))),
    );
  return result;
}
/** Applies the saved contract/account; explicit CP may fill only a zero GUID. Other fields stay intact. */
export async function applyBankSettlementSettings(
  conn: Connection,
  input: unknown,
  confirm: boolean,
): Promise<Record<string, unknown>> {
  const p = z.object(bankSettlementInput).parse(input);
  if (!bankSets.has(p.entitySet))
    throw new InputError("Поддержаны только существующие банковские поступления и списания.");
  const { doc, meta, path } = await readVersionedDocument(conn, p.entitySet, p.ref, p.expectedDataVersion);
  if (doc.Posted !== false)
    throw new InputError(
      "Документ проведён. Отмена проведения выполняется отдельно; автоматической отмены нет.",
    );
  requireTypes(meta, {
    [CONTRACT]: "Edm.Guid",
    [ACCOUNT]: "Edm.Guid",
    Организация_Key: "Edm.Guid",
    ВалютаДокумента_Key: "Edm.Guid",
    ВидОперации: "Edm.String",
  });
  const org = exactRef.parse(doc.Организация_Key),
    currency = exactRef.parse(doc.ВалютаДокумента_Key);
  // Resolve the published representation; never guess a composite reference's type.
  const props = new Map(meta.properties.map((field) => [field.name, field.type]));
  const cpField = props.get("Контрагент_Key") === "Edm.Guid" ? "Контрагент_Key" : "Контрагент";
  if (cpField === "Контрагент") {
    requireTypes(meta, { Контрагент: "Edm.String", Контрагент_Type: "Edm.String" });
    const cpSet = resolveEntity(CATALOGS.counterparties, new Set((await conn.getMetadata()).entities.keys()));
    if (!cpSet || doc.Контрагент_Type !== `StandardODATA.${cpSet}`)
      throw new InputError("Составная ссылка контрагента имеет неизвестный или неподдержанный тип.");
  }
  // Only an observed zero GUID is empty; missing, null or malformed values remain unknown.
  const fillCounterparty = doc[cpField] === EMPTY_GUID;
  const cp = fillCounterparty ? p.counterpartyRef : exactRef.parse(doc[cpField]);
  if (!cp)
    throw new InputError(
      "Контрагент документа пуст: требуется явный counterpartyRef; подбор по договору запрещён.",
    );
  if (p.counterpartyRef && p.counterpartyRef !== cp)
    throw new InputError("Замена заполненного контрагента документа запрещена.");
  const observedCounterparty = p.counterpartyRef
    ? await readCounterparty(conn, await counterpartyCatalog(conn), cp, false)
    : undefined;
  const md = await conn.getMetadata(),
    available = new Set(md.entities.keys());
  const contracts = resolveEntity(CATALOGS.contracts, available);
  if (!contracts) throw new InputError("Договоры не опубликованы.");
  requireTypes(md.entities.get(contracts), {
    Ref_Key: "Edm.Guid",
    Owner_Key: "Edm.Guid",
    Организация_Key: "Edm.Guid",
    ВалютаВзаиморасчетов_Key: "Edm.Guid",
    ВидДоговора: "Edm.String",
    РасчетыВУсловныхЕдиницах: "Edm.Boolean",
    IsFolder: "Edm.Boolean",
    DeletionMark: "Edm.Boolean",
  });
  const contract = await conn.client.getEntity(`${contracts}(${odataGuid(p.contractRef)})?$format=json`);
  if (
    String(contract.Ref_Key).toLowerCase() !== p.contractRef ||
    contract.DeletionMark !== false ||
    contract.IsFolder !== false ||
    String(contract.Owner_Key).toLowerCase() !== cp ||
    String(contract.Организация_Key).toLowerCase() !== org
  )
    throw new InputError(
      "Договор удалён, является группой или не совпадает по владельцу/организации/валюте документа.",
    );
  if (contract.РасчетыВУсловныхЕдиницах === true)
    throw new InputError("Расчёты в условных единицах этим инструментом не поддерживаются.");
  if (String(contract.ВалютаВзаиморасчетов_Key).toLowerCase() !== currency)
    throw new InputError("Валюта договора не совпадает с валютой документа.");
  const side = settlementSideOf(contract.ВидДоговора);
  if ((side && side !== p.side) || (!side && contract.ВидДоговора !== "Прочее"))
    throw new InputError("Сторона расчётов не соответствует виду договора.");
  const operation = String(doc.ВидОперации ?? "");
  const incoming = (DOCUMENTS.bankIn as readonly string[]).includes(p.entitySet);
  if (
    (incoming && ["ОплатаПоставщику", "ПрочееСписание"].includes(operation)) ||
    (!incoming && ["ОплатаПокупателя", "ПрочееПоступление"].includes(operation))
  )
    throw new InputError("Вид операции не соответствует направлению банковского документа.");
  if (
    !/^(ОплатаПоставщику|ОплатаПокупателя|РасчетыПоКредитамИЗаймам|ПрочееСписание|ПрочееПоступление)$/.test(
      operation,
    ) ||
    (/Поставщик/.test(operation) && p.side !== "supplier") ||
    (/Покупател/.test(operation) && p.side !== "customer")
  )
    throw new InputError("Вид операции не поддержан или требует другую явную сторону расчётов.");
  const resolution = await settlementAccountsFor(
    conn,
    { orgKey: org, contractRef: p.contractRef, side: p.side },
    { contract },
  );
  const accountRef = exactRef.parse(resolution.choice?.settle);
  const advanceRef = resolution.choice?.advance ? exactRef.parse(resolution.choice.advance) : undefined;
  const chart = resolveEntity(accountingChart(conn.cfg), available);
  if (!chart) throw new InputError("План счетов не опубликован.");
  requireTypes(md.entities.get(chart), {
    Ref_Key: "Edm.Guid",
    Code: "Edm.String",
    DeletionMark: "Edm.Boolean",
    ЗапретитьИспользоватьВПроводках: "Edm.Boolean",
  });
  const accounts = [];
  for (const ref of new Set([accountRef, ...(advanceRef ? [advanceRef] : [])])) {
    const account = await conn.client.getEntity(`${chart}(${odataGuid(ref)})?$format=json`);
    if (
      String(account.Ref_Key).toLowerCase() !== ref ||
      account.DeletionMark !== false ||
      account.IsFolder === true ||
      account.ЗапретитьИспользоватьВПроводках !== false ||
      typeof account.Code !== "string" ||
      !account.Code
    )
      throw new InputError("Выбранный счёт отсутствует, удалён или запрещён для проводок.");
    accounts.push(account);
  }
  const payload: ODataEntity = {
    [CONTRACT]: p.contractRef,
    [ACCOUNT]: accountRef,
    ...(fillCounterparty ? { [cpField]: cp } : {}),
    ...(advanceRef && props.get(ADVANCE) === "Edm.Guid" ? { [ADVANCE]: advanceRef } : {}),
  };
  if (!meta.properties.some((p) => p.name === ROWS))
    throw new InputError("Расшифровка платежа не опубликована; изменение запрещено.");
  if (!Array.isArray(doc[ROWS])) throw new InputError("Расшифровка платежа неизвестна; изменение запрещено.");
  const rows = doc[ROWS] as unknown[];
  if (rows.length > 1)
    throw new InputError("Несколько строк платежа: выбор строки неоднозначен, изменение запрещено.");
  if (rows.length) {
    const row = rows[0];
    const rowMeta = md.entities.get(`${p.entitySet}_${ROWS}`);
    requireTypes(rowMeta, { [CONTRACT]: "Edm.Guid", [ACCOUNT]: "Edm.Guid" });
    if (!row || typeof row !== "object" || Array.isArray(row))
      throw new InputError("Неверная строка расшифровки.");
    const r = row as ODataEntity;
    if (
      !same(r[CONTRACT], doc[CONTRACT]) ||
      Object.keys(r).some((k) => !rowMeta!.properties.some((p) => p.name === k))
    )
      throw new InputError(
        "Строка не соответствует договору шапки или содержит неопубликованные поля; изменение запрещено.",
      );
    const rowHasAdvance = rowMeta!.properties.some(
      (field) => field.name === ADVANCE && field.type === "Edm.Guid",
    );
    if (rowHasAdvance && !advanceRef && /^(ОплатаПокупателя|ОплатаПоставщику)$/.test(operation))
      throw new InputError(
        "Счёт авансов не найден: смена договора не должна сохранять счёт прежнего договора.",
      );
    payload[ROWS] = [
      {
        ...r,
        [CONTRACT]: p.contractRef,
        [ACCOUNT]: accountRef,
        ...(rowHasAdvance && advanceRef ? { [ADVANCE]: advanceRef } : {}),
      },
    ];
  }
  const warnings = paymentAmountWarnings(doc, rows);
  const headers = { "If-Match": p.expectedDataVersion };
  const prepared = {
    method: "PATCH",
    path,
    payload,
    headers,
    resolution,
    observedDocument: published(doc, meta),
    observedContract: contract,
    ...(observedCounterparty ? { observedCounterparty } : {}),
    account: accounts[0],
    accounts,
  };
  if (!confirm) {
    await conn.client.prepareCreate(p.entitySet, prepared, { withRefKey: false });
    return { dryRun: true, database: conn.cfg.name, ...prepared, ...(warnings.length ? { warnings } : {}) };
  }
  const result = await conn.client.writeRecordOnce(
    p.entitySet,
    prepared,
    () =>
      verifyAfter(
        () => conn.client.request(path, "PATCH", payload, headers),
        () => conn.client.getEntity(path),
        protectedSettlementResult({ ...published(doc, meta), ...payload }),
        protectedSettlementResult,
      ),
    { kind: "documentState", ref: p.ref },
  );
  return {
    updated: true,
    ...readbackResult(result),
    database: conn.cfg.name,
    entitySet: p.entitySet,
    ref: result.Ref_Key,
    dataVersion: result.DataVersion,
    resolution,
    ...(warnings.length ? { warnings } : {}),
  };
}

/**
 * Документ из контрольного чтения — сразу в ответе, без второго GET у вызывающего (#50), как у
 * версионного проведения. Повтор завершённой операции отвечает квитанцией журнала: это не новое
 * чтение 1С, поэтому verified=false и verificationScope=journal_receipt, документа в ответе нет.
 */
function readbackResult(result: ODataEntity): Record<string, unknown> {
  const replayed = result._operation_replayed === true;
  return {
    verified: !replayed,
    verificationScope: replayed ? "journal_receipt" : "document_state",
    readbackFetched: !replayed,
    ...(replayed ? { replayed: true } : { readback: result }),
  };
}

/**
 * Сумма платежа должна быть разнесена по расшифровке целиком: проводки банковского документа
 * формируются по строкам «Расшифровки платежа», и неразнесённая часть не попадёт в расчёты.
 * Форма 1С держит суммы согласованными сама, расходятся они при записи в обход формы (#61).
 * Инструмент суммы не меняет — только предупреждает до проведения.
 */
export function paymentAmountWarnings(doc: ODataEntity, rows: unknown[]): string[] {
  const problems = paymentAmountProblems(doc, rows);
  return problems.length
    ? [
        `${problems.join("; ")}. Проводки формируются по расшифровке — до проведения исправьте: ${ALIGN_HINT}.`,
      ]
    : [];
}

/** Как исправить суммы расшифровки штатно — для подсказок и отказа проведения. */
export const ALIGN_HINT =
  "write.document.align_bank_payment_amounts (одна строка, рубли) или в форме документа в 1С";

/**
 * Расхождения сумм «Расшифровки платежа» с документом, при которых 1С проведёт документ с неверными
 * проводками по расчётам или без них. Без строк (налоги, прочие операции без договора) — не проверяем:
 * сумма там в шапке, расшифровка пуста законно. Неизвестные (не числа) — не утверждаем.
 */
export function paymentAmountProblems(doc: ODataEntity, rows: unknown[]): string[] {
  const total = doc["СуммаДокумента"];
  if (typeof total !== "number" || rows.length === 0) return [];
  const list = rows as ODataEntity[];
  const problems: string[] = [];
  const amounts = list.map((r) => r["СуммаПлатежа"]);
  if (amounts.every((a) => typeof a === "number")) {
    const allocated = (amounts as number[]).reduce((a, b) => a + b, 0);
    if (Math.abs(allocated - total) >= 0.005)
      problems.push(
        `сумма документа ${total} не совпадает с суммой расшифровки платежа ${Math.round(allocated * 100) / 100}`,
      );
  }
  const emptySettlement = list.some(
    (r) =>
      typeof r[CONTRACT] === "string" &&
      r[CONTRACT] !== EMPTY_GUID &&
      "СуммаВзаиморасчетов" in r &&
      !r["СуммаВзаиморасчетов"] &&
      typeof r["СуммаПлатежа"] === "number" &&
      r["СуммаПлатежа"] !== 0,
  );
  if (emptySettlement) problems.push("в строке с договором не заполнена сумма взаиморасчётов");
  return problems;
}

/**
 * Перед проведением банковского документа: расхождения сумм расшифровки (см. paymentAmountProblems).
 * Не банковский документ — без запроса. Возвращает текст отказа или undefined.
 */
export async function bankPostingBlocker(
  conn: Connection,
  entitySet: string,
  ref: string,
): Promise<string | undefined> {
  if (!bankSets.has(entitySet)) return undefined;
  const props = new Set(
    ((await conn.getMetadata()).entities.get(entitySet)?.properties ?? []).map((p) => p.name),
  );
  if (!props.has(ROWS) || !props.has("СуммаДокумента")) return undefined;
  const doc = await conn.client.getEntity(
    `${entitySet}(${odataGuid(ref)})?$format=json&$select=СуммаДокумента,${ROWS}`,
  );
  const problems = paymentAmountProblems(doc, Array.isArray(doc[ROWS]) ? (doc[ROWS] as unknown[]) : []);
  if (!problems.length) return undefined;
  return (
    `Проведение не отправлено: ${problems.join("; ")}. 1С проведёт такой документ, но проводки по расчётам ` +
    `будут неверными или не сформируются. Исправьте суммы: ${ALIGN_HINT}, затем проведите снова.`
  );
}

export const alignBankPaymentInput = {
  entitySet: z.string(),
  ref: exactRef,
  expectedDataVersion: expectedDocumentVersion,
};

/** Процент НДС из значения перечисления: «НДС22», «НДС22_122» → 22; «БезНДС», «НДС0» → 0. */
function vatPercent(rate: unknown): number | undefined {
  if (rate === "БезНДС" || rate === "" || rate === undefined || rate === null) return 0;
  const m = /^НДС(\d+)(?:_\d+)?$/.exec(String(rate));
  return m ? Number(m[1]) : undefined;
}

/**
 * Выравнивает единственную строку «Расшифровки платежа» непроведённого рублёвого документа по сумме
 * документа (#61): сумма платежа и сумма взаиморасчётов = сумма документа, курс и кратность 1, НДС по
 * ставке строки (сумма платежа включает НДС: сумма × ставка / (100 + ставка), как считает форма).
 * Ничего не угадывает: несколько строк, валюта, условные единицы, неизвестная ставка — отказ.
 * Остальные поля документа и строки сохраняются; версия, журнал и контрольное чтение — как у
 * apply_bank_settlement_settings. Проведение — отдельным вызовом.
 */
export async function alignBankPaymentAmounts(
  conn: Connection,
  input: unknown,
  confirm: boolean,
): Promise<Record<string, unknown>> {
  const p = z.object(alignBankPaymentInput).parse(input);
  if (!bankSets.has(p.entitySet))
    throw new InputError("Поддержаны только существующие банковские поступления и списания.");
  const { doc, meta, path } = await readVersionedDocument(conn, p.entitySet, p.ref, p.expectedDataVersion);
  if (doc.Posted !== false)
    throw new InputError(
      "Документ проведён. Отмена проведения выполняется отдельно; автоматической отмены нет.",
    );
  requireTypes(meta, { СуммаДокумента: "Edm.Double", ВалютаДокумента_Key: "Edm.Guid" });
  const md = await conn.getMetadata();
  const rowMeta = md.entities.get(`${p.entitySet}_${ROWS}`);
  requireTypes(rowMeta, {
    [CONTRACT]: "Edm.Guid",
    СуммаПлатежа: "Edm.Double",
    СуммаВзаиморасчетов: "Edm.Double",
  });
  const rowTypes = new Map(rowMeta!.properties.map((f) => [f.name, f.type]));
  if (!Array.isArray(doc[ROWS])) throw new InputError("Расшифровка платежа неизвестна; изменение запрещено.");
  const rows = doc[ROWS] as ODataEntity[];
  if (rows.length !== 1)
    throw new InputError(
      rows.length === 0
        ? "В расшифровке платежа нет строк — выравнивать нечего."
        : "Несколько строк платежа: как разнести сумму, решает человек; изменение запрещено.",
    );
  const row = rows[0]!;
  const total = doc.СуммаДокумента;
  if (typeof total !== "number" || !(total > 0)) throw new InputError("Сумма документа не заполнена.");

  // Рубли и договор в рублях без условных единиц — тогда курс и кратность 1 и считать нечего.
  const available = new Set(md.entities.keys());
  const currencies = resolveEntity(CATALOGS.currencies, available);
  if (!currencies) throw new InputError("Справочник валют не опубликован.");
  const currency = exactRef.parse(doc.ВалютаДокумента_Key);
  const cur = await conn.client.getEntity(
    `${currencies}(${odataGuid(currency)})?$format=json&$select=Ref_Key,Code`,
  );
  if (String(cur.Code ?? "").trim() !== (conn.behavior.regulatedCurrencyCode ?? "643"))
    throw new InputError(
      "Документ не в валюте регламентированного учёта: курс не угадываем, изменение запрещено.",
    );
  const contractRef = exactRef.parse(row[CONTRACT]);
  if (contractRef === EMPTY_GUID)
    throw new InputError("В строке не указан договор — сначала apply_bank_settlement_settings.");
  const contracts = resolveEntity(CATALOGS.contracts, available);
  if (!contracts) throw new InputError("Договоры не опубликованы.");
  const contract = await conn.client.getEntity(`${contracts}(${odataGuid(contractRef)})?$format=json`);
  if (
    contract.РасчетыВУсловныхЕдиницах !== false ||
    String(contract.ВалютаВзаиморасчетов_Key).toLowerCase() !== currency
  )
    throw new InputError(
      "Договор в условных единицах или в другой валюте: курс не угадываем, изменение запрещено.",
    );
  const percent = vatPercent(row.СтавкаНДС);
  if (percent === undefined)
    throw new InputError(`Ставка НДС строки «${String(row.СтавкаНДС)}» не поддержана.`);

  const round2 = (n: number) => Math.round(n * 100) / 100;
  // Число — в том виде, в каком его отдаёт 1С: Int64 (кратность) приходит строкой «1». Иначе «1» → 1
  // выглядело бы изменением, а контрольное чтение после записи не совпало бы с отправленным.
  const asObserved = (field: string, n: number) => (typeof row[field] === "string" ? String(n) : n);
  const numeric: Record<string, number> = {
    СуммаПлатежа: total,
    СуммаВзаиморасчетов: total,
    ...(rowTypes.has("КурсВзаиморасчетов") ? { КурсВзаиморасчетов: 1 } : {}),
    ...(rowTypes.has("КратностьВзаиморасчетов") ? { КратностьВзаиморасчетов: 1 } : {}),
    ...(rowTypes.has("СуммаНДС") ? { СуммаНДС: round2((total * percent) / (100 + percent)) } : {}),
  };
  const aligned: ODataEntity = {
    ...row,
    ...Object.fromEntries(Object.entries(numeric).map(([k, n]) => [k, asObserved(k, n)])),
  };
  const changes = Object.entries(numeric)
    .filter(([k, n]) => !(Number.isFinite(Number(row[k])) && Math.abs(Number(row[k]) - n) < 1e-9))
    .map(([k]) => ({ field: k, from: row[k] ?? null, to: aligned[k] }));
  if (changes.length === 0)
    throw new InputError("Суммы строки уже совпадают с суммой документа; изменение не нужно.");
  const payload: ODataEntity = { [ROWS]: [aligned] };
  const headers = { "If-Match": p.expectedDataVersion };
  const prepared = {
    method: "PATCH",
    path,
    payload,
    headers,
    observedDocument: published(doc, meta),
    changes,
  };
  if (!confirm) {
    await conn.client.prepareCreate(p.entitySet, prepared, { withRefKey: false });
    return {
      dryRun: true,
      database: conn.cfg.name,
      ...prepared,
      note: "Предпросмотр. Меняется только эта строка расшифровки; проведение — отдельным вызовом.",
    };
  }
  const result = await conn.client.writeRecordOnce(
    p.entitySet,
    prepared,
    () =>
      verifyAfter(
        () => conn.client.request(path, "PATCH", payload, headers),
        () => conn.client.getEntity(path),
        withoutVersion({ ...published(doc, meta), ...payload }),
        withoutVersion,
      ),
    { kind: "documentState", ref: p.ref },
  );
  return {
    updated: true,
    ...readbackResult(result),
    database: conn.cfg.name,
    entitySet: p.entitySet,
    ref: result.Ref_Key,
    dataVersion: result.DataVersion,
    changes,
  };
}

const withoutVersion = (record: ODataEntity): ODataEntity =>
  Object.fromEntries(Object.entries(record).filter(([field]) => field !== "DataVersion"));
/** Separate bound action: fresh read guards version; OData Post CAS is not claimed. */
export async function guardedDocumentAction(
  conn: Connection,
  input: { entitySet: string; ref: string; post: boolean; repost?: boolean; expectedDataVersion: string },
  confirm: boolean,
): Promise<Record<string, unknown>> {
  const p = input;
  if (p.repost && !p.post) throw new InputError("repost=true требует post=true.");
  const { doc, meta, path } = await readVersionedDocument(conn, p.entitySet, p.ref, p.expectedDataVersion);
  if (doc.Posted === p.post && !p.repost)
    throw new InputError("Документ уже имеет требуемое состояние; действие не отправлено.");
  const action = p.post ? "Post" : "Unpost",
    actionPath = `${p.entitySet}(${odataGuid(p.ref)})/${action}?$format=json${p.post ? "&PostingModeOperational=false" : ""}`;
  const prepared = {
    method: "POST",
    path: actionPath,
    expectedDataVersion: p.expectedDataVersion,
    observedDocument: published(doc, meta),
  };
  const note =
    "Версия проверяется свежим чтением перед действием. Атомарная проверка If-Match для Post/Unpost не подтверждена; требуется последовательная работа без параллельных изменений.";
  if (!confirm) {
    await conn.client.prepareCreate(p.entitySet, prepared, { withRefKey: false });
    return { dryRun: true, database: conn.cfg.name, ...prepared, note };
  }
  const result = await conn.client.writeRecordOnce(
    p.entitySet,
    prepared,
    () =>
      verifyAfter(
        () => conn.client.action(actionPath),
        () => conn.client.getEntity(path),
        { Ref_Key: p.ref, DeletionMark: false, Posted: p.post },
      ),
    { kind: "documentState", ref: p.ref },
  );
  return {
    done: true,
    ...readbackResult(result),
    accountingMovementsVerified: false,
    database: conn.cfg.name,
    ref: result.Ref_Key,
    dataVersion: result.DataVersion,
    action,
    note,
  };
}
