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
    verified: true,
    database: conn.cfg.name,
    entitySet: p.entitySet,
    ref: result.Ref_Key,
    dataVersion: result.DataVersion,
    resolution,
    ...(warnings.length ? { warnings } : {}),
  };
}

/**
 * Сумма платежа должна быть разнесена по расшифровке целиком: проводки банковского документа
 * формируются по строкам «Расшифровки платежа», и неразнесённая часть не попадёт в расчёты.
 * Форма 1С держит суммы согласованными сама, расходятся они при записи в обход формы (#61).
 * Инструмент суммы не меняет — только предупреждает до проведения.
 */
export function paymentAmountWarnings(doc: ODataEntity, rows: unknown[]): string[] {
  const total = doc["СуммаДокумента"];
  if (typeof total !== "number") return [];
  const amounts = rows.map((r) => (r as ODataEntity)["СуммаПлатежа"]);
  if (amounts.some((a) => typeof a !== "number")) return [];
  const allocated = (amounts as number[]).reduce((a, b) => a + b, 0);
  if (Math.abs(allocated - total) < 0.005) return [];
  return [
    `Сумма документа ${total} не совпадает с суммой расшифровки платежа ${Math.round(allocated * 100) / 100}` +
      `${rows.length === 0 ? " (строк расшифровки нет)" : ""}. Проводки формируются по расшифровке — ` +
      "до проведения исправьте сумму строки в 1С; этот инструмент суммы не меняет.",
  ];
}
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
    verified: result._operation_replayed !== true,
    verificationScope: result._operation_replayed === true ? "journal_receipt" : "document_state",
    accountingMovementsVerified: false,
    readbackFetched: result._operation_replayed !== true,
    ...(result._operation_replayed === true ? { replayed: true } : { readback: result }),
    database: conn.cfg.name,
    ref: result.Ref_Key,
    dataVersion: result.DataVersion,
    action,
    note,
  };
}
