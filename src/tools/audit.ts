import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { CATALOGS, resolveEntity } from "../config/mapping.js";
import {
  EMPTY_GUID,
  GUID_RE,
  REGISTER_RECORDS,
  readDocumentHeader,
  recorderFilter,
  requireDocumentEntity,
  resolveNames,
  resolvePostingsSource,
} from "../odata/accounting.js";
import { ODataError } from "../odata/errors.js";
import { NotPublishedError } from "../odata/publication.js";
import { buildQuery, odataGuid } from "../odata/query.js";
import { uuidTimestamp } from "../odata/uuid.js";
import { InputError } from "../errors.js";
import { getDocumentHistoryResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok } from "./_shared.js";

/**
 * Хронология документа по данным стандартного OData — ТОЛЬКО то, что доказуемо.
 *
 * Три разных понятия, которые нельзя смешивать:
 *  - documentDate  — Document.Date, учётная дата документа (не время действия);
 *  - refCreatedAt  — метка генерации Ref_Key, выведенная из UUIDv1 (derived);
 *  - executedAt / modifiedAt — фактическое выполнение / последнее изменение:
 *    через стандартный OData недоступны (журнал регистрации и версии объектов
 *    не опубликованы), поэтому в ответе отсутствуют — никогда не подставляются.
 *
 * DataVersion отдаётся как есть и хронологически НЕ интерпретируется.
 * Period записей регистра — учётное время, не время записи; не используется.
 * Проверка движений по регистру отвечает только на вопрос «есть ли движения» и
 * на метки времени не влияет.
 */

/** Поля шапки, которые берём, если они есть в $metadata данного вида документа. */
export const HISTORY_HEADER_FIELDS = [
  "Ref_Key",
  "DataVersion",
  "Number",
  "Date",
  "Posted",
  "DeletionMark",
  "Организация_Key",
  "ВидОперации",
  "Состояние",
  "Ответственный_Key",
] as const;

export const RESPONSIBLE_FIELD = "Ответственный_Key";
export const USERS_CATALOG = ["Catalog_Пользователи"] as const;
export const REF_CREATED_AT_SOURCE = "Ref_Key UUIDv1 timestamp";
export const REF_CREATED_AT_DESCRIPTION =
  "Ref_Key creation timestamp derived from UUIDv1 (UTC). Это момент генерации ссылки Ref_Key, " +
  "вычисленный из UUID версии 1, — НЕ проверенное время создания, выполнения или проведения документа.";

/**
 * Ограничения, верные для ЛЮБОГО документа — независимо от версии UUID в Ref_Key
 * и от найденных данных.
 */
export const BASE_LIMITATIONS = [
  "Фактическое время выполнения/проведения (executedAt) через стандартный OData недоступно: журнал регистрации не публикуется.",
  "Время последнего изменения (modifiedAt) недоступно: версии объектов через OData не получены.",
  "Кто фактически выполнил или изменил документ — недоступно; responsible — это только реквизит «Ответственный» документа.",
  "Period записей регистра бухгалтерии — учётное время (обычно конец периода), а не время физической записи; для хронологии не используется.",
  "DataVersion — непрозрачное значение версии объекта; хронологически не интерпретируется.",
  "documentDate — учётная дата документа (Document.Date); она не является временем создания или выполнения.",
] as const;

/** Ограничения refCreatedAt — добавляются ТОЛЬКО когда refCreatedAt возвращён (Ref_Key — UUIDv1). */
export const REF_CREATED_AT_LIMITATIONS = [
  "refCreatedAt выведено из UUIDv1 Ref_Key (derived): это момент генерации ссылки, а не проверенное время создания, выполнения или проведения документа.",
  "Ref_Key мог быть задан программно или перенесён из другой базы (обмен, загрузка данных) — тогда refCreatedAt отражает генерацию ссылки в другой базе.",
  "Корректность часов и часовой пояс машины, сгенерировавшей Ref_Key, по данным стандартного OData проверить нельзя.",
  "Поле node в UUID не доказывает, что ссылку сгенерировал конкретный физический сервер 1С.",
] as const;

type Row = Record<string, unknown>;
const str = (v: unknown): string | undefined => (typeof v === "string" && v !== "" ? v : undefined);
const bool = (v: unknown): boolean | undefined => (typeof v === "boolean" ? v : undefined);
const normGuid = (v: unknown): string =>
  typeof v === "string" ? v.replace(/[{}]/g, "").trim().toLowerCase() : "";
const refValue = (v: unknown): string | undefined =>
  typeof v === "string" && v !== "" && normGuid(v) !== EMPTY_GUID ? v : undefined;
const errText = (e: unknown): string =>
  e instanceof ODataError ? `[${e.kind}] ${e.message}` : e instanceof Error ? e.message : String(e);

export type MovementsStatus = "exists" | "none" | "unsupported" | "error";

export interface AccountingMovements {
  /**
   * exists/none — запрос выполнен, ответ однозначен; unsupported — проверку нельзя
   * выполнить в этой публикации; error — запрос не удался или ответ противоречив.
   * unsupported/error НИКОГДА не означают «движений нет».
   */
  status: MovementsStatus;
  exists?: boolean;
  source?: string;
  filter?: string;
  detail?: string;
}

/**
 * Есть ли у документа движения в регистре бухгалтерии Хозрасчетный.
 * Один GET с отбором по регистратору на стороне 1С и $top=1 — без листания и
 * без какого-либо отката к выборке регистра целиком.
 */
export async function checkAccountingMovements(
  conn: Connection,
  documentEntity: string,
  ref: string,
): Promise<{ result: AccountingMovements; rowsScanned: number }> {
  let entitySet: string;
  let hasRecorderType: boolean;
  try {
    const source = await resolvePostingsSource(conn);
    entitySet = source.entitySet;
    hasRecorderType = source.properties.has(REGISTER_RECORDS.recorderType);
  } catch (e) {
    if (e instanceof ODataError) return { result: { status: "error", detail: errText(e) }, rowsScanned: 0 };
    const detail =
      e instanceof NotPublishedError
        ? "Регистр бухгалтерии «Хозрасчётный» не опубликован в OData."
        : errText(e);
    return { result: { status: "unsupported", detail }, rowsScanned: 0 };
  }

  const filter = recorderFilter(documentEntity, ref);
  const select = [REGISTER_RECORDS.recorder, ...(hasRecorderType ? [REGISTER_RECORDS.recorderType] : [])];
  const base = { source: entitySet, filter };
  let rows: Row[];
  try {
    const page = await conn.client.getCollection(`${entitySet}${buildQuery({ select, filter, top: 1 })}`);
    if (!page || !Array.isArray(page.value)) {
      return {
        result: { ...base, status: "error", detail: "1С вернула ответ без массива value." },
        rowsScanned: 0,
      };
    }
    rows = page.value as Row[];
  } catch (e) {
    return { result: { ...base, status: "error", detail: errText(e) }, rowsScanned: 0 };
  }

  if (rows.length === 0) return { result: { ...base, status: "none", exists: false }, rowsScanned: 0 };
  if (rows.length > 1) {
    return {
      result: {
        ...base,
        status: "error",
        detail: `Запрошена 1 запись, 1С вернула ${rows.length}: $top не применён.`,
      },
      rowsScanned: rows.length,
    };
  }
  const r = rows[0]!;
  const recType = r[REGISTER_RECORDS.recorderType];
  const typeOk =
    recType === undefined ||
    (typeof recType === "string" && recType.slice(recType.lastIndexOf(".") + 1) === documentEntity);
  if (normGuid(r[REGISTER_RECORDS.recorder]) !== normGuid(ref) || !typeOk) {
    return {
      result: {
        ...base,
        status: "error",
        detail: "1С вернула запись другого регистратора: серверный отбор по регистратору не применился.",
      },
      rowsScanned: 1,
    };
  }
  return { result: { ...base, status: "exists", exists: true }, rowsScanned: 1 };
}

export type ResponsibleResolution = "resolved" | "not_found" | "catalog_not_published" | "lookup_failed";

export interface Responsible {
  ref: string;
  name?: string;
  resolution: ResponsibleResolution;
  source: string;
  detail?: string;
}

/**
 * Ответственный_Key → Catalog_Пользователи.Description. Сбой поиска имени не
 * роняет инструмент и не превращается в «неизвестного автора»: ref сохраняется,
 * а resolution/detail говорят, что именно не удалось.
 */
export async function resolveResponsible(conn: Connection, ref: string): Promise<Responsible> {
  const source = `Document.${RESPONSIBLE_FIELD}`;
  const users = resolveEntity(USERS_CATALOG, await conn.available());
  if (!users) {
    return {
      ref,
      resolution: "catalog_not_published",
      source,
      detail: `${USERS_CATALOG.join(", ")} не опубликован в OData — имя не получено.`,
    };
  }
  const src = `${source} → ${users}.Description`;
  try {
    const row = (await conn.client.getEntity(
      `${users}(${odataGuid(ref)})${buildQuery({ select: ["Ref_Key", "Description"] })}`,
    )) as Row | undefined;
    if (!row || typeof row !== "object") {
      return { ref, resolution: "lookup_failed", source: src, detail: "1С вернула пустой ответ." };
    }
    if (typeof row["Description"] !== "string") {
      return { ref, resolution: "lookup_failed", source: src, detail: "В ответе нет поля Description." };
    }
    const name = str(row["Description"]);
    return { ref, ...(name ? { name } : {}), resolution: "resolved", source: src };
  } catch (e) {
    if (e instanceof ODataError && e.kind === "not_found") {
      return {
        ref,
        resolution: "not_found",
        source: src,
        detail: `Пользователь с Ref_Key ${ref} не найден.`,
      };
    }
    return { ref, resolution: "lookup_failed", source: src, detail: errText(e) };
  }
}

interface Evidence {
  sourceEntity: string;
  sourceRef?: string;
  sourceField: string;
  timestamp?: string;
  description: string;
  details?: Record<string, unknown>;
}

/**
 * История (хронология) одного документа. Только GET-запросы:
 * $metadata → шапка документа по ключу → (имя организации) → (имя ответственного)
 * → проверка движений по регистратору с $top=1.
 */
export async function getDocumentHistory(conn: Connection, documentEntity: string, documentRef: string) {
  const t0 = Date.now();
  const ref = documentRef.trim().replace(/^\{|\}$/g, "");
  if (!GUID_RE.test(ref))
    throw new InputError(`documentRef должен быть GUID (Ref_Key документа): ${documentRef}`);
  const em = await requireDocumentEntity(conn, documentEntity);
  const props = new Set(em.properties.map((p) => p.name));
  const doc = (await readDocumentHeader(conn, em, ref, HISTORY_HEADER_FIELDS)) as Row;

  // Ref_Key обязателен и должен совпадать с запрошенным — иначе это не тот документ.
  const docRef = doc["Ref_Key"];
  if (typeof docRef !== "string" || !GUID_RE.test(docRef.replace(/[{}]/g, ""))) {
    throw new Error(`1С вернула документ ${documentEntity} без корректного поля Ref_Key.`);
  }
  if (normGuid(docRef) !== normGuid(ref)) {
    throw new Error(`1С вернула документ с другим Ref_Key (${docRef}) вместо ${ref}.`);
  }

  const limitations: string[] = [...BASE_LIMITATIONS];
  const evidence: Evidence[] = [];

  // documentDate — только из Document.Date и только как учётная дата.
  const documentDate = str(doc["Date"]);
  if (documentDate) {
    evidence.push({
      sourceEntity: documentEntity,
      sourceRef: docRef,
      sourceField: "Date",
      timestamp: documentDate,
      description: "Document.Date — учётная дата документа; не время создания и не время выполнения.",
    });
  } else {
    limitations.push("У документа нет значения Date — documentDate не возвращается.");
  }

  // refCreatedAt — только из самого Ref_Key (UUIDv1 RFC 4122).
  const uuid = uuidTimestamp(docRef);
  let refCreatedAt: { value: string; source: string; confidence: "derived"; description: string } | undefined;
  if (uuid.kind === "v1") {
    refCreatedAt = {
      value: uuid.timestampIso,
      source: REF_CREATED_AT_SOURCE,
      confidence: "derived",
      description: REF_CREATED_AT_DESCRIPTION,
    };
    evidence.push({
      sourceEntity: documentEntity,
      sourceRef: docRef,
      sourceField: "Ref_Key",
      timestamp: uuid.timestampIso,
      description:
        "Ref_Key — UUID версии 1 варианта RFC 4122; 60-битная метка времени (100 нс от 1582-10-15Z) " +
        "переведена в UTC. Поле node — часть UUID, не доказанный идентификатор сервера.",
      details: {
        uuidVersion: uuid.version,
        uuidVariant: uuid.variant,
        timestamp100ns: uuid.timestamp100ns,
        clockSequence: uuid.clockSequence,
        node: uuid.node,
        nodeMulticastBit: uuid.nodeMulticastBit,
      },
    });
    limitations.push(...REF_CREATED_AT_LIMITATIONS);
  } else {
    evidence.push({
      sourceEntity: documentEntity,
      sourceRef: docRef,
      sourceField: "Ref_Key",
      description: uuid.reason,
      details: { uuidVersion: uuid.version, uuidVariant: uuid.variant },
    });
    limitations.push(`refCreatedAt не возвращается: ${uuid.reason}`);
  }

  // Организация — как в get_document_postings.
  const orgRef = refValue(doc["Организация_Key"]);
  let orgName: string | undefined;
  if (orgRef) {
    const orgSet = resolveEntity(CATALOGS.organizations, await conn.available());
    if (orgSet) orgName = (await resolveNames(conn, orgSet, [orgRef])).get(orgRef) || undefined;
  }

  // Ответственный — только реквизит документа, не автор/исполнитель.
  let responsible: Responsible | undefined;
  if (!props.has(RESPONSIBLE_FIELD)) {
    limitations.push(
      `У вида документа ${documentEntity} нет реквизита ${RESPONSIBLE_FIELD} — responsible не возвращается.`,
    );
  } else {
    const respRef = refValue(doc[RESPONSIBLE_FIELD]);
    if (!respRef)
      limitations.push(`${RESPONSIBLE_FIELD} у документа не заполнен — responsible не возвращается.`);
    else {
      responsible = await resolveResponsible(conn, respRef);
      if (responsible.resolution !== "resolved") {
        limitations.push(
          `Имя ответственного не получено (${responsible.resolution}): ${responsible.detail ?? ""}`.trim(),
        );
      }
    }
  }

  const movements = await checkAccountingMovements(conn, documentEntity, docRef);
  if (movements.result.status === "unsupported" || movements.result.status === "error") {
    limitations.push(
      `Наличие движений по регистру бухгалтерии не установлено (${movements.result.status}) — это не означает, что движений нет.`,
    );
  }

  const posted = bool(doc["Posted"]);
  const deletionMark = bool(doc["DeletionMark"]);
  const dataVersion = typeof doc["DataVersion"] === "string" ? doc["DataVersion"] : undefined;

  return {
    database: conn.cfg.name,
    document: {
      entitySet: documentEntity,
      ref: docRef,
      ...(str(doc["Number"]) ? { number: str(doc["Number"]) } : {}),
      ...(posted !== undefined ? { posted } : {}),
      ...(deletionMark !== undefined ? { deletionMark } : {}),
      ...(orgName ? { organization: orgName } : {}),
      ...(orgRef ? { organizationRef: orgRef } : {}),
      ...(str(doc["ВидОперации"]) ? { operation: str(doc["ВидОперации"]) } : {}),
      ...(str(doc["Состояние"]) ? { state: str(doc["Состояние"]) } : {}),
      ...(dataVersion !== undefined ? { dataVersion } : {}),
    },
    timestamps: {
      ...(documentDate ? { documentDate } : {}),
      ...(refCreatedAt ? { refCreatedAt } : {}),
    },
    ...(responsible ? { responsible } : {}),
    accountingMovements: movements.result,
    evidence,
    limitations,
    scan: { rowsScanned: movements.rowsScanned, elapsedMs: Date.now() - t0 },
  };
}

export function registerAuditTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "read.audit.get_document_history",
    {
      title: "Хронология документа (аудит)",
      description:
        "Что можно доказуемо сказать о времени документа по данным стандартного OData. " +
        "Возвращает раздельно: documentDate (Document.Date — учётная дата, не время действия) и " +
        "refCreatedAt (метка генерации Ref_Key, выведенная из UUIDv1; confidence=derived — НЕ " +
        "проверенное время создания/выполнения/проведения). Фактическое время выполнения и " +
        "последнего изменения через стандартный OData недоступны и не возвращаются. Также: " +
        "ответственный (реквизит документа, не автор), DataVersion как есть (не интерпретируется), " +
        "наличие движений в регистре Хозрасчетный (только факт наличия), evidence и limitations. " +
        "Только чтение. documentEntity — Document_* из list_entities, documentRef — Ref_Key.",
      inputSchema: {
        database: databaseField,
        documentEntity: z
          .string()
          .trim()
          .regex(/^Document_[^\s()'/?#&]+$/, "Имя документа вида Document_<Имя>")
          .describe("Имя документа, напр. Document_РегламентнаяОперация"),
        documentRef: z
          .string()
          .trim()
          .regex(
            /^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$/,
            "GUID",
          )
          .describe("Ref_Key документа (GUID)"),
      },
      outputSchema: getDocumentHistoryResultSchema,
    },
    ({ database, documentEntity, documentRef }) =>
      guard("read.audit.get_document_history", async () =>
        ok(await getDocumentHistory(ctx.db(database), documentEntity, documentRef)),
      ),
  );
}
