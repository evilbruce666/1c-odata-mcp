import { registerExistingBankSettlementTools } from "../tools/existing-bank-settlement.js";
import { registerCounterpartyGroupTools } from "../tools/counterparty-groups.js";
import { registerSettlementSettingsTools } from "../tools/settlement-settings.js";
import { readFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/server";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { ServerContext } from "../context.js";
import { createResultSchema } from "../schemas/output.js";
import { fingerprintWriteInput, withWriteOperation } from "../odata/write-operation-context.js";
import { registerMetaTools } from "../tools/meta.js";
import { registerCounterpartyTools } from "../tools/counterparties.js";
import { registerDocumentTools } from "../tools/documents.js";
import { registerRegisterTools } from "../tools/registers.js";
import { registerDebtTools } from "../tools/debts.js";
import { registerMonthCloseTools } from "../tools/month-close.js";
import { registerCardTools } from "../tools/card.js";
import { registerPrompts } from "./prompts.js";
import { APPS_EXTENSION_ID, appMetaFor, registerApps } from "./apps.js";
import { registerCashflowTools } from "../tools/cashflow.js";
import { registerSalesTools } from "../tools/sales.js";
import { registerOrganizationTools } from "../tools/organization.js";
import { registerWriteTools } from "../tools/write.js";
import { registerAuditTools } from "../tools/audit.js";
import { registerFreshnessTools } from "../tools/freshness.js";
import { READ_HINTS, WRITE_HINTS, DESTRUCTIVE_HINTS, fail, guard, ok } from "../tools/_shared.js";

/** Версия берётся из package.json (в собранном пакете он на два уровня выше dist/mcp/). */
function readVersion(): string {
  try {
    const url = new URL("../../package.json", import.meta.url);
    const pkg = JSON.parse(readFileSync(url, "utf8")) as { version?: string };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

const INSTRUCTIONS =
  "Доступ к данным 1С:Предприятие через OData. По умолчанию только чтение " +
  "(аналитика: дебиторка, остатки, продажи, движение денег; справочники и документы). " +
  "Запись включается отдельно и по умолчанию работает в режиме предпросмотра (dry-run): " +
  "сначала показывайте пользователю, что будет создано/изменено, и выполняйте запись " +
  "только после явного согласия (confirm=true). У инструментов есть параметр database " +
  "(см. read.system.list_databases) и organization (см. read.system.list_organizations). " +
  "Для создания объекта при подтверждении передавайте operationId из предпросмотра; " +
  "при таймауте повторяйте подтверждение с тем же id и теми же аргументами, а не с новым предпросмотром. " +
  "Это защищает от дубликата, если ответ 1С потерялся. " +
  "Если настроено несколько баз и из вопроса неясно какая, спросите пользователя или проверьте " +
  "read.system.data_freshness, а не полагайтесь на базу по умолчанию: при её подстановке в ответе чтения есть " +
  "databaseDefaulted, а запись без database при нескольких базах отказывает. " +
  "Ref_Key действуют только в своей базе: чужая ссылка даёт пустой результат, а не ошибку.";

/**
 * Правки документа, которые нельзя слепо повторять: инструмент перечитывает документ, и повтор
 * добавил бы ещё строку или удалил соседнюю. Как и создание, защищены operationId и журналом.
 */
/** Правки существующего банковского документа: итог несёт readback, повтор — квитанцию журнала. */
const BANK_DOCUMENT_TOOLS = new Set([
  "write.document.apply_bank_settlement_settings",
  "write.document.align_bank_payment_amounts",
]);
const JOURNALED_UPDATE_TOOLS = new Set([
  "write.document.apply_bank_settlement_settings",
  "write.document.align_bank_payment_amounts",
  "write.document.post_document",
  "write.accounting.set_contract_settlement_settings",
  "write.accounting.set_settlement_settings",
  "write.counterparty.create_group",
  "write.counterparty.set_group",
  "write.document.add_document_line",
  "write.document.remove_document_line",
]);

/** Подсказки клиенту о характере инструмента — по имени (одна точка вместо правок 55 конфигов). */
function annotationsFor(name: string) {
  if (name === "write.document.post_document" || name === "write.entity.mark_for_deletion")
    return DESTRUCTIVE_HINTS;
  if (name.startsWith("read.")) return READ_HINTS;
  return WRITE_HINTS;
}

/**
 * Ответ чтения всегда называет базу, а подставленная по умолчанию при нескольких настроенных базах
 * помечена: пустой результат в чужой базе выглядит как «данных нет». Опрашивать другие базы здесь
 * нельзя: это лишние запросы, разные юрлица и ложные подсказки там, где пусто законно.
 */
export function withDatabaseInfo(
  ctx: ServerContext,
  args: Record<string, unknown>,
  result: CallToolResult,
): CallToolResult {
  const data = result.structuredContent;
  if (result.isError || !data || typeof data !== "object") return result;
  const requested = typeof args.database === "string" ? args.database : undefined;
  const name = ctx.db(requested).cfg.name;
  const others = ctx
    .databases()
    .map((d) => d.name)
    .filter((n) => n !== name);
  const structuredContent = {
    database: name,
    ...data,
    ...(!requested && others.length
      ? {
          databaseDefaulted: true,
          databaseNote: `Использована база по умолчанию ${name}. Также доступны: ${others.join(", ")}.`,
        }
      : {}),
  };
  const text = JSON.stringify(data, null, 2);
  const content = result.content.map((item) =>
    item.type === "text" && item.text === text
      ? { ...item, text: JSON.stringify(structuredContent, null, 2) }
      : item,
  );
  return { ...result, content, structuredContent };
}

/**
 * Запись при нескольких настроенных базах — только с явной базой, и предпросмотр тоже: иначе
 * документ можно создать не в той организации, а предпросмотр и подтверждение разойдутся по аргументам.
 */
export function requireDatabase(
  ctx: ServerContext,
  cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
): (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult> {
  return async (args, extra) => {
    const names = ctx.databases().map((d) => d.name);
    if (names.length > 1 && typeof args.database !== "string") {
      return fail(
        `Настроено несколько баз (${names.join(", ")}): для записи укажите database. Ничего не записано. ` +
          "Если неясно какая база, спросите пользователя.",
      );
    }
    return cb(args, extra);
  };
}

/** Создаёт MCP-сервер и регистрирует все инструменты (чтение + гейтованная запись). */
export function createServer(ctx: ServerContext): McpServer {
  const server = new McpServer(
    { name: "1c-odata-mcp", version: readVersion() },
    { instructions: INSTRUCTIONS, capabilities: { extensions: { [APPS_EXTENSION_ID]: {} } } },
  );

  // Оборачиваем registerTool, чтобы каждому инструменту проставить аннотации по имени.
  const original = server.registerTool.bind(server) as unknown as (
    name: string,
    config: Record<string, unknown>,
    cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
  ) => unknown;
  const registerTool = (
    name: string,
    config: Record<string, unknown>,
    cb: (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
  ) => {
    const appMeta = appMetaFor(name);
    const annotatedConfig = {
      ...config,
      annotations: {
        ...annotationsFor(name),
        ...(config.annotations as Record<string, unknown> | undefined),
      },
      ...(appMeta ? { _meta: { ...(config._meta as Record<string, unknown> | undefined), ...appMeta } } : {}),
    };
    const isCreateTool = name.startsWith("write.") && config.outputSchema === createResultSchema;
    if (name.startsWith("read.") && name !== "read.system.list_databases") {
      return original(name, annotatedConfig, async (args, extra) =>
        withDatabaseInfo(ctx, args, await cb(args, extra)),
      );
    }
    if (!isCreateTool && !JOURNALED_UPDATE_TOOLS.has(name)) {
      return original(name, annotatedConfig, name.startsWith("write.") ? requireDatabase(ctx, cb) : cb);
    }

    // v2: inputSchema — z.object(...); добавляем operationId через extend (разворот zod-объекта
    // как простого объекта даёт не схему, а его внутренние поля).
    const operationIdField = z
      .string()
      .uuid()
      .optional()
      .describe("UUID из предпросмотра. Передайте его при confirm=true, чтобы повтор не создал дубликат.");
    const base = config.inputSchema;
    const inputSchema =
      base instanceof z.ZodObject
        ? base.extend({ operationId: operationIdField })
        : { ...(base as Record<string, unknown>), operationId: operationIdField };
    const wrapped = async (args: Record<string, unknown>, extra: unknown): Promise<CallToolResult> => {
      if (name === "write.document.post_document" && args.expectedDataVersion === undefined)
        return cb(args, extra);
      const operationId = typeof args.operationId === "string" ? args.operationId : undefined;
      if (args.confirm === true && !operationId) {
        return fail("Сначала выполните dry-run. При confirm=true передайте operationId из его результата.");
      }
      const token = operationId ?? randomUUID();
      const requestHash = fingerprintWriteInput(name, args as Record<string, unknown>);
      // Повтор уже подтверждённой операции: ответ из журнала — до того, как инструмент заново
      // прочитает 1С (иначе правка строк применилась бы к уже изменённому документу).
      if (args.confirm === true && operationId) {
        let replay: CallToolResult | undefined;
        const checked = await guard(name, async () => {
          const conn = ctx.db(typeof args.database === "string" ? args.database : undefined);
          const settled = await conn.client.operationSettled(operationId, requestHash);
          if (settled) {
            replay = ok({
              ...(isCreateTool ? { created: true } : { updated: true }),
              ...(name === "write.document.post_document"
                ? {
                    done: true,
                    action: args.post ? "Post" : "Unpost",
                    verified: false,
                    verificationScope: "journal_receipt",
                    readbackFetched: false,
                    accountingMovementsVerified: false,
                    readbackNote:
                      "Журнал хранит только квитанцию; полный документ и проводки повторно не читались.",
                  }
                : {}),
              // Правка банковского документа отдаёт документ из контрольного чтения (#50); повтор — нет.
              ...(BANK_DOCUMENT_TOOLS.has(name)
                ? {
                    verified: false,
                    verificationScope: "journal_receipt",
                    readbackFetched: false,
                    readbackNote: "Журнал хранит только квитанцию; документ повторно не читался.",
                  }
                : {}),
              replayed: true,
              note: "Операция с этим operationId уже выполнена; повторный запрос в 1С не отправлялся.",
              database: conn.cfg.name,
              entitySet: settled.entitySet,
              ref: settled.result["Ref_Key"],
              ...(settled.result["RecordPath"] ? { recordPath: settled.result["RecordPath"] } : {}),
              ...(settled.result["Code"] ? { code: settled.result["Code"] } : {}),
              operationId,
            });
          }
          return ok({});
        });
        if (checked.isError) return checked;
        if (replay) return replay;
      }
      return withWriteOperation(token, requestHash, async () => {
        const result = await cb(args, extra);
        if (result?.isError || !result?.structuredContent || typeof result.structuredContent !== "object")
          return result;
        const structuredContent = { ...result.structuredContent, operationId: token };
        const content = result.content.map((item) =>
          item.type === "text" ? { ...item, text: JSON.stringify(structuredContent, null, 2) } : item,
        );
        return { ...result, content, structuredContent };
      });
    };
    return original(name, { ...annotatedConfig, inputSchema }, requireDatabase(ctx, wrapped));
  };
  Reflect.set(server, "registerTool", (name: string, config: unknown, cb: unknown) =>
    registerTool(
      name,
      config as Record<string, unknown>,
      cb as (args: Record<string, unknown>, extra: unknown) => Promise<CallToolResult>,
    ),
  );

  registerMetaTools(server, ctx);
  registerCounterpartyTools(server, ctx);
  registerDocumentTools(server, ctx);
  registerRegisterTools(server, ctx);
  registerDebtTools(server, ctx);
  registerMonthCloseTools(server, ctx);
  registerCardTools(server, ctx);
  registerCashflowTools(server, ctx);
  registerSalesTools(server, ctx);
  registerOrganizationTools(server, ctx);
  registerAuditTools(server, ctx);
  registerFreshnessTools(server, ctx);
  registerWriteTools(server, ctx);
  registerExistingBankSettlementTools(server, ctx);
  registerSettlementSettingsTools(server, ctx);
  registerCounterpartyGroupTools(server, ctx);
  registerPrompts(server);
  registerApps(server);

  return server;
}
