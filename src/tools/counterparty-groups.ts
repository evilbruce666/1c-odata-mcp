import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import type { Connection, ServerContext } from "../context.js";
import { InputError } from "../errors.js";
import {
  EMPTY_GUID,
  exactRef,
  counterpartyAncestors,
  counterpartyCatalog,
  readCounterparty,
  MAX_GROUP_DEPTH,
} from "../odata/counterparty-hierarchy.js";
import { ODataError } from "../odata/errors.js";
import { odataGuid } from "../odata/query.js";
import { currentWriteOperationId } from "../odata/write-operation-context.js";
import { patchResultSchema } from "../schemas/output.js";
import { databaseField, guard, ok } from "./_shared.js";

const parentRef = z
  .string()
  .uuid()
  .transform((v) => v.toLowerCase());
const createSchema = z.object({ name: z.string().trim().min(1).max(200), parentRef });
const moveSchema = z.object({
  ref: exactRef,
  subjectKind: z.literal("counterparty"),
  parentRef,
  expectedDataVersion: z
    .string()
    .min(1)
    .max(200)
    .refine((v) => v !== "*" && !/[\r\n]/.test(v)),
});

async function validateParent(conn: Connection, ref: string): Promise<string[]> {
  if (ref === EMPTY_GUID) return [];
  const { groupRefs } = await counterpartyAncestors(conn, ref, true);
  if (groupRefs.length + 1 > MAX_GROUP_DEPTH)
    throw new InputError(
      `Новая цепочка родителей глубже ${MAX_GROUP_DEPTH}; создание или изменение группы запрещено.`,
    );
  return [ref, ...groupRefs];
}

async function writeGroupOperation(
  conn: Connection,
  set: string,
  recordPath: string,
  method: "POST" | "PATCH",
  payload: Record<string, unknown>,
  expected: Record<string, unknown>,
  headers: Record<string, string>,
  confirm: boolean,
) {
  const path = `${method === "POST" ? set : recordPath}?$format=json`;
  const operation = { method, path, payload, headers };
  if (!confirm) {
    await conn.client.prepareCreate(set, operation, { withRefKey: false });
    return { dryRun: true, database: conn.cfg.name, entitySet: set, recordPath, ...operation };
  }
  const record = await conn.client.writeRecordOnce(
    set,
    operation,
    async () => {
      await conn.client.request(path, method, payload, headers);
      try {
        const actual = await conn.client.getEntity(`${recordPath}?$format=json`);
        for (const [name, value] of Object.entries(expected)) {
          const equal =
            name.endsWith("_Key") && typeof value === "string"
              ? String(actual[name]).toLowerCase() === value.toLowerCase()
              : JSON.stringify(actual[name]) === JSON.stringify(value);
          if (!equal) throw new Error(`Field mismatch: ${name}`);
        }
        return { ...actual, RecordPath: recordPath };
      } catch (cause) {
        throw new ODataError({
          kind: "unknown",
          message:
            "Контрольное чтение группы/принадлежности не подтвердило результат; повторная запись запрещена, требуется ручная сверка.",
          cause,
        });
      }
    },
    method === "POST"
      ? { kind: "createdRef", ref: String(payload.Ref_Key) }
      : { kind: "manual", target: "counterpartyGroup" },
  );
  return { updated: true, verified: true, database: conn.cfg.name, entitySet: set, recordPath, record };
}

export function registerCounterpartyGroupTools(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "write.counterparty.create_group",
    {
      title: "Создать группу контрагентов",
      description:
        "Создаёт только явно названную группу контрагентов в указанной группе parentRef; пустой GUID parentRef означает корень. Проверяет родителя и иерархию, сначала предпросмотр, затем confirm с тем же operationId и контрольное чтение. Не создаёт произвольные бизнес-группы без поручения владельца.",
      inputSchema: z.object({
        database: databaseField,
        ...createSchema.shape,
        confirm: z.boolean().default(false),
      }),
      outputSchema: patchResultSchema,
    },
    (args) =>
      guard("write.counterparty.create_group", async () => {
        const p = createSchema.parse(args);
        const conn = ctx.db(args.database);
        const set = await counterpartyCatalog(conn);
        await validateParent(conn, p.parentRef);
        const props = new Map(
          (await conn.getMetadata()).entities.get(set)!.properties.map((v) => [v.name, v.type]),
        );
        if (props.get("Description") !== "Edm.String")
          throw new InputError("Наименование группы не опубликовано как Edm.String.");
        const ref = currentWriteOperationId();
        if (!ref) throw new InputError("Создание требует предварительный просмотр и operationId.");
        const payload = { Ref_Key: ref, Description: p.name, IsFolder: true, Parent_Key: p.parentRef };
        return ok(
          await writeGroupOperation(
            conn,
            set,
            `${set}(${odataGuid(ref)})`,
            "POST",
            payload,
            { ...payload, DeletionMark: false },
            {},
            args.confirm,
          ),
        );
      }),
  );
  server.registerTool(
    "write.counterparty.set_group",
    {
      title: "Указать группу отдельного контрагента",
      description:
        "Меняет только Parent_Key отдельного контрагента (subjectKind=counterparty, IsFolder=false). Перемещение самих групп запрещено: безопасность глубины их потомков не проверяется. parentRef — точная группа или пустой GUID для корня. Проверяет отсутствие удаления и цикла; expectedDataVersion обязателен, PATCH с If-Match, предпросмотр/confirm/operationId и контроль сохранения остальных полей. Удаление недоступно.",
      inputSchema: z.object({
        database: databaseField,
        ...moveSchema.shape,
        confirm: z.boolean().default(false),
      }),
      outputSchema: patchResultSchema,
    },
    (args) =>
      guard("write.counterparty.set_group", async () => {
        const p = moveSchema.parse(args);
        const conn = ctx.db(args.database);
        const set = await counterpartyCatalog(conn);
        const subject = await readCounterparty(conn, set, p.ref, false);
        const ancestors = await validateParent(conn, p.parentRef);
        if (ancestors.includes(p.ref))
          throw new InputError("Нельзя переместить группу в себя или своего потомка.");
        const props = new Map(
          (await conn.getMetadata()).entities.get(set)!.properties.map((v) => [v.name, v.type]),
        );
        if (props.get("DataVersion") !== "Edm.String" || subject.DataVersion !== p.expectedDataVersion)
          throw new InputError("Нужен текущий expectedDataVersion для точного объекта.");
        const payload = { Parent_Key: p.parentRef };
        const expected = {
          ...Object.fromEntries(
            Object.entries(subject).filter(([name]) => props.has(name) && name !== "DataVersion"),
          ),
          ...payload,
        };
        return ok(
          await writeGroupOperation(
            conn,
            set,
            `${set}(${odataGuid(p.ref)})`,
            "PATCH",
            payload,
            expected,
            { "If-Match": p.expectedDataVersion },
            args.confirm,
          ),
        );
      }),
  );
}
