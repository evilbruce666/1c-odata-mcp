import { afterEach, describe, expect, it, vi } from "vitest";
import type { CallToolResult } from "@modelcontextprotocol/server";
import { readFileSync } from "node:fs";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

/**
 * Служебные поля (Posted, DeletionMark, Ref_Key, DataVersion) не проходят через свободные «fields» ни одного
 * инструмента записи. update_entity обходил проверки mark_for_deletion и post_document, copy_document — гарантию
 * «копия создаётся непроведённой» и ссылку, которую журнал выдал при предпросмотре (аудит 25.11). Инструменты
 * берутся из manifest.json: новый инструмент со свободными полями попадёт в проверку сам.
 */
const REF = "33333333-3333-4333-8333-333333333333";
const manifest = JSON.parse(readFileSync("manifest.json", "utf8")) as {
  tools: Array<{ name: string; inputSchema: { properties?: Record<string, Record<string, unknown>> } }>;
};
const isFree = (p: Record<string, unknown> | undefined) =>
  !!p && p["type"] === "object" && !!p["additionalProperties"] && !p["properties"];
const freeHeader = manifest.tools
  .filter((t) => isFree(t.inputSchema.properties?.["fields"]))
  .map((t) => t.name);

const ENTITY_SETS = [
  "Document_РеализацияТоваровУслуг",
  "Document_ПоступлениеНаРасчетныйСчет",
  "Catalog_Контрагенты",
];
const SERVICE = [
  ["Posted", true],
  ["DeletionMark", true],
  ["Ref_Key", "44444444-4444-4444-8444-444444444444"],
  ["DataVersion", "AAAA"],
] as const;

afterEach(() => vi.restoreAllMocks());

describe("служебные поля в свободных fields — отказ до записи", () => {
  it("свободные поля шапки — у update_entity и copy_document (список из манифеста)", () => {
    expect(freeHeader.sort()).toEqual(["write.document.copy_document", "write.entity.update_entity"]);
  });

  for (const tool of freeHeader)
    for (const [field, value] of SERVICE)
      it(`${tool}: ${field} — отказ без записи, и в предпросмотре`, async () => {
        for (const entitySet of ENTITY_SETS) {
          const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
          const post = vi.spyOn(f.conn.client, "create");
          const patch = vi.spyOn(f.conn.client, "patch");
          const tools = (
            createServer(f.ctx) as unknown as {
              _registeredTools: Record<
                string,
                { handler: (a: object, e: object) => Promise<CallToolResult> }
              >;
            }
          )._registeredTools;
          // copy_document без предпросмотра не подтвердить (operationId из журнала): отказ в предпросмотре закрывает и
          // подтверждение. update_entity подтверждается без operationId — проверяются оба.
          for (const confirm of tool === "write.entity.update_entity" ? [false, true] : [false]) {
            const r = await tools[tool]!.handler(
              { entitySet, ref: REF, fields: { Комментарий: "x", [field]: value }, confirm },
              {},
            );
            expect(r.isError, `${entitySet} confirm=${confirm}`).toBe(true);
            expect(JSON.stringify(r.content)).toContain(field);
          }
          expect(post).not.toHaveBeenCalled();
          expect(patch).not.toHaveBeenCalled();
        }
      });

  it("copy_document: Ref_Key в полях строки — отказ", async () => {
    const f = fake1C({ READ_ONLY: "false", ODATA_WRITABLE: "true" });
    const post = vi.spyOn(f.conn.client, "create");
    const tools = (
      createServer(f.ctx) as unknown as {
        _registeredTools: Record<string, { handler: (a: object, e: object) => Promise<CallToolResult> }>;
      }
    )._registeredTools;
    const r = await tools["write.document.copy_document"]!.handler(
      {
        entitySet: "Document_РеализацияТоваровУслуг",
        ref: REF,
        lines: [{ lineNumber: 1, fields: { Ref_Key: REF } }],
        confirm: false,
      },
      {},
    );
    expect(r.isError).toBe(true);
    expect(JSON.stringify(r.content)).toContain("Ref_Key");
    expect(post).not.toHaveBeenCalled();
  });
});
