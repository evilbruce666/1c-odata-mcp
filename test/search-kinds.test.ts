import { describe, expect, it } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C } from "./helpers/fake-1c.js";

/**
 * Виды поиска документов Управления торговлей: цены, заказы, заявки на расходование, складские документы УТ. Поиск
 * без kind их не включает — его стоимость прежняя (норма запросов в test/fixtures/request-budget.json).
 */

async function search(args: Record<string, unknown>) {
  const fake = fake1C({ READ_ONLY: "true" }, { schema: "ut11" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await createServer(fake.ctx).connect(a);
  const client = new Client({ name: "t", version: "1" });
  await client.connect(b);
  const r = (await client.callTool({ name: "read.document.search_documents", arguments: args })) as {
    isError?: boolean;
    structuredContent: { searched?: { entitySet: string }[] };
    content: { text: string }[];
  };
  return { r, sets: (r.structuredContent.searched ?? []).map((s) => s.entitySet), fake };
}

describe("search_documents: виды Управления торговлей", () => {
  it.each([
    ["prices", ["Document_УстановкаЦенНоменклатуры"]],
    ["orders", ["Document_ЗаказКлиента", "Document_ЗаказПоставщику"]],
    ["payment_requests", ["Document_ЗаявкаНаРасходованиеДенежныхСредств"]],
  ] as const)("kind=%s ищет %j", async (kind, expected) => {
    const { r, sets, fake } = await search({ kind });
    expect(r.isError, r.content[0]?.text).toBeFalsy();
    expect(sets.sort()).toEqual([...expected].sort());
    expect(fake.violations).toEqual([]);
  });

  it("warehouse в УТ — и оприходование излишков, списание недостач, пересчёт", async () => {
    const { sets } = await search({ kind: "warehouse" });
    for (const set of [
      "Document_ОприходованиеИзлишковТоваров",
      "Document_СписаниеНедостачТоваров",
      "Document_ПересчетТоваров",
    ])
      expect(sets).toContain(set);
  });

  // Ревью: прежний тест проверял только отсутствие четырёх новых видов и пропустил три складских документа УТ
  // (12 → 15 запросов). Теперь — весь список и число запросов, как до появления видов УТ.
  it("без kind — прежний набор видов и прежнее число запросов", async () => {
    const { sets, fake } = await search({});
    expect(sets.sort()).toEqual([
      "Document_ВозвратТоваровОтКлиента",
      "Document_ВозвратТоваровПоставщику",
      "Document_ПеремещениеТоваров",
      "Document_ПоступлениеБезналичныхДенежныхСредств",
      "Document_ПриобретениеТоваровУслуг",
      "Document_ПриходныйКассовыйОрдер",
      "Document_РасходныйКассовыйОрдер",
      "Document_РеализацияТоваровУслуг",
      "Document_СписаниеБезналичныхДенежныхСредств",
      "Document_СчетНаОплатуКлиенту",
      "Document_СчетФактураВыданный",
      "Document_СчетФактураПолученный",
    ]);
    expect(fake.requests.filter((r) => r.resource.startsWith("Document_"))).toHaveLength(12);
  });
});
