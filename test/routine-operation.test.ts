import { afterEach, describe, expect, it, vi } from "vitest";
import { callTool } from "./helpers/call-tool.js";

const SET = "Document_РегламентнаяОперация";
const row = {
  Ref_Key: "11111111-1111-4111-8111-111111111111",
  Number: "0000-000001",
  Date: "2030-01-31T23:59:59",
  Posted: false,
  DeletionMark: false,
  ВидОперации: "ЗакрытиеСчетов",
  Состояние: "Выполнено",
};

afterEach(() => vi.unstubAllGlobals());

describe("регламентные операции: выполнение — по состоянию, а не по проведению", () => {
  it("в строке поиска есть operation и state", async () => {
    const { result, violations, queries } = await callTool(
      "read.document.search_documents",
      { entitySet: SET },
      (q) => (q.base === SET ? [row] : undefined),
    );
    expect(violations).toEqual([]);
    expect(result.structuredContent?.["rows"]).toMatchObject([
      { posted: false, operation: "ЗакрытиеСчетов", state: "Выполнено" },
    ]);
    expect(queries.find((q) => q.base === SET)!.params.get("$select")).toContain("Состояние");
  });

  it("postedOnly для них — отказ с причиной, а не пустой список", async () => {
    const { result } = await callTool("read.document.search_documents", { entitySet: SET, postedOnly: true });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.content)).toContain("смотрите state");
  });

  it("у других документов этих полей нет и они не запрашиваются", async () => {
    const other = "Document_РеализацияТоваровУслуг";
    const { result, queries } = await callTool(
      "read.document.search_documents",
      { entitySet: other, postedOnly: true },
      (q) => (q.base === other ? [row] : undefined),
    );
    expect(result.isError).toBeFalsy();
    expect((result.structuredContent?.["rows"] as object[])[0]).not.toHaveProperty("state");
    expect(queries.find((q) => q.base === other)!.params.get("$select")).not.toContain("Состояние");
  });
});
