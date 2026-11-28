import { afterEach, describe, expect, it, vi } from "vitest";
import { withoutBinary } from "../src/tools/binary-fields.js";
import { callTool } from "./helpers/call-tool.js";

const SET = "Document_РегламентнаяОперация";
const REF = "11111111-1111-4111-8111-111111111111";
const BLOB = "QUJD".repeat(5000);
const data = (q: { base: string }) =>
  q.base === SET
    ? [{ Ref_Key: REF, Number: "1", ОписаниеОшибок_Base64Data: BLOB, ОписаниеОшибок_Type: "application/xml" }]
    : undefined;

afterEach(() => vi.unstubAllGlobals());

describe("двоичные поля документа", () => {
  it.each([false, true])("get_document не отдаёт их и называет опущенные (compact=%s)", async (compact) => {
    const { result } = await callTool(
      "read.document.get_document",
      { entitySet: SET, ref: REF, compact },
      data,
    );
    expect(result.isError).toBeFalsy();
    expect(JSON.stringify(result)).not.toContain(BLOB.slice(0, 100));
    expect(result.structuredContent).not.toHaveProperty("ОписаниеОшибок_Base64Data");
    expect(result.structuredContent?.["omittedBinary"]).toEqual(["ОписаниеОшибок_Base64Data"]);
    expect(result.structuredContent?.["Number"]).toBe("1");
  });

  it("includeBinary=true возвращает поле как есть", async () => {
    const { result } = await callTool(
      "read.document.get_document",
      { entitySet: SET, ref: REF, includeBinary: true },
      data,
    );
    expect(result.structuredContent?.["ОписаниеОшибок_Base64Data"]).toBe(BLOB);
    expect(result.structuredContent).not.toHaveProperty("omittedBinary");
  });

  it("вложенные в строки — тоже, без повторов; пустые остаются", () => {
    const { entity, omitted } = withoutBinary({
      Пустое_Base64Data: "",
      Строки: [
        { Файл_Base64Data: "QUI=", Имя: "а" },
        { Файл_Base64Data: "Q0Q=", Имя: "б" },
      ],
    });
    expect(omitted).toEqual(["Строки[].Файл_Base64Data"]);
    expect(entity).toEqual({ Пустое_Base64Data: "", Строки: [{ Имя: "а" }, { Имя: "б" }] });
  });
});
