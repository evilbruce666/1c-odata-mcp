import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import { WriteOperationJournal } from "../src/odata/write-journal.js";

/**
 * 1.1: безопасный повтор создания — граница применимости в самом журнале. Повтор только у создания с назначенной
 * ссылкой; правки строк, записи регистров и операции без ссылки по-прежнему удерживаются до ручной сверки.
 */
const OP = "6f1c2a3b-4d5e-4f60-8a7b-9c0d1e2f3a4b";
const REF = "aaaaaaaa-0000-1000-8000-000000000001";
const SET = "Document_УстановкаЦенНоменклатуры";

function journal() {
  return new WriteOperationJournal(mkdtempSync(join(tmpdir(), "1c-odata-rj-")), "ut", "http://x/");
}
const lost = () => Promise.reject(new ODataError({ kind: "network", message: "fetch failed" }));

describe("журнал: повтор создания после неизвестного исхода", () => {
  it("данные изменились с предпросмотра, объекта нет — отказ без POST с причиной; объект появился — успех", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    const sent: string[] = [];
    await expect(j.execute(OP, SET, { a: 1 }, "h", lost, undefined, async () => undefined)).rejects.toThrow(
      /Повторите подтверждение/,
    );
    const send = async () => {
      sent.push("post");
      return { Ref_Key: REF };
    };
    await expect(j.execute(OP, SET, { a: 2 }, "h", send, undefined, async () => undefined)).rejects.toThrow(
      /изменились с предпросмотра/,
    );
    expect(sent).toEqual([]);
    const ok = await j.execute(OP, SET, { a: 2 }, "h", send, undefined, async () => ({ Ref_Key: REF }));
    expect(ok).toMatchObject({ Ref_Key: REF, _operation_reconciled: true });
    expect(sent).toEqual([]);
    expect((await j.lookup(OP))?.state).toBe("succeeded");
  });

  it("операция с другой проверкой (правка строк) — повтор не выполняется, исход по-прежнему неизвестен", async () => {
    const j = journal();
    await j.prepare(OP, SET, { ref: REF }, "h");
    const check = { kind: "lineCount" as const, ref: REF, section: "Товары", before: 1, expected: 2 };
    await expect(j.execute(OP, SET, { ref: REF }, "h", lost, check)).rejects.toThrow(/неизвестен/);
    let posts = 0;
    await expect(
      j.execute(
        OP,
        SET,
        { ref: REF },
        "h",
        async () => {
          posts += 1;
          return { Ref_Key: REF };
        },
        check,
        async () => undefined,
      ),
    ).rejects.toThrow(/Повторная отправка заблокирована/);
    expect(posts).toBe(0);
  });

  it("создание без назначенной ссылки (старая запись журнала) — повтор не выполняется", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h");
    await expect(j.execute(OP, SET, { a: 1 }, "h", lost)).rejects.toThrow(/неизвестен/);
    let posts = 0;
    await expect(
      j.execute(
        OP,
        SET,
        { a: 1 },
        "h",
        async () => {
          posts += 1;
          return {};
        },
        undefined,
        async () => undefined,
      ),
    ).rejects.toThrow(/Повторная отправка заблокирована/);
    expect(posts).toBe(0);
  });

  it("отказ 1С на повторе не делает операцию отклонённой: первый запрос мог ещё выполняться", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    await expect(j.execute(OP, SET, { a: 1 }, "h", lost, undefined, async () => undefined)).rejects.toThrow();
    const rejected = () => Promise.reject(new ODataError({ kind: "bad_request", message: "отказ" }));
    await expect(
      j.execute(OP, SET, { a: 1 }, "h", rejected, undefined, async () => undefined),
    ).rejects.toThrow(/неизвестен/);
    expect((await j.lookup(OP))?.state).toBe("outcome_unknown");
  });

  // Живая проверка 1.1: второе одновременное подтверждение попадало в окно между захватом lock и записью «executing»
  // и получало «исход неизвестен, повтор заблокирован». Тест подделки ждал, пока первый POST уже начнётся.
  it("два подтверждения одновременно, с самого начала — второе in_progress, POST один", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    let posts = 0;
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    const send = async () => {
      posts += 1;
      await gate;
      return { Ref_Key: REF };
    };
    const find = async () => undefined;
    const first = j.execute(OP, SET, { a: 1 }, "h", send, undefined, find);
    const second = j.execute(OP, SET, { a: 1 }, "h", send, undefined, find);
    await expect(second).rejects.toMatchObject({ kind: "in_progress" });
    release();
    await expect(first).resolves.toMatchObject({ Ref_Key: REF });
    expect(posts).toBe(1);
  });
});
