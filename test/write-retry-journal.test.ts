import { mkdtempSync, readdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import { WriteOperationJournal, markLive, unmarkLive } from "../src/odata/write-journal.js";

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
    // Lock достаётся любому из двух (ревью 1.1: тест ждал отказа именно от второго и падал через раз).
    const both = Promise.allSettled([
      j.execute(OP, SET, { a: 1 }, "h", send, undefined, find),
      j.execute(OP, SET, { a: 1 }, "h", send, undefined, find),
    ]);
    await vi.waitFor(() => expect(posts).toBe(1));
    release();
    const results = await both;
    expect(results.filter((r) => r.status === "fulfilled")).toHaveLength(1);
    const rejected = results.find((r) => r.status === "rejected") as PromiseRejectedResult;
    expect(rejected.reason).toMatchObject({ kind: "in_progress" });
    expect(posts).toBe(1);
  });

  // Ревью 1.1: подтверждение, захватившее lock уже после неизвестного исхода первого, уходило в retryCreate без await,
  // и finally снимал отметку «запрос идёт» — соседнее подтверждение отправило бы второй POST (объект один — индекс 1С).
  it("переход в повтор под lock: пока повтор шлёт POST, операция отмечена как идущая", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    const internal = j as unknown as { read: (id: string) => Promise<unknown> };
    const read = internal.read.bind(j);
    let releaseRead!: () => void;
    const readGate = new Promise<void>((r) => (releaseRead = r));
    let calls = 0;
    internal.read = async (id: string) => {
      const entry = await read(id);
      calls += 1;
      if (calls === 1) await readGate; // B прочитал «prepared» и ждёт, пока A закончится неизвестным исходом
      return entry;
    };
    let releasePost!: () => void;
    const postGate = new Promise<void>((r) => (releasePost = r));
    const b = j.execute(
      OP,
      SET,
      { a: 1 },
      "h",
      async () => {
        await postGate;
        return { Ref_Key: REF };
      },
      undefined,
      async () => undefined,
    );
    internal.read = read;
    await expect(j.execute(OP, SET, { a: 1 }, "h", lost, undefined, async () => undefined)).rejects.toThrow();
    internal.read = async (id: string) => read(id);
    releaseRead();
    await vi.waitFor(async () => expect((await j.lookup(OP))?.state).toBe("executing"));
    const liveDuringPost = j.isLive(OP);
    releasePost();
    await expect(b).resolves.toMatchObject({ Ref_Key: REF });
    expect(liveDuringPost).toBe(true);
  });

  // Ревью 1.1: процесс упал между захватом lock и записью «executing» (POST не уходил) — операция висела сутки до
  // чистки, а markNotApplied на «подготовленную» не действовал. Создание с назначенной ссылкой теперь повторяется.
  it("«подготовлено» и lock упавшего процесса: создание с ссылкой — повтор отправляет POST; правка — по-прежнему стоп", async () => {
    const dir = mkdtempSync(join(tmpdir(), "1c-odata-rj-"));
    const j = new WriteOperationJournal(dir, "ut", "http://x/");
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    const sub = readdirSync(dir)[0]!;
    writeFileSync(join(dir, sub, `${OP}.lock`), "");
    let posts = 0;
    const r = await j.execute(
      OP,
      SET,
      { a: 1 },
      "h",
      async () => {
        posts += 1;
        return { Ref_Key: REF };
      },
      undefined,
      async () => undefined,
    );
    expect(r).toMatchObject({ Ref_Key: REF });
    expect(posts).toBe(1);

    const OP2 = "6f1c2a3b-4d5e-4f60-8a7b-9c0d1e2f3a4c";
    await j.prepare(OP2, SET, { ref: REF }, "h");
    writeFileSync(join(dir, sub, `${OP2}.lock`), "");
    await expect(
      j.execute(OP2, SET, { ref: REF }, "h", async () => ({ Ref_Key: REF }), {
        kind: "lineCount",
        ref: REF,
        section: "Товары",
        before: 1,
        expected: 2,
      }),
    ).rejects.toThrow(/неизвестен/);
  });

  // Ревью 1.1: при двух процессах на журнале проигравший повтор с окончательным отказом 1С записывал «отклонено»
  // поверх созданного объекта — новый предпросмотр дал бы дубль.
  it("окончательный отказ при создании, а объект по ссылке есть — успех, не «отклонено»", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    const rejected = () =>
      Promise.reject(new ODataError({ kind: "bad_request", message: "Код не уникален" }));
    const r = await j.execute(OP, SET, { a: 1 }, "h", rejected, undefined, async () => ({ Ref_Key: REF }));
    expect(r).toMatchObject({ Ref_Key: REF, _operation_reconciled: true });
    expect((await j.lookup(OP))?.state).toBe("succeeded");
    // Объекта нет — отказ остаётся отказом.
    const j2 = journal();
    await j2.prepare(OP, SET, { a: 1 }, "h", REF);
    await expect(
      j2.execute(OP, SET, { a: 1 }, "h", rejected, undefined, async () => undefined),
    ).rejects.toThrow(/Код не уникален/);
    expect((await j2.lookup(OP))?.state).toBe("rejected");
  });

  // Ревью 1.1, второй раунд: отметка «запрос идёт» без владельца — вызов снимал отметку соседа.
  it("отметку «запрос идёт» снимает только тот, кто её поставил", () => {
    const j = journal();
    const key = (j as unknown as { liveKey: (id: string) => string }).liveKey(OP);
    const b = markLive(key);
    const c = markLive(key); // сосед перехватил операцию и ведёт запрос
    unmarkLive(key, b); // прежний владелец уходит — отметка соседа остаётся
    expect(j.isLive(OP)).toBe(true);
    unmarkLive(key, c);
    expect(j.isLive(OP)).toBe(false);
  });
});
