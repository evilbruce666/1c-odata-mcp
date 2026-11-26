import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ODataError } from "../src/odata/errors.js";
import { WriteOperationJournal } from "../src/odata/write-journal.js";

/**
 * Окно между созданием lock-файла и продолжением вызова, который его создал: отметка «запрос идёт» ставилась после
 * `await open`, и одновременное подтверждение, получив EEXIST раньше, считало lock брошенным и слало второй POST.
 * Тесты одновременных подтверждений это ловили только случайно: порядок ответов файловой системы они не задавали.
 * Здесь он задан: ответ первого успешного `open` lock-файла придерживается, пока остальные не отработают.
 */
const fs = vi.hoisted(() => ({
  /** Один раз: после создания lock-файла — сообщить и ждать. */
  hold: undefined as undefined | { reached: () => void; gate: Promise<void> },
  /** Один раз: ошибка открытия lock-файла. */
  fail: undefined as undefined | Error,
}));
vi.mock("node:fs/promises", async (original) => {
  const real = await original<typeof import("node:fs/promises")>();
  return {
    ...real,
    open: async (...args: Parameters<typeof real.open>) => {
      const isLock = String(args[0]).endsWith(".lock");
      if (isLock && fs.fail) {
        const error = fs.fail;
        fs.fail = undefined;
        throw error;
      }
      const handle = await real.open(...args);
      const hold = isLock ? fs.hold : undefined;
      if (hold) {
        fs.hold = undefined;
        hold.reached();
        await hold.gate;
      }
      return handle;
    },
  };
});

const OP = "6f1c2a3b-4d5e-4f60-8a7b-9c0d1e2f3a4b";
const REF = "aaaaaaaa-0000-1000-8000-000000000001";
const SET = "Document_УстановкаЦенНоменклатуры";
const check = { kind: "lineCount" as const, ref: REF, section: "Товары", before: 1, expected: 2 };

beforeEach(() => {
  fs.hold = undefined;
  fs.fail = undefined;
});

const journal = () =>
  new WriteOperationJournal(mkdtempSync(join(tmpdir(), "1c-odata-cc-")), "ut", "http://x/");

function holdFirstLock() {
  let reached!: () => void;
  let release!: () => void;
  const created = new Promise<void>((r) => (reached = r));
  fs.hold = { reached, gate: new Promise<void>((r) => (release = r)) };
  return { created, release };
}

describe("одновременные подтверждения, пока первый ещё не получил ответ о своём lock-файле", () => {
  it.each([1, 2])(
    "создание с назначенной ссылкой, ещё %i подтверждение(й): все — in_progress, POST один",
    async (others) => {
      const j = journal();
      await j.prepare(OP, SET, { a: 1 }, "h", REF);
      const send = vi.fn(async () => ({ Ref_Key: REF }));
      const confirm = () => j.execute(OP, SET, { a: 1 }, "h", send, undefined, async () => undefined);
      const lock = holdFirstLock();
      const first = confirm();
      await lock.created;
      const rest = await Promise.allSettled(Array.from({ length: others }, confirm));
      lock.release();
      await expect(first).resolves.toMatchObject({ Ref_Key: REF });
      expect(rest.map((r) => (r as PromiseRejectedResult).reason?.kind)).toEqual(
        Array(others).fill("in_progress"),
      );
      expect(send).toHaveBeenCalledTimes(1);
      expect(j.isLive(OP)).toBe(false);
    },
  );

  it("правка без повтора: второе — in_progress, а не «исход неизвестен»; POST один", async () => {
    const j = journal();
    await j.prepare(OP, SET, { ref: REF }, "h");
    const send = vi.fn(async () => ({ Ref_Key: REF }));
    const confirm = () => j.execute(OP, SET, { ref: REF }, "h", send, check);
    const lock = holdFirstLock();
    const first = confirm();
    await lock.created;
    const second = await confirm().catch((e: unknown) => e);
    lock.release();
    await expect(first).resolves.toMatchObject({ Ref_Key: REF });
    expect(second).toMatchObject({ kind: "in_progress" });
    expect(send).toHaveBeenCalledTimes(1);
    expect(j.isLive(OP)).toBe(false);
  });

  it("lock-файл не открылся (нет прав) — отметка снята, следующее подтверждение проходит", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    const send = vi.fn(async () => ({ Ref_Key: REF }));
    const confirm = () => j.execute(OP, SET, { a: 1 }, "h", send, undefined, async () => undefined);
    fs.fail = Object.assign(new Error("EACCES"), { code: "EACCES" });
    await expect(confirm()).rejects.toThrow("EACCES");
    expect(j.isLive(OP)).toBe(false);
    await expect(confirm()).resolves.toMatchObject({ Ref_Key: REF });
    expect(send).toHaveBeenCalledTimes(1);
  });

  // Чужой lock без живого запроса: все выходы из этой ветки снимают отметку — иначе операция до перезапуска
  // отвечала бы in_progress.
  it("правка с брошенным lock-файлом: «исход неизвестен» и при повторе, отметка снята", async () => {
    const j = journal();
    await j.prepare(OP, SET, { ref: REF }, "h");
    const dir = (j as unknown as { directory: string }).directory;
    const { writeFileSync } = await import("node:fs");
    writeFileSync(join(dir, `${OP}.lock`), "");
    const send = vi.fn(async () => ({ Ref_Key: REF }));
    for (let i = 0; i < 2; i++) {
      await expect(j.execute(OP, SET, { ref: REF }, "h", send, check)).rejects.toThrow(/неизвестен/);
      expect(j.isLive(OP)).toBe(false);
    }
    expect(send).not.toHaveBeenCalled();
  });

  it("повтор по брошенному lock-файлу закончился неизвестным исходом — отметка снята, следующий проходит", async () => {
    const j = journal();
    await j.prepare(OP, SET, { a: 1 }, "h", REF);
    const dir = (j as unknown as { directory: string }).directory;
    const { writeFileSync } = await import("node:fs");
    writeFileSync(join(dir, `${OP}.lock`), "");
    const lost = () => Promise.reject(new ODataError({ kind: "network", message: "fetch failed" }));
    await expect(j.execute(OP, SET, { a: 1 }, "h", lost, undefined, async () => undefined)).rejects.toThrow();
    expect(j.isLive(OP)).toBe(false);
    await expect(
      j.execute(
        OP,
        SET,
        { a: 1 },
        "h",
        async () => ({ Ref_Key: REF }),
        undefined,
        async () => undefined,
      ),
    ).resolves.toMatchObject({ Ref_Key: REF });
    expect(j.isLive(OP)).toBe(false);
  });
});
