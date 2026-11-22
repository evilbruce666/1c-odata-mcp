import { mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Client, InMemoryTransport } from "@modelcontextprotocol/client";
import { createServer } from "../src/mcp/server.js";
import { fake1C, type FakeQuery, type FakeSchema } from "./helpers/fake-1c.js";

/**
 * Сбои и повторы записи партнёра и контрагента Управления торговлей (перенесено из ревью выпуска 2). Подделка 1С здесь
 * помнит созданное: объект находится по ссылке и отбором по ИНН. Сбои — на каждый POST: ответ потерян после записи,
 * запрос не дошёл, отказ 1С; перезапуск — новый сервер на том же журнале. Критерий: ни второго объекта, ни ложного успеха.
 */
type Fail =
  | "lost" // 1С записала, ответ потерян (обрыв)
  | "lostNoWrite" // запрос не дошёл
  | "reject" // 1С отказала (400)
  | "write500" // 1С записала и ответила 500
  | "nowrite500" // 500 до записи
  | "hang" // 1С записала и молчит до таймаута клиента (ODATA_TIMEOUT_MS)
  | "held" // ждёт w.release(): «живой» запрос, 1С пишет после разблокировки
  | undefined;
interface W {
  method: string;
  set: string;
  body?: Record<string, unknown>;
}

const PS = "Catalog_Партнеры";
const CS = "Catalog_Контрагенты";
const BA = "Catalog_БанковскиеСчета";

/** Мир 1С с состоянием: созданное видно GET по ссылке и отбором по ИНН (как в настоящей базе). */
function world(
  o: { schema?: FakeSchema; ka?: boolean; noPartners?: boolean; env?: Record<string, string> } = {},
) {
  const dir = mkdtempSync(join(tmpdir(), "1c-odata-retry-"));
  const writes: W[] = [];
  const store: Record<string, Map<string, Record<string, unknown>>> = {
    [PS]: new Map(),
    [CS]: new Map(),
    [BA]: new Map(),
  };
  const fails: Record<string, Fail[]> = { [PS]: [], [CS]: [], [BA]: [] };
  /** Сбой PATCH владельца (makeMain): запрос не дошёл. */
  const patchFails: Array<"lostNoWrite"> = [];
  let fake = make();
  function make() {
    const f = fake1C(
      { READ_ONLY: "false", ODATA_WRITABLE: "true", ODATA_WRITE_JOURNAL_DIR: dir, ...o.env },
      {
        schema: o.schema ?? "ut11",
        data: (q: FakeQuery) => {
          const flt = q.params.get("$filter") ?? "";
          if (q.base === CS && flt.includes("ИНН eq")) {
            const inn = /ИНН eq '([^']*)'/.exec(flt)?.[1];
            return [...store[CS]!.values()].filter((r) => r["ИНН"] === inn && r["DeletionMark"] !== true);
          }
          if (q.base === PS && !flt) return [...store[PS]!.values()];
          const byName = /Description eq '((?:[^']|'')*)'/.exec(flt)?.[1]?.replace(/''/g, "'");
          if (q.base === PS && byName !== undefined)
            return [...store[PS]!.values()].filter(
              (r) => r["Description"] === byName && r["DeletionMark"] !== true,
            );
          return undefined;
        },
      },
    );
    if (o.ka || o.noPartners) {
      const orig = f.conn.available.bind(f.conn);
      (f.conn as unknown as { available: () => Promise<Set<string>> }).available = async () => {
        const s = new Set(await orig());
        if (o.ka) s.add("AccountingRegister_Хозрасчетный");
        if (o.noPartners) s.delete(PS);
        return s;
      };
    }
    return f;
  }
  const inner = globalThis.fetch;
  const hide = new Map<string, number>();
  let release!: () => void;
  const gate = new Promise<void>((r) => (release = r));
  const json = (b: unknown, status = 200) =>
    new Response(JSON.stringify(b), { status, headers: { "content-type": "application/json" } });
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = decodeURIComponent(
      typeof input === "string" ? input : input instanceof URL ? input.href : input.url,
    );
    const method = (init?.method ?? "GET").toUpperCase();
    const set = /standard\.odata\/([^?(/]+)/.exec(url)?.[1] ?? "";
    if (method === "PATCH") {
      writes.push({ method, set, body: JSON.parse(String(init?.body)) as Record<string, unknown> });
      if (patchFails.shift() === "lostNoWrite") {
        writes.pop();
        throw new TypeError("fetch failed");
      }
    }
    if (method === "POST" && store[set]) {
      const body = JSON.parse(String(init?.body)) as Record<string, unknown>;
      writes.push({ method, set, body });
      const f = fails[set]!.shift();
      if (f === "reject")
        return json(
          { "odata.error": { code: "-1", message: { lang: "ru", value: "Synthetic rejection" } } },
          400,
        );
      if (f === "lostNoWrite") throw new TypeError("fetch failed");
      const err500 = (value: string) =>
        json({ "odata.error": { code: "-1", message: { lang: "ru", value } } }, 500);
      if (f === "nowrite500") return err500("Synthetic server error");
      if (f === "held") await gate;
      const ref = String(body["Ref_Key"]);
      // Как уникальный индекс 1С: вторая запись с той же ссылкой — отказ, объект прежний.
      if (store[set]!.has(ref)) return err500("Нарушено условие уникальности данных");
      const row = { ...body, DeletionMark: false, ГоловнойКонтрагент_Key: ref, Code: "0001" };
      store[set]!.set(ref, row);
      if (f === "lost") throw new TypeError("fetch failed");
      if (f === "write500") return err500("Synthetic server error");
      if (f === "hang")
        await new Promise<never>((_, reject) =>
          init?.signal?.addEventListener("abort", () => reject(new DOMException("aborted", "AbortError")), {
            once: true,
          }),
        );
      return json(row, 201);
    }
    const m = /standard\.odata\/(Catalog_[^?(]+)\(guid'([0-9a-f-]{36})'\)/i.exec(url);
    if (method === "GET" && m && store[m[1]!]) {
      // Чтение «не видит» объект один раз: первый запрос дописался между чтением повтора и его POST.
      const hidden = hide.get(m[2]!) ?? 0;
      if (hidden > 0 && hide.set(m[2]!, hidden - 1))
        return json({ "odata.error": { code: "-1", message: { lang: "ru", value: "not found" } } }, 404);
      const row = store[m[1]!]!.get(m[2]!);
      return row
        ? json(row)
        : json({ "odata.error": { code: "-1", message: { lang: "ru", value: "not found" } } }, 404);
    }
    return inner(input, init);
  });
  const call = async (name: string, args: Record<string, unknown>) => {
    const [a, b] = InMemoryTransport.createLinkedPair();
    await createServer(fake.ctx).connect(a);
    const client = new Client({ name: "t", version: "1" });
    await client.connect(b);
    const r = (await client.callTool({ name, arguments: args })) as {
      isError?: boolean;
      structuredContent?: Record<string, unknown>;
      content: { text: string }[];
    };
    return {
      isError: !!r.isError,
      sc: r.structuredContent ?? {},
      text: r.content.map((c) => c.text).join(""),
    };
  };
  const age = (op: string) => {
    for (const sub of readdirSync(dir)) {
      const f = join(dir, sub, `${op}.json`);
      try {
        const e = JSON.parse(readFileSync(f, "utf8"));
        writeFileSync(f, JSON.stringify({ ...e, updatedAt: new Date(Date.now() - 600_000).toISOString() }));
      } catch {
        /* */
      }
    }
  };
  const restart = () => {
    fake = make();
  };
  return {
    call,
    writes,
    store,
    fails,
    age,
    restart,
    release,
    dir,
    hide,
    patchFails,
    get fake() {
      return fake;
    },
  };
}

const PT = "write.counterparty.create_partner";
const CP = "write.counterparty.create_counterparty";
const ST = "write.operation.status";
const pArgs = { name: "ZZ-TEST партнёр", relations: ["client"] };
afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("create_partner: сбои и повторы", () => {
  // 1.1: повтор подтверждения после неизвестного исхода не блокируется — сначала чтение по назначенной ссылке.
  it("ответ потерян, 1С записала: повтор находит объект по ссылке без второго POST; после перезапуска — из журнала; партнёр один", async () => {
    const w = world();
    w.fails[PS] = ["lost"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    const c1 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    const c2 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(c2.sc).toMatchObject({ replayed: true, reconciled: true, ref: p.sc.ref });
    w.restart();
    const st = await w.call(ST, { operationId: op });
    const c3 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    expect(c1.isError).toBe(true);
    expect(st.sc.status).toBe("succeeded");
    expect(c3.sc).toMatchObject({ replayed: true, ref: st.sc.ref });
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
    expect(w.store[PS]!.size).toBe(1);
  });

  it("запрос не дошёл: status not_found, закрытие not_applied, тот же id — отказ, новый предпросмотр — один партнёр", async () => {
    const w = world();
    w.fails[PS] = ["lostNoWrite"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    const st1 = await w.call(ST, { operationId: op });
    w.age(op);
    const st2 = await w.call(ST, { operationId: op, markNotApplied: true });
    const c2 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    const p2 = await w.call(PT, pArgs);
    const c3 = await w.call(PT, { ...pArgs, confirm: true, operationId: p2.sc.operationId });
    expect(st1.sc.status).toBe("not_found");
    expect(st2.sc.status).toBe("not_applied");
    expect(c2.isError).toBe(true);
    expect(c3.isError).toBe(false);
    expect(w.store[PS]!.size).toBe(1);
  });

  it("1С отказала: status rejected, тот же id — отказ, ничего не создано", async () => {
    const w = world();
    w.fails[PS] = ["reject"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    const c1 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    const st = await w.call(ST, { operationId: op });
    const c2 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    expect(c1.isError && c2.isError).toBe(true);
    expect(st.sc.status).toBe("rejected");
    expect(w.store[PS]!.size).toBe(0);
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
  });

  it("ответ потерян, затем новый предпросмотр без status: предупреждение о тёзке со ссылкой на созданного", async () => {
    const w = world();
    w.fails[PS] = ["lost"];
    const p = await w.call(PT, pArgs);
    await w.call(PT, { ...pArgs, confirm: true, operationId: p.sc.operationId });
    const p2 = await w.call(PT, pArgs);
    const created = [...w.store[PS]!.keys()][0]!;
    expect(JSON.stringify(p2.sc.notes)).toContain(created);
    expect(JSON.stringify(p2.sc.notes)).toContain("write.operation.status");
  });

  it("тело записи — как в предпросмотре; privatePerson по умолчанию и явное false — одна операция", async () => {
    const w = world();
    const p = await w.call(PT, pArgs);
    const c = await w.call(PT, {
      ...pArgs,
      privatePerson: false,
      confirm: true,
      operationId: p.sc.operationId,
    });
    expect(c.isError, c.text.slice(0, 300)).toBe(false);
    const body = { ...w.writes[0]!.body! };
    delete body["Ref_Key"];
    expect(body).toEqual(p.sc.payload);
  });

  it("допуск: УТ без партнёров — отказ, КА (Хозрасчётный и партнёры) — допуск, Бухгалтерия — отказ; записей нет", async () => {
    const ut = world({ noPartners: true });
    const r1 = await ut.call(PT, pArgs);
    vi.unstubAllGlobals();
    const ka = world({ ka: true });
    const r2 = await ka.call(PT, pArgs);
    vi.unstubAllGlobals();
    const bp = world({ schema: "bp30" });
    const r3 = await bp.call(PT, pArgs);
    expect(r1.isError).toBe(true);
    expect(r2.isError).toBe(false);
    expect(r3.isError).toBe(true);
    expect(ut.writes.length + ka.writes.length + bp.writes.length).toBe(0);
  });
});

describe("create_counterparty в УТ: сбои и повторы", () => {
  const seedPartner = (w: ReturnType<typeof world>) => {
    const ref = "aaaaaaaa-0000-0000-0000-00000000000a";
    w.store[PS]!.set(ref, { Ref_Key: ref, Description: "P", DeletionMark: false });
    return ref;
  };
  it("ответ потерян, 1С записала: status сверяет, повтор — replayed; тело — как в предпросмотре; контрагент один", async () => {
    const w = world();
    const partnerRef = seedPartner(w);
    const args = { name: "ZZ-TEST Ромашка", inn: "7700000001", partnerRef, phone: "+7 900 000-00-00" };
    w.fails[CS] = ["lost"];
    const p = await w.call(CP, args);
    const op = p.sc.operationId as string;
    const c1 = await w.call(CP, { ...args, confirm: true, operationId: op });
    // 1.1: повтор находит созданного по ссылке раньше, чем инструмент проверит дубль ИНН (иначе — ложный отказ).
    const c2 = await w.call(CP, { ...args, confirm: true, operationId: op });
    const st = await w.call(ST, { operationId: op });
    expect(c1.isError).toBe(true);
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(c2.sc).toMatchObject({ reconciled: true });
    expect(st.sc.status).toBe("succeeded");
    const c3 = await w.call(CP, { ...args, confirm: true, operationId: op });
    const body = { ...w.writes[0]!.body! };
    delete body["Ref_Key"];
    expect(body).toEqual(p.sc.payload);
    expect(w.store[CS]!.size).toBe(1);
    expect(c3.sc.replayed).toBe(true);
  });

  it("партнёр с потерянным ответом сверен через status — контрагент к нему; каждого по одному", async () => {
    const w = world();
    w.fails[PS] = ["lost"];
    const p = await w.call(PT, pArgs);
    const pop = p.sc.operationId as string;
    const c1 = await w.call(PT, { ...pArgs, confirm: true, operationId: pop });
    expect(c1.isError).toBe(true);
    const st = await w.call(ST, { operationId: pop });
    expect(st.sc.status).toBe("found_reconciled");
    const args = { name: "ZZ-TEST", inn: "7700000003", partnerRef: st.sc.ref };
    const cp = await w.call(CP, args);
    const done = await w.call(CP, { ...args, confirm: true, operationId: cp.sc.operationId });
    expect(done.isError, done.text.slice(0, 300)).toBe(false);
    expect(w.store[PS]!.size).toBe(1);
    expect(w.store[CS]!.size).toBe(1);
  });

  it("1С отказала контрагенту: новый предпросмотр создаёт одного", async () => {
    const w = world();
    const partnerRef = seedPartner(w);
    const args = { name: "ZZ-TEST", inn: "7700000004", partnerRef };
    w.fails[CS] = ["reject"];
    const p = await w.call(CP, args);
    const c1 = await w.call(CP, { ...args, confirm: true, operationId: p.sc.operationId });
    const st = await w.call(ST, { operationId: p.sc.operationId });
    expect(c1.isError).toBe(true);
    expect(st.sc.status).toBe("rejected");
    const p2 = await w.call(CP, args);
    const c2 = await w.call(CP, { ...args, confirm: true, operationId: p2.sc.operationId });
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(w.store[CS]!.size).toBe(1);
  });

  it("ответ потерян, затем новый предпросмотр: с ИНН — отказ по дублю со ссылкой", async () => {
    const w = world();
    const partnerRef = seedPartner(w);
    const withInn = { name: "ZZ-TEST Р", inn: "7700000005", partnerRef };
    w.fails[CS] = ["lost"];
    const q = await w.call(CP, withInn);
    await w.call(CP, { ...withInn, confirm: true, operationId: q.sc.operationId });
    const q2 = await w.call(CP, withInn);
    expect(q2.isError).toBe(true);
    expect(q2.text).toContain("уже есть");
  });

  it("партнёра пометили после подтверждения — повтор всё равно правдиво replayed", async () => {
    const w = world();
    const partnerRef = seedPartner(w);
    const args = { name: "ZZ-TEST", inn: "7700000006", partnerRef };
    const p = await w.call(CP, args);
    const c = await w.call(CP, { ...args, confirm: true, operationId: p.sc.operationId });
    expect(c.isError, c.text.slice(0, 300)).toBe(false);
    w.store[PS]!.get(partnerRef)!["DeletionMark"] = true;
    const again = await w.call(CP, { ...args, confirm: true, operationId: p.sc.operationId });
    expect(again.sc.replayed).toBe(true);
  });

  it("Бухгалтерия: create_counterparty без партнёра, relations схемой отброшен", async () => {
    const w = world({ schema: "bp30" });
    const r = await w.call(CP, { name: "ZZ", inn: "7700000007", relations: ["client"] });
    expect(r.isError, r.text.slice(0, 300)).toBe(false);
    expect(r.sc.payload).not.toHaveProperty("Партнер_Key");
  });
});

// ─── 1.1: безопасный повтор создания с назначенной ссылкой ───

/** Запись журнала операции на диске (каталог журнала — подкаталог базы). */
function journalFile(dir: string, op: string): string {
  for (const sub of readdirSync(dir)) {
    const f = join(dir, sub, `${op}.json`);
    try {
      readFileSync(f);
      return f;
    } catch {
      /* */
    }
  }
  throw new Error("нет записи журнала");
}

describe("1.1: повтор подтверждения создания после неизвестного исхода", () => {
  const confirm = (w: ReturnType<typeof world>, op: string) =>
    w.call(PT, { ...pArgs, confirm: true, operationId: op });

  it("предпросмотр отдаёт назначенную ссылку — с ней и создаётся объект", async () => {
    const w = world();
    const p = await w.call(PT, pArgs);
    expect(p.sc.ref).toMatch(/^[0-9a-f-]{36}$/);
    const c = await confirm(w, p.sc.operationId as string);
    expect(c.sc.ref).toBe(p.sc.ref);
    expect([...w.store[PS]!.keys()]).toEqual([p.sc.ref]);
  });

  for (const fail of ["lostNoWrite", "nowrite500"] as const)
    it(`${fail}: до 1С не дошло — повтор отправляет тот же POST с той же ссылкой и создаёт один объект`, async () => {
      const w = world();
      w.fails[PS] = [fail];
      const p = await w.call(PT, pArgs);
      const op = p.sc.operationId as string;
      const c1 = await confirm(w, op);
      expect(c1.isError).toBe(true);
      expect(c1.text).toContain("Повторите подтверждение с тем же operationId");
      const st = await w.call(ST, { operationId: op });
      expect(st.sc).toMatchObject({ status: "not_found", retrySafe: true });
      const c2 = await confirm(w, op);
      expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
      expect(c2.sc).toMatchObject({ created: true, retried: true, ref: p.sc.ref });
      const writes = w.writes.filter((x) => x.set === PS);
      expect(writes).toHaveLength(2);
      expect(writes[1]!.body).toEqual(writes[0]!.body);
      expect(w.store[PS]!.size).toBe(1);
      const c3 = await confirm(w, op);
      expect(c3.sc).toMatchObject({ replayed: true, ref: p.sc.ref });
      expect(w.writes.filter((x) => x.set === PS)).toHaveLength(2);
    });

  it("write500: 1С записала и ответила 500 — повтор находит объект, второго POST нет", async () => {
    const w = world();
    w.fails[PS] = ["write500"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    expect((await confirm(w, op)).isError).toBe(true);
    const c2 = await confirm(w, op);
    expect(c2.sc).toMatchObject({ reconciled: true, ref: p.sc.ref });
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
    expect(w.store[PS]!.size).toBe(1);
  });

  it("таймаут при медленной 1С (записала, ответ не дождались) — повтор находит объект, второго POST нет", async () => {
    const w = world({ env: { ODATA_TIMEOUT_MS: "100" } });
    w.fails[PS] = ["hang"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    const c1 = await confirm(w, op);
    expect(c1.isError).toBe(true);
    expect(c1.text).toContain("неизвестен");
    const c2 = await confirm(w, op);
    expect(c2.sc).toMatchObject({ reconciled: true, ref: p.sc.ref });
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
    expect(w.store[PS]!.size).toBe(1);
  });

  it("первый запрос дописался между чтением и POST повтора: отказ по уникальности, повторное чтение — успех, объект один", async () => {
    const w = world();
    w.fails[PS] = ["lost"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    await confirm(w, op);
    // Повтор читает по ссылке дважды (до сборки записи и перед POST) — оба раза «не видит».
    w.hide.set(String(p.sc.ref), 2);
    const c2 = await confirm(w, op);
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(c2.sc).toMatchObject({ reconciled: true, ref: p.sc.ref });
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(2);
    expect(w.store[PS]!.size).toBe(1);
  });

  it("повтор снова не дошёл — исход по-прежнему неизвестен, повтор безопасен; третий создаёт один объект", async () => {
    const w = world();
    w.fails[PS] = ["lostNoWrite", "lostNoWrite"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    await confirm(w, op);
    const c2 = await confirm(w, op);
    expect(c2.isError).toBe(true);
    expect(c2.text).toContain("Повторите подтверждение с тем же operationId");
    const c3 = await confirm(w, op);
    expect(c3.sc).toMatchObject({ retried: true });
    expect(w.store[PS]!.size).toBe(1);
  });

  it("два одновременных подтверждения: второе — in_progress, без POST; первое создаёт; объект один", async () => {
    const w = world();
    w.fails[PS] = ["held"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    const first = confirm(w, op);
    await vi.waitFor(() => expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1));
    const second = await confirm(w, op);
    const st = await w.call(ST, { operationId: op });
    w.release();
    const r1 = await first;
    expect(second.isError).toBe(true);
    expect(second.text).toContain("[in_progress]");
    expect(st.sc).toMatchObject({ status: "not_found", inProgress: true, retrySafe: false });
    expect(r1.isError, r1.text.slice(0, 300)).toBe(false);
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
    expect(w.store[PS]!.size).toBe(1);
  });

  for (const written of [false, true])
    it(`рестарт коннектора посреди POST (lock упавшего процесса, 1С ${written ? "записала" : "не получила"}) — повтор доводит до одного объекта`, async () => {
      const w = world();
      const p = await w.call(PT, pArgs);
      const op = p.sc.operationId as string;
      const ref = String(p.sc.ref);
      const file = journalFile(w.dir, op);
      const entry = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>;
      writeFileSync(file, JSON.stringify({ ...entry, state: "executing" }));
      writeFileSync(file.replace(/\.json$/, ".lock"), "boot-of-crashed-process");
      if (written) w.store[PS]!.set(ref, { Ref_Key: ref, Description: pArgs.name, DeletionMark: false });
      // Процесс упал: в журнале «executing», lock с чужой меткой запуска, живого запроса нет.
      const c = await confirm(w, op);
      expect(c.isError, c.text.slice(0, 300)).toBe(false);
      expect(c.sc).toMatchObject(written ? { reconciled: true } : { retried: true });
      expect(c.sc.ref).toBe(ref);
      expect(w.store[PS]!.size).toBe(1);
      expect(w.writes.filter((x) => x.set === PS)).toHaveLength(written ? 0 : 1);
    });

  it("повтор через сутки при сроке предпросмотра 15 минут: срок повтору не мешает, объект один", async () => {
    const w = world({ env: { ODATA_WRITE_PREVIEW_MAX_AGE: "15" } });
    w.fails[PS] = ["lostNoWrite"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    await confirm(w, op);
    // Часы ушли на сутки вперёд: данные документа (дата) — от предпросмотра, а не от «сейчас».
    vi.useFakeTimers({ toFake: ["Date"], now: Date.now() + 24 * 3600_000 });
    const c = await confirm(w, op);
    expect(c.isError, c.text.slice(0, 300)).toBe(false);
    expect(c.sc).toMatchObject({ retried: true });
    expect(w.store[PS]!.size).toBe(1);
  });

  it("найден помеченный на удаление — успех с deletionMark и предупреждением, без POST", async () => {
    const w = world();
    w.fails[PS] = ["lost"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    await confirm(w, op);
    w.store[PS]!.get(String(p.sc.ref))!["DeletionMark"] = true;
    const c = await confirm(w, op);
    expect(c.sc).toMatchObject({ reconciled: true, deletionMark: true });
    expect(String(c.sc.warning)).toContain("помечен на удаление");
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
  });

  it("другие аргументы при повторе — отказ без POST", async () => {
    const w = world();
    w.fails[PS] = ["lostNoWrite"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    await confirm(w, op);
    const c = await w.call(PT, { ...pArgs, name: "ZZ-TEST другой", confirm: true, operationId: op });
    expect(c.isError).toBe(true);
    // Ревью 1.1: раньше совет «выполните новый предпросмотр» — он создал бы второй объект, если первая запись дошла.
    expect(c.text).toContain("Повторите подтверждение с теми же аргументами");
    expect(w.writes.filter((x) => x.set === PS)).toHaveLength(1);
    expect(w.store[PS]!.size).toBe(0);
  });
});

// ─── 1.1: шаг после создания (makeMain) при повторе ───

describe("1.1: повтор создания с makeMain — владелец получает основной счёт", () => {
  const owner = "bbbbbbbb-0000-0000-0000-00000000000b";
  const args = { ownerRef: owner, accountNumber: "40702810000000000001", bik: "044525225", makeMain: true };
  const BT = "write.counterparty.create_bank_account";
  const mainPatches = (w: ReturnType<typeof world>) =>
    w.writes.filter(
      (x) => x.method === "PATCH" && x.set.startsWith(CS) && "ОсновнойБанковскийСчет_Key" in (x.body ?? {}),
    );

  // Ревью 1.1: ответ из журнала (и сверка по ссылке) шёл мимо инструмента — PATCH владельца не выполнялся, молча.
  it("ответ на создание потерян: повтор находит счёт по ссылке и делает его основным", async () => {
    const w = world({ schema: "bp30" });
    w.fails[BA] = ["lost"];
    const p = await w.call(BT, args);
    expect(p.isError, p.text.slice(0, 300)).toBe(false);
    const op = p.sc.operationId as string;
    const c1 = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c1.isError).toBe(true);
    expect(mainPatches(w)).toHaveLength(0);
    const c2 = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(c2.sc).toMatchObject({ reconciled: true, setAsMain: true, ref: p.sc.ref });
    expect(mainPatches(w).map((x) => x.body!["ОсновнойБанковскийСчет_Key"])).toEqual([p.sc.ref]);
    expect(w.store[BA]!.size).toBe(1);
  });

  it("счёт создан, PATCH владельца не дошёл: повтор — из журнала, без второго счёта, и PATCH владельца", async () => {
    const w = world({ schema: "bp30" });
    w.patchFails.push("lostNoWrite");
    const p = await w.call(BT, args);
    const op = p.sc.operationId as string;
    const c1 = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c1.isError).toBe(true);
    const c2 = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(c2.sc).toMatchObject({ replayed: true, setAsMain: true });
    expect(mainPatches(w)).toHaveLength(1);
    expect(w.writes.filter((x) => x.method === "POST" && x.set === BA)).toHaveLength(1);
    expect(w.store[BA]!.size).toBe(1);
  });

  it("без makeMain повтор по-прежнему отвечает из журнала, без PATCH", async () => {
    const w = world({ schema: "bp30" });
    const a = { ...args, makeMain: false };
    const p = await w.call(BT, a);
    const op = p.sc.operationId as string;
    await w.call(BT, { ...a, confirm: true, operationId: op });
    const again = await w.call(BT, { ...a, confirm: true, operationId: op });
    expect(again.sc).toMatchObject({ replayed: true });
    expect(mainPatches(w)).toHaveLength(0);
  });

  // Ревью 1.1: флаг makeMain выводил подтверждение из-под проверки срока предпросмотра.
  it("срок предпросмотра действует и с makeMain: давний предпросмотр — отказ без записи", async () => {
    const w = world({ schema: "bp30", env: { ODATA_WRITE_PREVIEW_MAX_AGE: "15" } });
    const p = await w.call(BT, args);
    const op = p.sc.operationId as string;
    const file = journalFile(w.dir, op);
    const entry = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>;
    writeFileSync(
      file,
      JSON.stringify({ ...entry, createdAt: new Date(Date.now() - 3 * 3600_000).toISOString() }),
    );
    const c = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c.isError).toBe(true);
    expect(c.text).toContain("ODATA_WRITE_PREVIEW_MAX_AGE");
    expect(w.writes).toHaveLength(0);
  });

  // Ревью 1.1: повтор уже созданного через инструмент сверял заново собранное тело — сменившаяся подстановка из 1С
  // давала совет «новый предпросмотр», а он создал бы второй счёт.
  it("счёт создан, PATCH не дошёл, тело собралось иначе — повтор всё равно из журнала, второго счёта нет", async () => {
    const w = world({ schema: "bp30" });
    w.patchFails.push("lostNoWrite");
    const p = await w.call(BT, args);
    const op = p.sc.operationId as string;
    const c1 = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c1.text).toContain("Объект создан");
    expect(c1.text).toContain("Новый предпросмотр не делайте");
    const file = journalFile(w.dir, op);
    const entry = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>;
    writeFileSync(file, JSON.stringify({ ...entry, payloadHash: "changed" }));
    const c2 = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c2.isError, c2.text.slice(0, 300)).toBe(false);
    expect(c2.sc).toMatchObject({ replayed: true, setAsMain: true });
    expect(w.writes.filter((x) => x.method === "POST" && x.set === BA)).toHaveLength(1);
  });

  it("найденный по ссылке счёт помечен на удаление — основным не ставится", async () => {
    const w = world({ schema: "bp30" });
    w.fails[BA] = ["lost"];
    const p = await w.call(BT, args);
    const op = p.sc.operationId as string;
    await w.call(BT, { ...args, confirm: true, operationId: op });
    w.store[BA]!.get(String(p.sc.ref))!["DeletionMark"] = true;
    const c = await w.call(BT, { ...args, confirm: true, operationId: op });
    expect(c.sc).toMatchObject({ reconciled: true, deletionMark: true, setAsMain: false });
    expect(mainPatches(w)).toHaveLength(0);
  });
});
