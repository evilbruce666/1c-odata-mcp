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
type Fail = "lost" | "lostNoWrite" | "reject" | undefined;
interface W {
  method: string;
  set: string;
  body?: Record<string, unknown>;
}

const PS = "Catalog_Партнеры";
const CS = "Catalog_Контрагенты";

/** Мир 1С с состоянием: созданное видно GET по ссылке и отбором по ИНН (как в настоящей базе). */
function world(o: { schema?: FakeSchema; ka?: boolean; noPartners?: boolean } = {}) {
  const dir = mkdtempSync(join(tmpdir(), "1c-odata-retry-"));
  const writes: W[] = [];
  const store: Record<string, Map<string, Record<string, unknown>>> = { [PS]: new Map(), [CS]: new Map() };
  const fails: Record<string, Fail[]> = { [PS]: [], [CS]: [] };
  let fake = make();
  function make() {
    const f = fake1C(
      { READ_ONLY: "false", ODATA_WRITABLE: "true", ODATA_WRITE_JOURNAL_DIR: dir },
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
  const json = (b: unknown, status = 200) =>
    new Response(JSON.stringify(b), { status, headers: { "content-type": "application/json" } });
  vi.stubGlobal("fetch", async (input: string | URL | Request, init?: RequestInit) => {
    const url = decodeURIComponent(
      typeof input === "string" ? input : input instanceof URL ? input.href : input.url,
    );
    const method = (init?.method ?? "GET").toUpperCase();
    const set = /standard\.odata\/([^?(/]+)/.exec(url)?.[1] ?? "";
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
      const ref = String(body["Ref_Key"]);
      if (store[set]!.has(ref))
        return json({ "odata.error": { code: "-1", message: { lang: "ru", value: "Duplicate key" } } }, 500);
      const row = { ...body, DeletionMark: false, ГоловнойКонтрагент_Key: ref, Code: "0001" };
      store[set]!.set(ref, row);
      if (f === "lost") throw new TypeError("fetch failed");
      return json(row, 201);
    }
    const m = /standard\.odata\/(Catalog_[^?(]+)\(guid'([0-9a-f-]{36})'\)/i.exec(url);
    if (method === "GET" && m && store[m[1]!]) {
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
  it("ответ потерян, 1С записала: повтор удержан, после перезапуска status сверяет, повтор — replayed; партнёр один", async () => {
    const w = world();
    w.fails[PS] = ["lost"];
    const p = await w.call(PT, pArgs);
    const op = p.sc.operationId as string;
    const c1 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    const c2 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    expect(c2.isError).toBe(true); // повтор удержан журналом: исход неизвестен
    w.restart();
    const st = await w.call(ST, { operationId: op });
    const c3 = await w.call(PT, { ...pArgs, confirm: true, operationId: op });
    expect(c1.isError).toBe(true);
    expect(st.sc.status).toBe("found_reconciled");
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
    const c2 = await w.call(CP, { ...args, confirm: true, operationId: op });
    const st = await w.call(ST, { operationId: op });
    expect(c1.isError && c2.isError).toBe(true);
    expect(st.sc.status).toBe("found_reconciled");
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
