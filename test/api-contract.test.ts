import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * Совместимость публичного API с последним выпуском (docs/VERSIONING.md): ядро встраивают в чужие боты и
 * интеграции, и до 2.0 нельзя удалить инструмент, поле ответа или параметр, сменить их тип, сделать параметр
 * обязательным, убрать значение перечисления из параметров или пометить инструмент чтения как пишущий. Раньше это
 * держалось на внимательности. Эталон — инструменты и сценарии из manifest.json выпуска
 * (test/fixtures/api-contract/release.json); обновляется при выпуске: npm run contract:update.
 */
type Schema = {
  type?: string | string[];
  properties?: Record<string, Schema>;
  required?: string[];
  items?: Schema;
  enum?: unknown[];
  anyOf?: Schema[];
  default?: unknown;
};
type Tool = {
  name: string;
  inputSchema: Schema;
  outputSchema?: Schema;
  annotations?: Record<string, unknown>;
};
type Api = { version: string; tools: Tool[]; prompts: Array<{ name: string; arguments?: string[] }> };

const release = JSON.parse(readFileSync("test/fixtures/api-contract/release.json", "utf8")) as Api;
const current = JSON.parse(readFileSync("manifest.json", "utf8")) as Api;

/**
 * Изменения, признанные совместимыми вопреки правилу, — с причиной. Пусто: каждое исключение обсуждается и
 * записывается в «Что изменится при обновлении» CHANGELOG.
 */
const ALLOWED: Record<string, string> = {};

const typeOf = (s: Schema | undefined) =>
  s?.type !== undefined
    ? JSON.stringify([s.type].flat().sort())
    : s?.anyOf
      ? JSON.stringify(s.anyOf.flatMap((v) => [v.type ?? "?"].flat()).sort())
      : undefined;

/** Нарушения при переходе old → now. side: input — параметры вызова, output — ответ. */
export function breaks(
  old: Schema | undefined,
  now: Schema | undefined,
  side: "input" | "output",
  path: string,
): string[] {
  if (!old) return [];
  if (!now) return [`${path}: схема удалена`];
  const out: string[] = [];
  const [was, is] = [typeOf(old), typeOf(now)];
  if (was && is && was !== is) out.push(`${path}: тип ${was} → ${is}`);
  if (old.enum && side === "input") {
    const gone = old.enum.filter((v) => !(now.enum ?? []).includes(v));
    if (now.enum && gone.length) out.push(`${path}: убраны значения ${JSON.stringify(gone)}`);
  }
  const oldReq = new Set(old.required ?? []);
  const nowReq = new Set(now.required ?? []);
  for (const [name, prop] of Object.entries(old.properties ?? {})) {
    const p = `${path}.${name}`;
    if (!now.properties?.[name]) {
      out.push(`${p}: удалено`);
      continue;
    }
    if (side === "output" && oldReq.has(name) && !nowReq.has(name))
      out.push(`${p}: было всегда, стало необязательным`);
    out.push(...breaks(prop, now.properties[name], side, p));
  }
  if (side === "input")
    for (const name of nowReq)
      if (!oldReq.has(name))
        out.push(`${path}.${name}: стал обязательным${old.properties?.[name] ? "" : " (новый)"}`);
  if (old.items) out.push(...breaks(old.items, now.items, side, `${path}[]`));
  return out;
}

function contractBreaks(before: Api, after: Api): string[] {
  const found: string[] = [];
  const now = new Map(after.tools.map((t) => [t.name, t]));
  for (const t of before.tools) {
    const n = now.get(t.name);
    if (!n) {
      found.push(`${t.name}: инструмент удалён`);
      continue;
    }
    found.push(...breaks(t.inputSchema, n.inputSchema, "input", `${t.name} параметры`));
    found.push(...breaks(t.outputSchema, n.outputSchema, "output", `${t.name} ответ`));
    if (t.annotations?.["readOnlyHint"] === true && n.annotations?.["readOnlyHint"] !== true)
      found.push(`${t.name}: был только для чтения, стал пишущим`);
    if (t.annotations?.["destructiveHint"] === false && n.annotations?.["destructiveHint"] === true)
      found.push(`${t.name}: стал разрушающим (destructiveHint)`);
  }
  const prompts = new Map(after.prompts.map((p) => [p.name, p]));
  for (const p of before.prompts) {
    const n = prompts.get(p.name);
    if (!n) found.push(`сценарий ${p.name}: удалён`);
    else
      for (const a of p.arguments ?? [])
        if (!(n.arguments ?? []).includes(a)) found.push(`сценарий ${p.name}: удалён аргумент ${a}`);
  }
  return found.filter((f) => !Object.keys(ALLOWED).some((k) => f.startsWith(k)));
}

describe(`публичный API совместим с выпуском ${release.version}`, () => {
  it("ничего не удалено, не сменило тип, не стало обязательным", () => {
    expect(contractBreaks(release, current)).toEqual([]);
  });

  it("эталон — последний выпуск из CHANGELOG", () => {
    const latest = /^## \[(\d+\.\d+\.\d+)\]/m.exec(readFileSync("CHANGELOG.md", "utf8"))?.[1];
    expect(release.version, "обновите эталон: npm run contract:update").toBe(latest);
  });

  it("исключения в ALLOWED — с причиной", () => {
    for (const [k, why] of Object.entries(ALLOWED)) expect(why.length, k).toBeGreaterThan(20);
  });
});

describe("сама проверка ловит несовместимые изменения", () => {
  const tool = (inputSchema: Schema, outputSchema: Schema, annotations = { readOnlyHint: true }): Api => ({
    version: "0",
    prompts: [{ name: "p", arguments: ["database"] }],
    tools: [{ name: "read.x", inputSchema, outputSchema, annotations }],
  });
  const base = tool(
    {
      type: "object",
      properties: { a: { type: "string" }, kind: { type: "string", enum: ["x", "y"] } },
      required: [],
    },
    {
      type: "object",
      properties: {
        rows: {
          type: "array",
          items: { type: "object", properties: { sum: { type: "number" } }, required: ["sum"] },
        },
      },
      required: ["rows"],
    },
  );
  it.each([
    [
      "удалён параметр",
      tool(
        { type: "object", properties: { kind: base.tools[0]!.inputSchema.properties!["kind"]! } },
        base.tools[0]!.outputSchema!,
      ),
      /a: удалено/,
    ],
    [
      "новый обязательный параметр",
      tool(
        {
          ...base.tools[0]!.inputSchema,
          properties: { ...base.tools[0]!.inputSchema.properties, b: { type: "string" } },
          required: ["b"],
        },
        base.tools[0]!.outputSchema!,
      ),
      /b: стал обязательным/,
    ],
    [
      "убрано значение перечисления",
      tool(
        {
          ...base.tools[0]!.inputSchema,
          properties: { ...base.tools[0]!.inputSchema.properties, kind: { type: "string", enum: ["x"] } },
        },
        base.tools[0]!.outputSchema!,
      ),
      /убраны значения \["y"\]/,
    ],
    [
      "сменился тип поля ответа",
      tool(base.tools[0]!.inputSchema, {
        type: "object",
        properties: {
          rows: {
            type: "array",
            items: { type: "object", properties: { sum: { type: "string" } }, required: ["sum"] },
          },
        },
        required: ["rows"],
      }),
      /sum: тип/,
    ],
    [
      "поле ответа стало необязательным",
      tool(base.tools[0]!.inputSchema, {
        type: "object",
        properties: {
          rows: { type: "array", items: { type: "object", properties: { sum: { type: "number" } } } },
        },
        required: ["rows"],
      }),
      /sum: было всегда/,
    ],
    [
      "инструмент чтения стал пишущим",
      tool(base.tools[0]!.inputSchema, base.tools[0]!.outputSchema!, { readOnlyHint: false } as never),
      /стал пишущим/,
    ],
  ])("%s", (_, after, why) => {
    expect(contractBreaks(base, after).join("\n")).toMatch(why);
  });
  it("добавить необязательный параметр, поле ответа и значение перечисления в ответе можно", () => {
    const after = structuredClone(base);
    after.tools[0]!.inputSchema.properties!["extra"] = { type: "boolean" };
    after.tools[0]!.outputSchema!.properties!["note"] = { type: "string" };
    expect(contractBreaks(base, after)).toEqual([]);
  });
});
