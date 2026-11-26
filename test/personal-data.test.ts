import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * personal-data.json — карта полей ответов с персональными данными: по ней бот или шлюз маскирует ответы, не
 * угадывая. Карта устаревает молча (поле переименовали — маска перестала срабатывать), поэтому тест сверяет её с
 * manifest.json: каждый путь существует, а новое поле с «персональным» именем без разметки не проходит.
 */
type Schema = { type?: unknown; properties?: Record<string, Schema>; items?: Schema; anyOf?: Schema[] };
type Field = { tool: string; path: string; category: string; when: string };
const map = JSON.parse(readFileSync("personal-data.json", "utf8")) as {
  fields: Field[];
  when: Record<string, string>;
  categories: Record<string, string>;
  rawTools: { tools: string[] };
};
const tools = new Map(
  (
    JSON.parse(readFileSync("manifest.json", "utf8")) as {
      tools: Array<{ name: string; outputSchema?: Schema }>;
    }
  ).tools.map((t) => [t.name, t.outputSchema ?? {}]),
);

/** Листовые строковые поля схемы ответа: путь → схема. */
function leaves(s: Schema, path = "", out = new Map<string, Schema>()): Map<string, Schema> {
  for (const v of [s, ...(s.anyOf ?? [])]) {
    for (const [k, p] of Object.entries(v.properties ?? {})) {
      const at = path ? `${path}.${k}` : k;
      const types = [p.type, ...(p.anyOf ?? []).map((a) => a.type)].flat();
      if (types.includes("string")) out.set(at, p);
      leaves(p, at, out);
    }
    if (v.items) leaves(v.items, `${path}[]`, out);
  }
  return out;
}

/** Имена полей, которые почти наверняка несут персональные данные. */
const PERSONAL_NAME =
  /^(inn|fullName|shortName|responsible|director|accountant|manager|ogrn|accountNumber|phone|email|address|passport|snils|birthDate|counterparty|partner|comment|purpose)$/;
/** Совпало по имени, но не персональное — с причиной. */
const NOT_PERSONAL: Record<string, string> = {};

describe("personal-data.json", () => {
  it("каждый путь есть в ответе своего инструмента, категории и условия описаны", () => {
    const missing = map.fields.filter((f) => !tools.has(f.tool) || !leaves(tools.get(f.tool)!).has(f.path));
    expect(missing).toEqual([]);
    for (const f of map.fields) {
      expect(map.categories[f.category], f.category).toBeTruthy();
      expect(map.when[f.when], f.when).toBeTruthy();
    }
  });

  it("инструменты с реквизитами 1С как есть существуют", () => {
    expect(map.rawTools.tools.filter((t) => !tools.has(t))).toEqual([]);
  });

  it("каждое поле ответа с «персональным» именем размечено", () => {
    const marked = new Set(map.fields.map((f) => `${f.tool} ${f.path}`));
    const unmarked: string[] = [];
    for (const [tool, schema] of tools)
      for (const path of leaves(schema).keys()) {
        const key = `${tool} ${path}`;
        if (
          PERSONAL_NAME.test(path.split(".").pop()!.replace("[]", "")) &&
          !marked.has(key) &&
          !NOT_PERSONAL[key]
        )
          unmarked.push(key);
      }
    expect(unmarked, "добавьте поле в personal-data.json или в NOT_PERSONAL с причиной").toEqual([]);
  });

  it("проверка ловит новое неразмеченное поле", () => {
    const schema: Schema = { properties: { rows: { items: { properties: { phone: { type: "string" } } } } } };
    expect([...leaves(schema).keys()]).toEqual(["rows[].phone"]);
    expect(PERSONAL_NAME.test("phone")).toBe(true);
  });
});
