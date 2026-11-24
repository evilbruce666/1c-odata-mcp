import { describe, expect, it } from "vitest";
import { nomenclatureAccountsMany } from "../src/odata/accounting.js";

/**
 * Счета учёта позиций из регистра «Счета учёта номенклатуры»: самая специфичная запись (номенклатура 4, вид 2,
 * организация 1; пустое измерение — «джокер»). Регистр читается один раз на вызов, вид позиций — пачками по 20 и
 * только если в регистре есть записи с видом (раньше — 2 запроса на каждую позицию документа).
 */

const E = "00000000-0000-0000-0000-000000000000";
const g = (p: string, i: number) => `${p}-0000-4000-8000-${String(i).padStart(12, "0")}`;
const ORG = g("0a0a0a0a", 1);
const OTHER_ORG = g("0a0a0a0a", 2);
const K1 = g("0b0b0b0b", 1);
const K2 = g("0b0b0b0b", 2);
const item = (i: number) => g("0c0c0c0c", i);
const acc = (i: number) => g("0d0d0d0d", i);
const row = (o: { org?: string; nom?: string; kind?: string }, goods: number) => ({
  Организация_Key: o.org ?? E,
  Номенклатура_Key: o.nom ?? E,
  ВидНоменклатуры_Key: o.kind ?? E,
  СчетУчета_Key: acc(goods),
  СчетУчетаНДСПоПриобретеннымЦенностям_Key: E,
  СчетУчетаНДСПоРеализации_Key: E,
  СчетДоходовОтРеализации_Key: E,
  СчетРасходовОтРеализации_Key: E,
});

function conn(o: {
  register: Record<string, unknown>[] | null;
  kinds: Record<string, string>;
  kindField?: boolean;
  failChunks?: boolean;
}) {
  const log: string[] = [];
  const kindOf = (ref: string) => o.kinds[ref.toLowerCase()];
  return {
    log,
    conn: {
      behavior: { pageSize: 1000 },
      available: async () =>
        new Set([
          ...(o.register ? ["InformationRegister_СчетаУчетаНоменклатуры"] : []),
          "Catalog_Номенклатура",
        ]),
      getMetadata: async () => ({
        entities: new Map([
          [
            "Catalog_Номенклатура",
            {
              properties: [
                { name: "Ref_Key" },
                ...(o.kindField === false ? [] : [{ name: "ВидНоменклатуры_Key" }]),
              ],
            },
          ],
        ]),
      }),
      client: {
        getCollection: async (path: string) => {
          const p = decodeURIComponent(path);
          log.push(p);
          if (p.startsWith("InformationRegister_")) return { value: o.register ?? [] };
          if (o.failChunks) throw new Error("сбой");
          const refs = [...p.matchAll(/guid'([0-9a-fA-F-]{36})'/g)].map((m) => m[1]!);
          return {
            value: refs
              .filter((r) => kindOf(r))
              .map((r) => ({ Ref_Key: r.toLowerCase(), ВидНоменклатуры_Key: kindOf(r) })),
          };
        },
        getEntity: async (path: string) => {
          const p = decodeURIComponent(path);
          log.push(p);
          const ref = /guid'([^']+)'/.exec(p)![1]!;
          if (ref === item(3)) throw new Error("сбой одной позиции");
          if (!kindOf(ref)) throw new Error("404");
          return { ВидНоменклатуры_Key: kindOf(ref) };
        },
      },
    } as never,
  };
}

const register = [
  row({}, 1), // джокер
  row({ kind: K1 }, 2),
  row({ org: ORG }, 3),
  row({ org: ORG, kind: K1 }, 4),
  row({ nom: item(1) }, 5),
  row({ org: OTHER_ORG, nom: item(2) }, 6), // чужая организация
  row({ kind: K2, org: OTHER_ORG }, 7), // чужая организация
];
const kinds = { [item(1)]: K1, [item(2)]: K1, [item(3)]: K1, [item(4)]: K2 };

describe("nomenclatureAccountsMany", () => {
  it("самая специфичная запись на каждую позицию; регистр — один запрос, вид — одной пачкой", async () => {
    const c = conn({ register, kinds });
    const refs = [item(1), item(2), item(4), item(5), item(2).toUpperCase(), `{${item(2)}}`];
    const m = (await nomenclatureAccountsMany(c.conn, ORG, refs))!;
    expect(m.get(item(1))?.goods).toBe(acc(5)); // номенклатура
    expect(m.get(item(2))?.goods).toBe(acc(4)); // вид + организация; запись другой организации не берётся
    expect(m.get(item(4))?.goods).toBe(acc(3)); // вид K2 только у чужой организации → организация
    expect(m.get(item(5))?.goods).toBe(acc(3)); // нет в справочнике → без вида
    // Ссылка в другом регистре или в скобках: вид находится (1С сравнивает GUID без регистра).
    expect(m.get(item(2).toUpperCase())?.goods).toBe(acc(4));
    expect(m.get(`{${item(2)}}`)?.goods).toBe(acc(4));
    expect(c.log.filter((l) => l.startsWith("InformationRegister_"))).toHaveLength(1);
    expect(c.log.filter((l) => l.startsWith("Catalog_"))).toHaveLength(1);
  });

  it("в регистре нет записей с видом — вид позиций не читается", async () => {
    const c = conn({ register: register.filter((r) => r.ВидНоменклатуры_Key === E), kinds });
    const m = (await nomenclatureAccountsMany(c.conn, ORG, [item(1), item(2)]))!;
    expect(m.get(item(1))?.goods).toBe(acc(5));
    expect(m.get(item(2))?.goods).toBe(acc(3));
    expect(c.log).toHaveLength(1);
  });

  it("45 позиций — вид тремя пачками по 20", async () => {
    const c = conn({ register, kinds });
    await nomenclatureAccountsMany(
      c.conn,
      ORG,
      Array.from({ length: 45 }, (_, i) => item(i + 1)),
    );
    expect(c.log.filter((l) => l.startsWith("Catalog_"))).toHaveLength(3);
  });

  it("пачка не прочиталась — по одной позиции; сбой одной не лишает вида остальные", async () => {
    const c = conn({ register, kinds, failChunks: true });
    const m = (await nomenclatureAccountsMany(c.conn, ORG, [item(2), item(3)]))!;
    expect(m.get(item(2))?.goods).toBe(acc(4));
    expect(m.get(item(3))?.goods).toBe(acc(3)); // её вид не прочитался → без вида
  });

  it("поле вида не опубликовано — подбор без вида, без запросов к справочнику", async () => {
    const c = conn({ register, kinds, kindField: false });
    const m = (await nomenclatureAccountsMany(c.conn, ORG, [item(2)]))!;
    expect(m.get(item(2))?.goods).toBe(acc(3));
    expect(c.log.filter((l) => l.startsWith("Catalog_"))).toHaveLength(0);
  });

  it("регистр не опубликован — undefined (вызывающий берёт счета по кодам)", async () => {
    const c = conn({ register: null, kinds });
    expect(await nomenclatureAccountsMany(c.conn, ORG, [item(1)])).toBeUndefined();
    expect(c.log).toHaveLength(0);
  });
});
