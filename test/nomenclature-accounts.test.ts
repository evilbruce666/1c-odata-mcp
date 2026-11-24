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
const row = (
  o: { org?: string; nom?: string; kind?: string; wh?: string; whType?: string },
  goods: number,
) => ({
  Организация_Key: o.org ?? E,
  Номенклатура_Key: o.nom ?? E,
  ВидНоменклатуры_Key: o.kind ?? E,
  ...(o.wh || o.whType ? { Склад_Key: o.wh ?? E, ТипСклада: o.whType ?? "" } : {}),
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
  /** ODATA_ANALYTICS_MAX_ROWS. */
  maxRows?: number;
  /** Родитель позиции или группы; есть — в схеме опубликованы Parent_Key, склад и тип склада. */
  parents?: Record<string, string>;
  warehouseTypes?: Record<string, string>;
}) {
  const log: string[] = [];
  const kindOf = (ref: string) => o.kinds[ref.toLowerCase()];
  return {
    log,
    conn: {
      behavior: { pageSize: 1000, analyticsMaxRows: o.maxRows ?? 200_000 },
      available: async () =>
        new Set([
          ...(o.register ? ["InformationRegister_СчетаУчетаНоменклатуры"] : []),
          "Catalog_Номенклатура",
          "Catalog_Склады",
        ]),
      getMetadata: async () => ({
        entities: new Map([
          [
            "Catalog_Номенклатура",
            {
              properties: [
                { name: "Ref_Key" },
                ...(o.kindField === false ? [] : [{ name: "ВидНоменклатуры_Key" }]),
                ...(o.parents ? [{ name: "Parent_Key" }] : []),
              ],
            },
          ],
          [
            "InformationRegister_СчетаУчетаНоменклатуры",
            { properties: o.parents ? [{ name: "Склад_Key" }, { name: "ТипСклада" }] : [] },
          ],
        ]),
      }),
      client: {
        getCollection: async (path: string) => {
          const p = decodeURIComponent(path);
          log.push(p);
          if (p.startsWith("InformationRegister_")) {
            // Как 1С: $top режет, отбор по Номенклатура_Key — без учёта регистра букв.
            const top = Number(/\$top=(\d+)/.exec(p)?.[1] ?? Infinity);
            const f = /\$filter=([^&]*)/.exec(p)?.[1];
            const wanted =
              f && new Set([...f.matchAll(/guid'([0-9a-fA-F-]{36})'/g)].map((m) => m[1]!.toLowerCase()));
            const rows = (o.register ?? []).filter(
              (r) => !wanted || wanted.has(String(r["Номенклатура_Key"]).toLowerCase()),
            );
            return { value: rows.slice(0, top) };
          }
          if (o.failChunks) throw new Error("сбой");
          const refs = [...p.matchAll(/guid'([0-9a-fA-F-]{36})'/g)].map((m) => m[1]!);
          return {
            value: refs
              .filter((r) => kindOf(r) || o.parents?.[r.toLowerCase()] || o.parents)
              .map((r) => ({
                Ref_Key: r.toLowerCase(),
                ВидНоменклатуры_Key: kindOf(r) ?? E,
                ...(o.parents ? { Parent_Key: o.parents[r.toLowerCase()] ?? E } : {}),
              })),
          };
        },
        getEntity: async (path: string) => {
          const p = decodeURIComponent(path);
          log.push(p);
          const ref = /guid'([^']+)'/.exec(p)![1]!;
          if (p.startsWith("Catalog_Склады")) {
            const t = o.warehouseTypes?.[ref];
            if (t === undefined) throw new Error("404");
            return { ТипСклада: t };
          }
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

  it("регистр больше 500 строк — подбор по всем записям: общие и записи позиций документа отбором", async () => {
    // Запись для позиции — 550-я из 600: раньше её отрезал предел 500, и позиция получала счёт общей записи.
    const big = [
      row({}, 1),
      ...Array.from({ length: 548 }, (_, i) => row({ nom: g("0f0f0f0f", i + 1) }, 100 + i)),
      row({ nom: item(1) }, 5),
      ...Array.from({ length: 50 }, (_, i) => row({ nom: g("0f0f0f0f", 1000 + i) }, 900 + i)),
    ];
    const c = conn({ register: big, kinds });
    const m = (await nomenclatureAccountsMany(c.conn, ORG, [item(1), item(2)]))!;
    expect(m.get(item(1))?.goods).toBe(acc(5));
    expect(m.get(item(2))?.goods).toBe(acc(1));
    // Начало регистра (501 строка), записи без номенклатуры, одна пачка позиций — регистр целиком не скачивается.
    const reads = c.log.filter((l) => l.startsWith("InformationRegister_"));
    expect(reads).toHaveLength(3);
    expect(reads[0]).toContain("$top=501");
    expect(reads.slice(1).every((l) => l.includes("$filter="))).toBe(true);
  });

  it("подбор по отбору — тот же, что по полному регистру", async () => {
    // Записи других позиций подбор не меняют: результат на большом регистре равен результату на малом без них.
    const filler = Array.from({ length: 600 }, (_, i) =>
      row({ nom: g("0f0f0f0f", i + 1), org: ORG }, 100 + i),
    );
    const refs = [item(1), item(2), item(4), item(5), item(2).toUpperCase(), `{${item(1)}}`, "не GUID"];
    const small = (await nomenclatureAccountsMany(conn({ register, kinds }).conn, ORG, refs))!;
    const m = (await nomenclatureAccountsMany(
      conn({ register: [...register, ...filler], kinds }).conn,
      ORG,
      refs,
    ))!;
    expect([...m]).toEqual([...small]);
    expect(m.get(item(1))?.goods).toBe(acc(5));
    expect(m.get(item(2))?.goods).toBe(acc(4));
  });

  it("подходящих записей больше предела выборки — понятный отказ, а не неточные счета", async () => {
    const big = Array.from({ length: 600 }, (_, i) => row({ kind: i % 2 ? K1 : K2 }, 100 + i));
    const c = conn({ register: big, kinds, maxRows: 10 });
    await expect(nomenclatureAccountsMany(c.conn, ORG, [item(1)])).rejects.toThrow(
      /Счета учёта номенклатуры.*ODATA_ANALYTICS_MAX_ROWS = 10/,
    );
  });

  it("ссылка позиции и организации в верхнем регистре или в скобках — своя запись, как у обычной формы", async () => {
    const forms = (r: string) => [r, r.toUpperCase(), `{${r}}`, ` ${r.toUpperCase()} `];
    for (const org of forms(ORG)) {
      const c = conn({ register, kinds });
      const m = (await nomenclatureAccountsMany(c.conn, org, [...forms(item(1)), ...forms(item(2))]))!;
      for (const f of forms(item(1))) expect(m.get(f)?.goods, `${org} ${f}`).toBe(acc(5)); // номенклатура
      for (const f of forms(item(2))) expect(m.get(f)?.goods, `${org} ${f}`).toBe(acc(4)); // вид + организация
    }
  });

  it("группа номенклатуры: ближайшая группа — выше вида и дальней группы, своя запись позиции — выше группы", async () => {
    const G1 = g("0e0e0e0e", 1); // группа позиций 1 и 5, внутри G2
    const G2 = g("0e0e0e0e", 2); // группа позиции 2
    const parents = { [item(1)]: G1, [item(5)]: G1, [G1]: G2, [item(2)]: G2 };
    const reg = [
      row({}, 1),
      row({ kind: K1 }, 2),
      row({ nom: G2 }, 20),
      row({ nom: G1 }, 21),
      row({ nom: item(5) }, 50),
    ];
    const c = conn({ register: reg, kinds, parents });
    const m = (await nomenclatureAccountsMany(c.conn, ORG, [item(1), item(2), item(3), item(5)]))!;
    expect(m.get(item(1))?.goods).toBe(acc(21)); // ближайшая группа, хотя есть запись вида
    expect(m.get(item(2))?.goods).toBe(acc(20)); // группа выше вида
    expect(m.get(item(3))?.goods).toBe(acc(2)); // без групп — вид
    expect(m.get(item(5))?.goods).toBe(acc(50)); // своя запись выше группы
    // Вид и родитель — одним запросом пачки; верхний уровень групп — ещё один.
    expect(c.log.filter((l) => l.startsWith("Catalog_Номенклатура"))).toHaveLength(2);
  });

  it("запись склада или типа склада — только для склада документа", async () => {
    const W1 = g("0f0f0f0f", 1);
    const W2 = g("0f0f0f0f", 2);
    const W3 = g("0f0f0f0f", 3);
    const reg = [row({}, 1), row({ wh: W1 }, 30), row({ whType: "РозничныйМагазин" }, 31)];
    const types = { [W1]: "ОптовыйСклад", [W2]: "РозничныйМагазин", [W3]: "ОптовыйСклад" };
    const pick = async (warehouseKey?: string) =>
      (await nomenclatureAccountsMany(
        conn({ register: reg, kinds, parents: {}, warehouseTypes: types }).conn,
        ORG,
        [item(1)],
        { warehouseKey },
      ))!.get(item(1))?.goods;
    expect(await pick(W1)).toBe(acc(30));
    expect(await pick(W1.toUpperCase())).toBe(acc(30));
    expect(await pick(W2)).toBe(acc(31));
    expect(await pick(W3)).toBe(acc(1));
    expect(await pick(undefined)).toBe(acc(1)); // без склада записи склада не подходят
  });

  it("регистр не опубликован — undefined (вызывающий берёт счета по кодам)", async () => {
    const c = conn({ register: null, kinds });
    expect(await nomenclatureAccountsMany(c.conn, ORG, [item(1)])).toBeUndefined();
    expect(c.log).toHaveLength(0);
  });
});
